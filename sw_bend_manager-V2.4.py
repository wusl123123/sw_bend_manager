# 项目名称: SW折弯管理器
# 制作人: wsl
# 制作日期: 2026-06-14
# 说明: 读取并修改SolidWorks钣金件的折弯参数

import sys
import os
import time

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QComboBox, QPushButton, QGroupBox,
    QLabel, QTextEdit, QHeaderView, QMessageBox, QFileDialog
)
from PySide6.QtCore import QTimer, Qt, QThread, Signal
from PySide6.QtGui import QColor

import win32com.client
import pythoncom
import math


# 脚本实际文件名（解决SW内嵌Python的__file__不一致问题）
_SCRIPT_NAME = "sw_bend_manager-V2.2.py"


def _format_exc():
    """格式化异常堆栈，将文件名修正为实际脚本名"""
    import traceback
    raw = traceback.format_exc()
    # 修正可能出现的错误文件名
    raw = raw.replace('sw_bend_manager-V2.py', _SCRIPT_NAME)
    return raw


# ==================== 通用COM工具函数 ====================
# 以下函数兼容CDispatch和PyIDispatch，解决late binding下的调用问题

def _get_oleobj(obj):
    """获取底层OLE对象，兼容CDispatch和PyIDispatch"""
    return obj._oleobj_ if hasattr(obj, '_oleobj_') else obj


def _invoke(obj, name):
    """调用COM方法/属性，返回COM对象，兼容CDispatch和PyIDispatch"""
    try:
        oleobj = _get_oleobj(obj)
        dispid = oleobj.GetIDsOfNames(0, name)
    except Exception:
        return None

    # 策略1: PROPERTYGET + bResultWanted=True
    try:
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
        if result is not None:
            return result
    except Exception:
        pass

    # 策略2: METHOD + bResultWanted=True
    try:
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True)
        if result is not None:
            return result
    except Exception:
        pass

    return None


def _invoke_method(obj, name, *args):
    """调用COM方法（带参数），兼容CDispatch和PyIDispatch
    使用DISPATCH_METHOD标志，通过Invoke传递参数
    返回: 方法返回值或None"""
    try:
        oleobj = _get_oleobj(obj)
        dispid = oleobj.GetIDsOfNames(0, name)
        # 使用Invoke方法调用，bResultWanted=True获取返回值
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True, *args)
        return result
    except Exception:
        return None


def _com_attr(obj, name):
    """获取COM属性值（基本类型），兼容CDispatch和PyIDispatch"""
    try:
        oleobj = _get_oleobj(obj)
        dispid = oleobj.GetIDsOfNames(0, name)
        return oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYGET, True)
    except Exception:
        return None


def _com_set(obj, name, value):
    """设置COM属性值，兼容CDispatch和PyIDispatch
    使用DISPATCH_PROPERTYPUT标志，通过InvokeTypes精确控制参数类型
    回退策略: InvokeTypes(PROPERTYPUT) → Invoke(PROPERTYPUT) → Invoke(PROPERTYPUTREF)
    """
    try:
        oleobj = _get_oleobj(obj)
        dispid = oleobj.GetIDsOfNames(0, name)
        # 根据值类型选择VT常量
        if isinstance(value, float):
            arg_descs = ((pythoncom.VT_R8, 0),)
        elif isinstance(value, int):
            arg_descs = ((pythoncom.VT_I4, 0),)
        elif isinstance(value, bool):
            arg_descs = ((pythoncom.VT_BOOL, 0),)
        else:
            arg_descs = ((pythoncom.VT_VARIANT, 0),)
        # 策略1: 使用InvokeTypes精确控制参数类型
        oleobj.InvokeTypes(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, (pythoncom.VT_EMPTY, 0), arg_descs, value)
        return True
    except Exception:
        pass

    # 回退1: 普通Invoke + PROPERTYPUT
    try:
        oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
        return True
    except Exception:
        pass

    # 回退2: PROPERTYPUTREF（某些COM对象需要）
    try:
        oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUTREF, 0, value)
        return True
    except Exception:
        return False


def _com_method(obj, name, *args):
    """调用COM方法并传参，兼容CDispatch和PyIDispatch
    策略: InvokeTypes(METHOD) → 普通Invoke(METHOD) → 仅用VT_DISPATCH的Invoke
    返回: (bool, result) - (是否成功, 返回值)
    """
    oleobj = None
    dispid = None
    try:
        oleobj = _get_oleobj(obj)
        dispid = oleobj.GetIDsOfNames(0, name)
    except Exception:
        return False, None

    # 构建参数描述符
    arg_descs = []
    for arg in args:
        if isinstance(arg, float):
            arg_descs.append((pythoncom.VT_R8, 0))
        elif isinstance(arg, int):
            arg_descs.append((pythoncom.VT_I4, 0))
        elif isinstance(arg, bool):
            arg_descs.append((pythoncom.VT_BOOL, 0))
        elif arg is None:
            arg_descs.append((pythoncom.VT_VARIANT, 0))
        elif hasattr(arg, '_oleobj_') or hasattr(arg, 'InvokeTypes'):
            arg_descs.append((pythoncom.VT_DISPATCH, 0))
        else:
            arg_descs.append((pythoncom.VT_VARIANT, 0))

    # 策略1: InvokeTypes（精确类型控制）
    if arg_descs:
        try:
            result = oleobj.InvokeTypes(dispid, 0, pythoncom.DISPATCH_METHOD,
                                       (pythoncom.VT_VARIANT, 0),
                                       tuple(arg_descs), *args)
            return True, result
        except Exception:
            pass
    else:
        try:
            result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True)
            return True, result
        except Exception:
            pass

    # 策略2: 普通Invoke（有参数）
    if args:
        try:
            result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True, *args)
            return True, result
        except Exception:
            pass

    # 策略3: 普通Invoke（无参数）
    try:
        result = oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, True)
        return True, result
    except Exception:
        pass

    # 策略4: 无参数，不要求返回值
    try:
        oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_METHOD, False)
        return True, None
    except Exception:
        return False, None


def _safe_get(obj, name):
    """安全获取COM属性/方法返回值——对返回基本类型的成员有效"""
    try:
        attr = getattr(obj, name)
        if callable(attr):
            return attr()
        return attr
    except Exception:
        return None


# ==================== 常量定义 ====================
# 钣金特征类型列表（用于查找折弯子特征，参考VBA宏）
# SheetMetal是顶层钣金特征，必须包含，否则修改线程遍历不到其子特征
SHEET_METAL_FEATURE_TYPES = [
    "SheetMetal", "EdgeFlange", "FlattenBends", "SMBaseFlange",
    "SMMiteredFlange", "Hem", "Jog", "SM3dBend", "LoftedBend",
    "SolidToSheetMetal", "BaseFlange"
]

# 折弯子特征类型
BEND_FEATURE_TYPES = ["OneBend", "SketchBend"]

# 钣金判断特征类型（参考VBA IsSheetMetalByFeature函数）
SHEET_METAL_IDENTIFY_TYPES = ["BaseFlange", "SMBaseFlange", "SheetMetal"]

# 折弯系数类型映射（统一使用中文显示）
BEND_ALLOWANCE_TYPE_MAP = {
    2: "K因子",        # K因子
    4: "折弯扣除"       # 折弯扣除
}

BEND_ALLOWANCE_TYPE_REVERSE_MAP = {
    "K因子": 2,
    "折弯扣除": 4
}


# ==================== 业务逻辑处理器 ====================
class BendDataProcessor:
    """钣金折弯数据处理器"""

    def __init__(self):
        self.sw_app = None
        self.sw_model = None
        self.sheet_metal_thickness = 0.0
        self.debug_info = []  # 调试信息

    def connect_to_sw(self):
        """连接SolidWorks实例（在工作线程中调用，需先初始化COM）"""
        try:
            # 初始化COM（必须在当前线程中调用）
            pythoncom.CoInitialize()
            self.sw_app = win32com.client.Dispatch("SldWorks.Application")
            # 确保SolidWorks可见，以便用户知道已连接
            try:
                self.sw_app.Visible = True
            except:
                pass
            return True, "已成功连接到SolidWorks"
        except Exception as e:
            self.debug_info.append(f"connect_to_sw详细错误: {_format_exc()}")
            return False, f"连接SolidWorks失败: {str(e)}"

    def refresh_model(self):
        """刷新当前激活的模型引用（需要COM初始化）"""
        try:
            pythoncom.CoInitialize()
            self.sw_model = self.get_active_model()
            return self.sw_model
        except Exception as e:
            self.debug_info.append(f"刷新模型引用出错: {str(e)}")
            return None

    def get_active_model(self):
        """获取当前激活的模型"""
        if not self.sw_app:
            return None
        try:
            return self.sw_app.ActiveDoc
        except:
            return None

    def is_sheet_metal_part(self, model):
        """
        判断零件是否为钣金件
        参考VBA宏 IsSheetMetalByFeature 函数
        使用 GetTypeName 方法检测
        返回: (bool, list) - (是否钣金件, 检测到的特征类型列表)
        """
        if not model:
            return False, ["模型为空"]

        detected_types = []
        self.debug_info = []

        try:
            # 检查文档类型 (1=Part, 2=Assembly, 3=Drawing)
            doc_type = model.GetType()
            self.debug_info.append(f"文档类型: {doc_type}")

            if doc_type != 1:
                return False, [f"文档类型={doc_type}, 非零件"]

            # 遍历特征，使用GetTypeName检测（与VBA一致）
            feat_count = 0
            feat = model.FirstFeature
            while feat:
                feat_count += 1
                try:
                    feat_type = feat.GetTypeName()
                except:
                    try:
                        feat_type = feat.GetTypeName2()
                    except:
                        feat_type = ""
                detected_types.append(feat_type)

                if feat_type in SHEET_METAL_IDENTIFY_TYPES:
                    self.debug_info.append(f"找到钣金特征: {feat_type}")
                    return True, detected_types

                try:
                    feat = feat.GetNextFeature()
                except:
                    try:
                        feat = feat.GetNextFeature
                    except:
                        break

            self.debug_info.append(f"共遍历 {feat_count} 个顶层特征")

            # 回退：通过Body判断是否有钣金实体
            try:
                bodies = model.GetBodies2(0, False)
                if bodies:
                    for body in bodies:
                        try:
                            if body.IsSheetMetal:
                                self.debug_info.append("通过Body.IsSheetMetal检测到钣金实体")
                                return True, detected_types
                        except:
                            pass
            except:
                pass

            self.debug_info.append("未找到钣金特征")
            return False, detected_types

        except Exception as e:
            self.debug_info.append(f"检测出错: {str(e)}")
            return False, detected_types

    def get_sheet_metal_thickness(self, model):
        """获取钣金厚度"""
        if not model:
            return 0.0

        try:
            feat = model.FirstFeature
            while feat:
                try:
                    feat_type = feat.GetTypeName2()
                except:
                    try:
                        feat_type = feat.GetTypeName()
                    except:
                        feat_type = ""
                if feat_type == "SMBaseFlange":
                    sw_smfd = feat.GetDefinition()
                    return sw_smfd.Thickness
                try:
                    feat = feat.GetNextFeature()
                except:
                    try:
                        feat = feat.GetNextFeature
                    except:
                        break
            return 0.0
        except:
            return 0.0

    def collect_bend_data(self, model, progress_callback=None):
        """
        收集钣金折弯数据
        参考VBA宏 显示并修改折弯扣除.txt 的逻辑
        返回: list[dict] - 折弯数据列表
        """
        bend_data_list = []

        if not model:
            return bend_data_list

        try:
            self.sheet_metal_thickness = self.get_sheet_metal_thickness(model)
            feat = model.FirstFeature
            index = 0

            while feat:
                try:
                    feat_type = feat.GetTypeName2()
                except:
                    try:
                        feat_type = feat.GetTypeName()
                    except:
                        feat_type = ""

                if feat_type in SHEET_METAL_FEATURE_TYPES:
                    try:
                        sub_feat = feat.GetFirstSubFeature()
                    except:
                        sub_feat = feat.GetFirstSubFeature

                    while sub_feat:
                        try:
                            sub_feat_type = sub_feat.GetTypeName2()
                        except:
                            try:
                                sub_feat_type = sub_feat.GetTypeName()
                            except:
                                sub_feat_type = ""

                        if sub_feat_type in BEND_FEATURE_TYPES:
                            index += 1
                            try:
                                obfd = sub_feat.GetDefinition()
                                cba = obfd.GetCustomBendAllowance()

                                # 获取折弯系数类型
                                allowance_type = cba.Type
                                allowance_type_str = BEND_ALLOWANCE_TYPE_MAP.get(allowance_type, "K因子")

                                # 获取折弯系数数值
                                if allowance_type == 2:  # K因子
                                    allowance_value = cba.KFactor
                                elif allowance_type == 4:  # 折弯扣除
                                    allowance_value = cba.BendDeduction * 1000  # 转换为mm
                                else:
                                    allowance_value = 0

                                # 保存原始值用于还原
                                original_data = {
                                    "index": index,
                                    "feature_name": sub_feat.Name,
                                    "bend_angle": obfd.BendAngle * 180.0 / 3.14159265358979,  # 精确转换为度
                                    "thickness": self.sheet_metal_thickness * 1000,  # 转换为mm
                                    "bend_radius": obfd.BendRadius * 1000,  # 转换为mm
                                    "allowance_type": allowance_type_str,
                                    "allowance_value": allowance_value,
                                    "feature_ptr": sub_feat,
                                    "obfd_ptr": obfd,
                                    "cba_ptr": cba,
                                    "original_angle": obfd.BendAngle,
                                    "original_radius": obfd.BendRadius,
                                    "original_allowance_type": allowance_type,
                                    "original_allowance_value": allowance_value if allowance_type == 2 else cba.BendDeduction,
                                }

                                bend_data_list.append(original_data)
                            except Exception as e:
                                self.debug_info.append(f"获取折弯数据出错(第{index}个): {str(e)}")

                        try:
                            sub_feat = sub_feat.GetNextSubFeature()
                        except:
                            try:
                                sub_feat = sub_feat.GetNextSubFeature
                            except:
                                try:
                                    sub_feat = sub_feat.GetNextFeature()
                                except:
                                    break

                try:
                    feat = feat.GetNextFeature()
                except:
                    try:
                        feat = feat.GetNextFeature
                    except:
                        break

        except Exception as e:
            self.debug_info.append(f"收集折弯数据时出错: {str(e)}")

        return bend_data_list


# ==================== 修改折弯参数工作线程 ====================
class ModifyBendWorker(QThread):
    """修改折弯参数的工作线程——COM修改操作必须在独立线程中执行
    严格对齐VBA修改逻辑:
      1. 遍历特征树, 找到钣金子特征(OneBend/SketchBend)
      2. 通过名称匹配定位目标特征
      3. cba.SetType → cba.SetKFactor/SetBendDeduction → feat.ModifyDefinition(obfd, model, None) → model.EditRebuild3
      4. 重新读取验证修改是否真正生效"""
    progress_updated = Signal(int, str)          # (row_index, message)
    modify_completed = Signal(bool, int, str)     # (success, success_count, message)
    single_modify_done = Signal(bool, int, str)   # (success, row, message)

    def __init__(self, task_type, bend_data_list=None, row=-1, new_type="KFactor", new_value=0.0):
        super().__init__()
        self.task_type = task_type          # "single" 或 "batch"
        self.bend_data_list = bend_data_list
        self.row = row                      # 单行修改的行号
        self.new_type = new_type            # 新的折弯系数类型字符串
        self.new_value = new_value          # 新的折弯系数数值

    def run(self):
        """在工作线程中执行COM修改操作——严格对齐VBA修改流程"""
        import pythoncom as _pc
        import win32com.client as _w32
        import traceback as _tb

        try:
            _pc.CoInitialize()

            # 连接SolidWorks（与VBA Set swApp = Application.SldWorks 对齐）
            # 优先使用GetObject获取已运行实例（数据抓取线程已验证此方式可行）
            sw_app = None
            try:
                sw_app = _w32.GetObject(None, "SldWorks.Application")
            except Exception:
                try:
                    sw_app = _w32.Dispatch("SldWorks.Application")
                except Exception:
                    try:
                        sw_app = _w32.gencache.EnsureDispatch("SldWorks.Application")
                    except Exception:
                        self._emit_failure(0, "无法连接SolidWorks")
                        return

            # 获取活动文档（与VBA Set swmodel = swApp.ActiveDoc 对齐）
            model = sw_app.ActiveDoc
            if model is None:
                self._emit_failure(0, "没有激活的SolidWorks文档")
                return

            # 在当前线程重新构建特征名称→(feat, obfd, cba) 映射
            # 使用与VBA完全一致的遍历逻辑
            feature_map = self._build_feature_map_vba_style(model)

            if not feature_map:
                self._emit_failure(0, "无法遍历特征树，请确保当前打开的是钣金零件")
                return

            allowance_type_enum = BEND_ALLOWANCE_TYPE_REVERSE_MAP.get(self.new_type, 2)

            if self.task_type == "single":
                self._modify_single_vba_style(model, feature_map, allowance_type_enum)
            elif self.task_type == "batch":
                self._modify_batch_vba_style(model, feature_map)

        except Exception as e:
            self._emit_failure(0, f"线程执行出错: {str(e)}\n{_format_exc()}")

    def _emit_failure(self, count, message):
        """根据任务类型发送失败信号，附加调试信息"""
        debug_info = getattr(self, '_debug_info', [])
        if debug_info:
            message += "\n[修改线程调试信息]\n" + "\n".join(debug_info)
        if self.task_type == "single":
            self.single_modify_done.emit(False, self.row, message)
        else:
            self.modify_completed.emit(False, count, message)

    def _build_feature_map_vba_style(self, model):
        """严格对齐VBA遍历逻辑构建特征映射
        使用底层_invoke/_com_attr调用，兼容CDispatch和PyIDispatch
        VBA逻辑:
          While Not swfeat Is Nothing
            If swfeat.GetTypeName2 是钣金特征类型 Then
              Set swsubfeat = swfeat.GetFirstSubFeature
              While Not swsubfeat Is Nothing
                If swsubfeat.GetTypeName2 = "OneBend" Or "SketchBend" Then
                  收集 swsubfeat, swOBFD, swCBA
                End If
                Set swsubfeat = swsubfeat.GetNextSubFeature
              Wend
            End If
            Set swfeat = swfeat.GetNextFeature
          Wend
        返回: {特征名称: (feat, obfd, cba)}
        """
        import traceback as _tb
        feature_map = {}
        debug_lines = []

        # ---- 底层COM辅助函数（线程内使用，兼容CDispatch）----
        def _ft(f):
            """获取特征类型名称（底层_invoke调用）"""
            result = _invoke(f, "GetTypeName2")
            if result is not None and isinstance(result, str):
                return result
            result = _invoke(f, "GetTypeName")
            if result is not None and isinstance(result, str):
                return result
            return ""

        def _fn(f):
            """获取特征名称（底层_com_attr调用）"""
            result = _com_attr(f, "Name")
            if result is not None and isinstance(result, str):
                return result
            return ""

        def _nf(f):
            """获取下一个特征（底层_invoke调用）"""
            result = _invoke(f, "GetNextFeature")
            return result if result is not None else None

        def _fsf(f):
            """获取第一个子特征（底层_invoke调用）"""
            result = _invoke(f, "GetFirstSubFeature")
            return result if result is not None else None

        def _nsf(f):
            """获取下一个子特征（底层_invoke调用）"""
            result = _invoke(f, "GetNextSubFeature")
            return result if result is not None else None

        try:
            # 获取第一个顶层特征（VBA: Set swfeat = swmodel.FirstFeature）
            feat = _invoke(model, "FirstFeature")
            feat_idx = 0
            while feat is not None:
                feat_idx += 1
                feat_type = _ft(feat)

                # 最多遍历100个特征防止死循环
                if feat_idx > 200:
                    debug_lines.append("警告: 特征遍历超过200个，强制终止")
                    break

                # VBA: If swfeat.GetTypeName2 是钣金特征类型 Then
                if feat_type in SHEET_METAL_FEATURE_TYPES:
                    debug_lines.append(f"修改线程-匹配钣金特征[{feat_idx}]: {feat_type}")
                    # VBA: Set swsubfeat = swfeat.GetFirstSubFeature
                    sub_feat = _fsf(feat)

                    sub_idx = 0
                    # VBA: While Not swsubfeat Is Nothing
                    while sub_feat is not None:
                        sub_idx += 1
                        if sub_idx > 100:
                            debug_lines.append(f"  警告: 子特征遍历超过100个，跳出")
                            break

                        sub_type = _ft(sub_feat)

                        # VBA: If swsubfeat.GetTypeName2 = "OneBend" Or "SketchBend" Then
                        if sub_type in BEND_FEATURE_TYPES:
                            try:
                                feat_name = _fn(sub_feat)
                                # VBA: Set swOBFD = swsubfeat.GetDefinition
                                obfd = _invoke(sub_feat, "GetDefinition")
                                # VBA: Set swCBA = swOBFD.GetCustomBendAllowance
                                if obfd is not None:
                                    cba = _invoke(obfd, "GetCustomBendAllowance")
                                else:
                                    cba = None
                                if feat_name and obfd and cba:
                                    feature_map[feat_name] = (sub_feat, obfd, cba)
                                    debug_lines.append(f"  收集折弯: {feat_name} ({sub_type})")
                                else:
                                    debug_lines.append(f"  跳过(缺少数据): name={feat_name}, obfd={obfd is not None}, cba={cba is not None}")
                            except Exception as e:
                                debug_lines.append(f"  收集出错: {str(e)}")
                        else:
                            debug_lines.append(f"  子特征[{sub_idx}]: {sub_type} (非折弯类型)")

                        # VBA: Set swsubfeat = swsubfeat.GetNextSubFeature
                        sub_feat = _nsf(sub_feat)

                # VBA: Set swfeat = swfeat.GetNextFeature
                feat = _nf(feat)

            debug_lines.append(f"修改线程-特征映射总数: {len(feature_map)}")

        except Exception as e:
            debug_lines.append(f"遍历异常: {str(e)}\n{_format_exc()}")

        # 存储调试信息供外部读取
        self._debug_info = debug_lines
        return feature_map

    def _modify_single_vba_style(self, model, feature_map, allowance_type_enum):
        """单行修改——严格对齐VBA修改流程:
          cba.SetType(type)
          cba.SetKFactor(value) 或 cba.SetBendDeduction(value)
          feat.ModifyDefinition(obfd, model, Nothing)
          model.EditRebuild
        所有COM调用使用底层_invoke/_com_set，兼容CDispatch
        """
        if self.row < 0 or not self.bend_data_list or self.row >= len(self.bend_data_list):
            self.single_modify_done.emit(False, self.row, "行号无效")
            return

        bend_data = self.bend_data_list[self.row]
        feature_name = bend_data.get("feature_name", "未知")

        if feature_name not in feature_map:
            self.single_modify_done.emit(False, self.row, f"修改失败: 未找到特征 '{feature_name}'，请刷新后重试")
            return

        try:
            bend_feat, obfd, cba = feature_map[feature_name]

            # ===== 步骤1: 设置折弯系数类型 (VBA: cba.SetType) =====
            # 使用底层_com_set回退到SetType方法调用
            set_type_ok = False
            try:
                cba.SetType(allowance_type_enum)
                set_type_ok = True
            except Exception:
                try:
                    _com_set(cba, "Type", allowance_type_enum)
                    set_type_ok = True
                except Exception:
                    pass

            if not set_type_ok:
                raise Exception(f"无法设置折弯系数类型 (期望={allowance_type_enum})")

            # ===== 步骤2: 设置折弯系数数值 =====
            if allowance_type_enum == 2:  # K因子 (VBA: cba.SetKFactor)
                try:
                    cba.SetKFactor(self.new_value)
                except Exception:
                    try:
                        _com_set(cba, "KFactor", self.new_value)
                    except Exception:
                        raise Exception(f"无法设置K因子值: {self.new_value}")
            elif allowance_type_enum == 4:  # 折弯扣除 (VBA: cba.SetBendDeduction)
                expected_m = self.new_value / 1000.0  # mm转m（SW内部单位）
                try:
                    cba.SetBendDeduction(expected_m)
                except Exception:
                    try:
                        _com_set(cba, "BendDeduction", expected_m)
                    except Exception:
                        raise Exception(f"无法设置折弯扣除值: {self.new_value}mm")

            # ===== 步骤2.5: 将修改后的cba显式写回obfd（兼容CDispatch副本问题）=====
            # 在CDispatch模式下，GetCustomBendAllowance可能返回副本，
            # 需要显式SetCustomBendAllowance将修改写回obfd，确保ModifyDefinition生效
            _invoke_method(obfd, "SetCustomBendAllowance", cba)

            # ===== 步骤3: 应用修改 (VBA: feat.ModifyDefinition obfd, swmodel, Nothing) =====
            mod_ok, mod_msg = self._call_modify_definition_vba_style(bend_feat, obfd, model)
            if not mod_ok:
                raise Exception(f"ModifyDefinition失败: {mod_msg}")

            # ===== 步骤4: 重建模型 (VBA: swmodel.EditRebuild) =====
            try:
                _invoke(model, "EditRebuild3")
            except Exception:
                try:
                    _invoke(model, "EditRebuild")
                except Exception:
                    pass

            # ===== 步骤5: 重新读取验证（使用底层_invoke调用） =====
            verify_obfd = _invoke(bend_feat, "GetDefinition")
            if verify_obfd is None:
                raise Exception("验证失败: 无法获取特征定义")
            verify_cba = _invoke(verify_obfd, "GetCustomBendAllowance")
            if verify_cba is None:
                raise Exception("验证失败: 无法获取折弯系数对象")
            final_type = _com_attr(verify_cba, "Type")
            if allowance_type_enum == 2:
                final_val = _com_attr(verify_cba, "KFactor")
            elif allowance_type_enum == 4:
                final_val = _com_attr(verify_cba, "BendDeduction")
                if final_val is not None:
                    final_val = final_val * 1000.0  # m转mm
            else:
                final_val = None

            if final_type != allowance_type_enum:
                raise Exception(f"修改后验证失败: 类型未改变 (期望={allowance_type_enum}, 实际={final_type})")
            if final_val is not None and abs(final_val - self.new_value) > 0.001:
                raise Exception(f"修改后验证失败: 数值不匹配 (期望={self.new_value}, 实际={final_val})")

            self.single_modify_done.emit(True, self.row,
                f"已修改: {feature_name} -> 类型={self.new_type}, 数值={self.new_value} (验证值={final_val}, 策略={mod_msg})")

        except Exception as e:
            self.single_modify_done.emit(False, self.row,
                f"修改失败({feature_name}): {str(e)}\n{_format_exc()}")

    def _modify_batch_vba_style(self, model, feature_map):
        """批量修改——按每行各自选择的系数类型和数值分别回填（等价分别点击每行修改按钮）
        所有COM调用使用底层_invoke/_com_set，兼容CDispatch"""
        if not self.bend_data_list:
            self.modify_completed.emit(False, 0, "没有可修改的数据")
            return

        success_count = 0
        fail_count = 0
        detail_msgs = []  # 收集详细结果信息

        for row in range(len(self.bend_data_list)):
            bend_data = self.bend_data_list[row]
            feature_name = bend_data.get("feature_name", "未知")

            # 从bend_data中读取该行的折弯系数类型和数值（由调用方在UI层已填入）
            row_type_str = bend_data.get("ui_allowance_type", "")
            row_value = bend_data.get("ui_allowance_value", None)

            if not row_type_str or row_value is None:
                fail_count += 1
                self.progress_updated.emit(row, f"第 {row + 1} 行跳过: 缺少折弯系数类型或数值")
                continue

            # 将类型字符串转换为枚举值
            row_type_enum = BEND_ALLOWANCE_TYPE_REVERSE_MAP.get(row_type_str, 2)

            if feature_name not in feature_map:
                fail_count += 1
                self.progress_updated.emit(row, f"第 {row + 1} 行跳过: 未找到特征 '{feature_name}'")
                continue

            try:
                bend_feat, obfd, cba = feature_map[feature_name]

                # 步骤1: 设置类型（使用底层_com_set回退）
                try:
                    cba.SetType(row_type_enum)
                except Exception:
                    try:
                        _com_set(cba, "Type", row_type_enum)
                    except Exception:
                        raise Exception(f"无法设置类型({row_type_str})")

                # 步骤2: 设置数值（使用底层_com_set回退）
                if row_type_enum == 2:  # K因子
                    try:
                        cba.SetKFactor(row_value)
                    except Exception:
                        try:
                            _com_set(cba, "KFactor", row_value)
                        except Exception:
                            raise Exception(f"无法设置K因子值({row_value})")
                elif row_type_enum == 4:  # 折弯扣除
                    expected_m = row_value / 1000.0
                    try:
                        cba.SetBendDeduction(expected_m)
                    except Exception:
                        try:
                            _com_set(cba, "BendDeduction", expected_m)
                        except Exception:
                            raise Exception(f"无法设置折弯扣除值({row_value}mm)")

                # 步骤2.5: 将修改后的cba显式写回obfd（兼容CDispatch副本问题）
                _invoke_method(obfd, "SetCustomBendAllowance", cba)

                # 步骤3: 应用修改
                mod_ok, mod_msg = self._call_modify_definition_vba_style(bend_feat, obfd, model)
                if not mod_ok:
                    raise Exception(f"ModifyDefinition失败: {mod_msg}")

                # 步骤4: 单特征修改后立即重建模型
                try:
                    _invoke(model, "EditRebuild3")
                except Exception:
                    try:
                        _invoke(model, "EditRebuild")
                    except Exception:
                        pass

                # 步骤5: 验证修改结果
                verify_obfd = _invoke(bend_feat, "GetDefinition")
                if verify_obfd is None:
                    raise Exception("验证失败: 无法获取特征定义")
                verify_cba = _invoke(verify_obfd, "GetCustomBendAllowance")
                if verify_cba is None:
                    raise Exception("验证失败: 无法获取折弯系数对象")
                final_type = _com_attr(verify_cba, "Type")
                if row_type_enum == 2:
                    final_val = _com_attr(verify_cba, "KFactor")
                elif row_type_enum == 4:
                    final_val = _com_attr(verify_cba, "BendDeduction")
                    if final_val is not None:
                        final_val = final_val * 1000.0
                else:
                    final_val = None

                if final_type != row_type_enum:
                    raise Exception(f"验证: 类型不匹配(期望={row_type_str}, 实际类型枚举={final_type})")
                if final_val is not None and abs(final_val - row_value) > 0.001:
                    raise Exception(f"验证: 数值不匹配(期望={row_value}, 实际={final_val})")

                success_count += 1
                detail_msg = f"第 {row + 1} 行修改成功: {feature_name} -> 类型={row_type_str}, 数值={row_value} (验证={final_val})"
                detail_msgs.append(detail_msg)
                self.progress_updated.emit(row, detail_msg)

            except Exception as e:
                fail_count += 1
                err_msg = f"第 {row + 1} 行修改失败({feature_name}): {str(e)}"
                detail_msgs.append(err_msg)
                self.progress_updated.emit(row, err_msg)

        # 组装最终结果消息
        summary = f"统一修改完成: 成功 {success_count} 条, 失败 {fail_count} 条"
        if detail_msgs:
            summary += "\n" + "\n".join(detail_msgs)

        self.modify_completed.emit(success_count > 0, success_count, summary)

    def _call_modify_definition_vba_style(self, bend_feat, obfd, model):
        """调用 ModifyDefinition — 严格对齐VBA: feat.ModifyDefinition(obfd, model, Nothing)
        SW API签名: IFeature.ModifyDefinition(IDispatch* Definition, IModelDoc2* Model, IComponent2* Component)
        VBA传Nothing表示空组件引用

        策略优先级:
          1. 三参数(feat, obfd, model) — 使用底层InvokeTypes精确类型
          2. 三参数(feat, obfd, model) — 使用普通Invoke
          3. 三参数传递None作为Component
          4. 双参数(obfd, model)
          5. 先EditRebuild3再重试

        返回: (bool, str) - (是否成功, 策略描述)
        """
        import pythoncom as _pc

        # 策略1: InvokeTypes三参数 (obfd=VT_DISPATCH, model=VT_DISPATCH, None=VT_DISPATCH)
        try:
            oleobj = _get_oleobj(bend_feat)
            dispid = oleobj.GetIDsOfNames(0, "ModifyDefinition")
            # 第三个参数传None(空指针)模拟VBA的Nothing
            result = oleobj.InvokeTypes(
                dispid, 0, _pc.DISPATCH_METHOD,
                (_pc.VT_BOOL, 0),  # 返回类型: BOOL
                ((_pc.VT_DISPATCH, 0), (_pc.VT_DISPATCH, 0), (_pc.VT_DISPATCH, 0)),
                obfd, model, None
            )
            if result is True:
                return True, "InvokeTypes三参(None)"
            elif result is False:
                # 返回值是False但调用本身成功，继续尝试其他策略
                pass
        except Exception:
            pass

        # 策略2: 普通Invoke三参数
        try:
            oleobj = _get_oleobj(bend_feat)
            dispid = oleobj.GetIDsOfNames(0, "ModifyDefinition")
            result = oleobj.Invoke(dispid, 0, _pc.DISPATCH_METHOD, True, obfd, model, None)
            if result is True:
                return True, "Invoke三参(None)"
        except Exception:
            pass

        # 策略3: InvokeTypes双参数 (不含Component，某些SW版本允许)
        try:
            oleobj = _get_oleobj(bend_feat)
            dispid = oleobj.GetIDsOfNames(0, "ModifyDefinition")
            result = oleobj.InvokeTypes(
                dispid, 0, _pc.DISPATCH_METHOD,
                (_pc.VT_BOOL, 0),
                ((_pc.VT_DISPATCH, 0), (_pc.VT_DISPATCH, 0)),
                obfd, model
            )
            if result is True:
                return True, "InvokeTypes双参"
        except Exception:
            pass

        # 策略4: 普通Invoke双参数
        try:
            oleobj = _get_oleobj(bend_feat)
            dispid = oleobj.GetIDsOfNames(0, "ModifyDefinition")
            result = oleobj.Invoke(dispid, 0, _pc.DISPATCH_METHOD, True, obfd, model)
            if result is True:
                return True, "Invoke双参"
        except Exception:
            pass

        # 策略5: 先重建模型再重试双参数
        try:
            model.EditRebuild3()
            oleobj = _get_oleobj(bend_feat)
            dispid = oleobj.GetIDsOfNames(0, "ModifyDefinition")
            result = oleobj.InvokeTypes(
                dispid, 0, _pc.DISPATCH_METHOD,
                (_pc.VT_BOOL, 0),
                ((_pc.VT_DISPATCH, 0), (_pc.VT_DISPATCH, 0)),
                obfd, model
            )
            if result is True:
                return True, "Rebuild后InvokeTypes双参"
        except Exception:
            pass

        return False, "所有ModifyDefinition策略均失败"


# ==================== 数据收集工作线程 ====================
class WorkerThread(QThread):
    """异步工作线程——所有COM操作在run()内直接完成，不依赖任何外部对象"""
    progress_updated = Signal(int, str)
    task_completed = Signal(bool, str, list, list, bool)  # (success, message, data, debug_info, is_sheet_metal)

    def __init__(self, task_type, data=None):
        super().__init__()
        self.task_type = task_type
        self.data = data

    def run(self):
        """线程执行入口——使用模块级COM工具函数兼容late binding"""
        import pythoncom as _pc
        import win32com.client as _w32
        import traceback as _tb

        debug_info = []
        try:
            # 步骤1: 在当前线程初始化COM（STA模式）
            _pc.CoInitialize()

            # 步骤2: 连接SolidWorks
            sw_app = None
            try:
                sw_app = _w32.gencache.EnsureDispatch("SldWorks.Application")
                debug_info.append("连接方式: 早期绑定")
            except Exception:
                try:
                    sw_app = _w32.GetObject(None, "SldWorks.Application")
                    debug_info.append("连接方式: GetObject")
                except Exception:
                    try:
                        sw_app = _w32.Dispatch("SldWorks.Application")
                        debug_info.append("连接方式: Dispatch")
                    except Exception:
                        self.task_completed.emit(False, "无法连接SolidWorks", [], debug_info, False)
                        return

            try:
                sw_app.Visible = True
            except:
                pass

            if self.task_type == "connect":
                self.task_completed.emit(True, "已成功连接到SolidWorks", [], debug_info, False)
                return

            # 步骤3: 获取激活的模型
            model = sw_app.ActiveDoc
            if model is None:
                self.task_completed.emit(True, "未找到激活的模型", [], debug_info, False)
                return

            # 步骤4: 获取文件名
            filename = _safe_get(model, "GetTitle") or "(未知文件)"

            # 步骤5+6: 检测钣金件 + 收集折弯数据
            doc_type = _safe_get(model, "GetType")
            debug_info.append(f"文档类型: {doc_type}")

            if doc_type != 1:  # 非零件
                self.task_completed.emit(True, filename, [], debug_info, False)
                return

            # ============================================================
            # 核心策略：实体归属思维（多实体钣金适用）
            #   路径: Feature → GetFaces → Face.GetBody → Body.Name
            #   第1遍: 遍历所有特征，通过面→实体路径构建 Body→基体法兰 字典
            #   第2遍: 遍历顶层特征树，找到折弯→通过面→实体→查字典获取基体法兰
            #   注意: 统一使用_invoke/_com_attr底层COM调用（兼容CDispatch）
            # ============================================================

            # --- 辅助函数：使用_invoke底层COM调用（兼容GetObject返回的CDispatch）---
            def _feat_type(f):
                """获取特征类型名称，使用_invoke底层COM调用"""
                result = _invoke(f, "GetTypeName2")
                if result is not None and isinstance(result, str):
                    return result
                result = _invoke(f, "GetTypeName")
                if result is not None and isinstance(result, str):
                    return result
                return ""

            def _feat_name(f):
                """获取特征名称"""
                result = _com_attr(f, "Name")
                if result is not None and isinstance(result, str):
                    return result
                return ""

            def _next_feat(f):
                """获取下一个特征"""
                result = _invoke(f, "GetNextFeature")
                if result is not None:
                    return result
                return None

            def _first_sub(f):
                """获取第一个子特征"""
                result = _invoke(f, "GetFirstSubFeature")
                if result is not None:
                    return result
                return None

            def _next_sub(f):
                """获取下一个子特征"""
                result = _invoke(f, "GetNextSubFeature")
                if result is not None:
                    return result
                return None

            def _get_feat_body_name(feat):
                """通过特征的面获取所属实体名称（Feature→Faces→Face.GetBody→Body.Name）
                这是多实体钣金中唯一正确的折弯→基体法兰归属方法"""
                try:
                    faces = _invoke(feat, "GetFaces")
                    if faces is not None:
                        # COM数组可能不支持len()，用try/except包装
                        try:
                            fc = len(faces)
                        except Exception:
                            fc = 0
                        if fc > 0:
                            first_face = faces[0]
                            body = _invoke(first_face, "GetBody")
                            if body is not None:
                                body_name = _com_attr(body, "Name")
                                if body_name:
                                    return body_name
                except Exception:
                    pass
                return None

            def _get_feat_thickness_mm(feat):
                """从特征GetDefinition获取厚度(mm)"""
                try:
                    d = _invoke(feat, "GetDefinition")
                    if d is not None:
                        t = _com_attr(d, "Thickness")
                        if t is not None:
                            return t * 1000.0
                except Exception:
                    pass
                return 0.0

            # --- 第1遍: 检测钣金件 + 构建 Body→基体法兰 字典 ---
            is_sheet_metal = False
            feat_count = 0
            body_base_flange_map = {}  # {body_name: (base_flange_feat, base_flange_name, thickness_mm)}
            feat_type_samples = []  # 收集前20个特征类型用于调试

            try:
                feat = _invoke(model, "FirstFeature")
                while feat:
                    feat_count += 1
                    ftype = _feat_type(feat)
                    fname = _feat_name(feat)

                    # 收集特征类型样本用于调试
                    if feat_count <= 20:
                        feat_type_samples.append(f"{feat_count}:{ftype}({fname})")

                    # 检测钣金特征
                    if ftype in SHEET_METAL_IDENTIFY_TYPES:
                        is_sheet_metal = True
                        debug_info.append(f"找到钣金特征: {ftype}")

                    # 对基体法兰(SMBaseFlange)和SheetMetal特征：通过面获取实体，建立Body→基体法兰映射
                    if ftype in ("SMBaseFlange", "BaseFlange", "SheetMetal"):
                        body_name = _get_feat_body_name(feat)
                        if body_name:
                            if ftype in ("SMBaseFlange", "BaseFlange"):
                                # 基体法兰：记录名称和厚度
                                thk = _get_feat_thickness_mm(feat)
                                body_base_flange_map[body_name] = (feat, fname, thk)
                                debug_info.append(f"Body→基体法兰: '{body_name}' -> '{fname}', 厚度={thk:.2f}mm")
                            elif ftype == "SheetMetal" and body_name not in body_base_flange_map:
                                # SheetMetal 作为备选
                                body_base_flange_map[body_name] = (feat, fname, 0.0)
                                debug_info.append(f"Body→SheetMetal(备选): '{body_name}' -> '{fname}'")

                    feat = _next_feat(feat)
            except Exception as e:
                debug_info.append(f"第1遍遍历出错: {str(e)}")

            debug_info.append(f"特征总数: {feat_count}")
            debug_info.append(f"特征类型样本: {', '.join(feat_type_samples)}")
            debug_info.append(f"Body→基体法兰映射: {len(body_base_flange_map)} 条")

            # 回退方案1: 如果特征遍历未检测到钣金，尝试GetBendState
            if not is_sheet_metal:
                try:
                    bend_state = _com_attr(model, "GetBendState")
                    debug_info.append(f"GetBendState: {bend_state}")
                    if bend_state is not None and bend_state != 0:
                        is_sheet_metal = True
                        debug_info.append("通过GetBendState检测到钣金件")
                except Exception as e:
                    debug_info.append(f"GetBendState调用失败: {str(e)}")

            # 回退方案2: 通过Body.IsSheetMetal检测
            if not is_sheet_metal:
                try:
                    bodies = _invoke(model, "GetBodies2")
                    if bodies is not None:
                        try:
                            bc = len(bodies)
                        except Exception:
                            bc = 0
                        for bi in range(bc):
                            try:
                                body = bodies[bi]
                                is_sm = _com_attr(body, "IsSheetMetal")
                                if is_sm:
                                    is_sheet_metal = True
                                    debug_info.append(f"通过Body[{bi}].IsSheetMetal检测到钣金实体")
                                    break
                            except Exception:
                                pass
                except Exception as e:
                    debug_info.append(f"GetBodies2调用失败: {str(e)}")

            if not is_sheet_metal:
                debug_info.append("未找到钣金特征")
                self.task_completed.emit(True, filename, [], debug_info, False)
                return

            # --- 第2遍: 遍历特征树收集折弯数据（VBA逻辑 + 实体归属）---
            # VBA逻辑:
            #   While 顶层特征
            #     If 顶层是钣金特征类型 Then
            #       遍历子特征 → 找OneBend/SketchBend → 抓取数据
            #     End If
            #   顶层 = 顶层.GetNextFeature
            #   Wend
            # 实体归属:
            #   对找到的折弯特征 → _get_feat_body_name → 查body_base_flange_map → 获取基体名称和厚度

            bend_data_list = []
            index = 0

            def _collect_bend_from_feat(bend_feat, body_map):
                """从一个折弯特征抓取完整数据，通过实体归属获取基体信息"""
                nonlocal index
                bname = _feat_name(bend_feat)

                # 通过面→实体路径获取所属Body，查字典获取基体法兰信息
                body_name = _get_feat_body_name(bend_feat)

                if body_name and body_name in body_map:
                    _, base_name, base_thk = body_map[body_name]
                    debug_info.append(f"折弯[{index+1}] {bname}: 实体'{body_name}' -> 基体法兰='{base_name}', 厚度={base_thk:.2f}mm")
                else:
                    # 回退：从body_map取第一个
                    if body_map:
                        first_key = next(iter(body_map))
                        _, base_name, base_thk = body_map[first_key]
                    else:
                        base_name, base_thk = "", 0.0
                    debug_info.append(f"折弯[{index+1}] {bname}: 实体查找失败，回退默认基体: '{base_name}', 厚度={base_thk:.2f}mm")

                # 抓取折弯数据（使用_invoke底层COM调用）
                obfd = _invoke(bend_feat, "GetDefinition")
                if obfd is None:
                    debug_info.append(f"折弯[{index+1}] {bname}: GetDefinition返回None")
                    return

                cba = _invoke(obfd, "GetCustomBendAllowance")
                if cba is None:
                    debug_info.append(f"折弯[{index+1}] {bname}: GetCustomBendAllowance返回None")
                    return

                try:
                    allowance_type = _com_attr(cba, "Type")
                    if allowance_type is None:
                        debug_info.append(f"折弯[{index+1}] {bname}: cba.Type为None")
                        return

                    at_str = BEND_ALLOWANCE_TYPE_MAP.get(allowance_type, "K因子")
                    if allowance_type == 2:
                        av = _com_attr(cba, "KFactor") or 0
                    elif allowance_type == 4:
                        av = (_com_attr(cba, "BendDeduction") or 0) * 1000
                    else:
                        av = 0

                    ba = _com_attr(obfd, "BendAngle")
                    br = _com_attr(obfd, "BendRadius")
                    if ba is None or br is None:
                        debug_info.append(f"折弯[{index+1}] {bname}: obfd属性为None")
                        return

                    index += 1
                    angle_deg = ba * 180.0 / 3.14159265358979
                    radius_mm = br * 1000.0

                    bd = {
                        "index": index,
                        "sheet_metal_name": base_name,
                        "feature_name": bname,
                        "bend_angle": angle_deg,
                        "thickness": base_thk,
                        "bend_radius": radius_mm,
                        "allowance_type": at_str,
                        "allowance_value": av,
                        "feature_ptr": bend_feat,
                        "obfd_ptr": obfd,
                        "cba_ptr": cba,
                        "original_angle": ba,
                        "original_radius": br,
                        "original_allowance_type": allowance_type,
                        "original_allowance_value": av if allowance_type == 2 else (_com_attr(cba, "BendDeduction") or 0),
                    }
                    bend_data_list.append(bd)
                    debug_info.append(f"折弯[{index}]: {bname}, 角度={angle_deg:.1f}°, 基体名称={base_name}, 基体厚度={base_thk:.2f}mm, 半径={radius_mm:.2f}mm, 类型={at_str}")
                except Exception as e:
                    debug_info.append(f"折弯[{index+1}] {bname}: 读取属性出错: {str(e)}")

            try:
                top_feat = _invoke(model, "FirstFeature")
                while top_feat:
                    top_type = _feat_type(top_feat)
                    top_name = _feat_name(top_feat)

                    if top_type in SHEET_METAL_FEATURE_TYPES:
                        debug_info.append(f"遍历父特征'{top_name}'({top_type})")

                        # 遍历直接子特征（VBA标准方式）
                        sub = _first_sub(top_feat)
                        while sub:
                            stype = _feat_type(sub)
                            if stype in BEND_FEATURE_TYPES:
                                _collect_bend_from_feat(sub, body_base_flange_map)
                            # 递归子特征的子特征
                            sub2 = _first_sub(sub)
                            while sub2:
                                stype2 = _feat_type(sub2)
                                if stype2 in BEND_FEATURE_TYPES:
                                    _collect_bend_from_feat(sub2, body_base_flange_map)
                                sub2 = _next_sub(sub2)
                            sub = _next_sub(sub)
                    else:
                        # 非钣金顶层特征也检查子特征
                        sub = _first_sub(top_feat)
                        while sub:
                            stype = _feat_type(sub)
                            if stype in BEND_FEATURE_TYPES:
                                _collect_bend_from_feat(sub, body_base_flange_map)
                            sub = _next_sub(sub)

                    top_feat = _next_feat(top_feat)

                debug_info.append(f"共找到 {index} 个折弯特征, 有效数据 {len(bend_data_list)} 条")
            except Exception as e:
                debug_info.append(f"收集折弯数据出错: {str(e)}")

            self.task_completed.emit(True, filename, bend_data_list, debug_info, True)

        except Exception as e:
            debug_info.append(f"详细错误: {_format_exc()}")
            self.task_completed.emit(False, f"线程执行出错: {str(e)}", [], debug_info, False)



# ==================== 主窗口类 ====================
class BendManagerWindow(QMainWindow):
    """折弯管理器主窗口"""

    def __init__(self):
        super().__init__()
        self.bend_data_list = []
        self.worker_thread = None
        self.modify_worker = None      # 修改折弯参数的工作线程
        self.is_connected = False
        self.current_filename = ""

        self.init_ui()

    def init_ui(self):
        """初始化UI布局"""
        self.setWindowTitle("SW折弯管理器")
        self.setGeometry(100, 100, 1100, 800)

        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(10, 10, 10, 10)

        # 区域1：连接状态区
        self.create_connection_area(main_layout)

        # 区域2：折弯数据表格区
        self.create_bend_table_area(main_layout)

        # 区域3：日志区
        self.create_log_area(main_layout)

    def create_connection_area(self, parent_layout):
        """创建第一区域：连接状态区"""
        group_box = QGroupBox("SolidWorks 连接状态")
        layout = QVBoxLayout(group_box)

        # 状态行
        status_layout = QHBoxLayout()

        # 连接状态
        self.label_connection = QLabel("连接状态: 未连接")
        self.label_connection.setStyleSheet("color: red; font-weight: bold;")
        status_layout.addWidget(self.label_connection)

        # 文件名
        self.label_filename = QLabel("文件名: 未打开文件")
        status_layout.addWidget(self.label_filename)

        status_layout.addStretch()

        # 刷新按钮
        self.btn_refresh = QPushButton("刷新状态")
        self.btn_refresh.clicked.connect(self.on_refresh_clicked)
        self.btn_refresh.setFixedWidth(100)
        status_layout.addWidget(self.btn_refresh)

        layout.addLayout(status_layout)

        parent_layout.addWidget(group_box)

    def create_bend_table_area(self, parent_layout):
        """创建第二区域：折弯数据表格区"""
        group_box = QGroupBox("折弯数据列表")
        layout = QVBoxLayout(group_box)

        # 创建表格
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(8)
        self.table_widget.setHorizontalHeaderLabels([
            "折弯特征名称", "基体名称", "折弯角度(°)", "基体厚度(mm)",
            "折弯半径(mm)", "折弯系数类型", "折弯系数数值", "修改"
        ])

        # 设置列宽（均布）
        header = self.table_widget.horizontalHeader()
        for i in range(self.table_widget.columnCount()):
            header.setSectionResizeMode(i, QHeaderView.Stretch)

        layout.addWidget(self.table_widget)

        # 按钮行
        btn_layout = QHBoxLayout()

        self.btn_modify_all = QPushButton("统一修改")
        self.btn_modify_all.clicked.connect(self.on_modify_all_clicked)
        self.btn_modify_all.setFixedWidth(100)
        btn_layout.addWidget(self.btn_modify_all)

        self.btn_reload = QPushButton("重新加载")
        self.btn_reload.clicked.connect(self.on_reload_clicked)
        self.btn_reload.setFixedWidth(100)
        btn_layout.addWidget(self.btn_reload)

        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        parent_layout.addWidget(group_box)

    def create_log_area(self, parent_layout):
        """创建第三区域：日志区"""
        group_box = QGroupBox("日志")
        layout = QVBoxLayout(group_box)

        # 日志文本区
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        layout.addWidget(self.log_text)

        # 按钮行
        btn_layout = QHBoxLayout()

        self.btn_save_log = QPushButton("保存日志")
        self.btn_save_log.clicked.connect(self.save_log)
        self.btn_save_log.setFixedWidth(100)
        btn_layout.addWidget(self.btn_save_log)

        self.btn_clear_log = QPushButton("清空日志")
        self.btn_clear_log.clicked.connect(self.clear_log)
        self.btn_clear_log.setFixedWidth(100)
        btn_layout.addWidget(self.btn_clear_log)

        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        parent_layout.addWidget(group_box)

    def add_log(self, message):
        """添加日志消息"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.log_text.append(f"[{timestamp}] {message}")
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def save_log(self):
        """保存日志到文件"""
        if not self.log_text.toPlainText():
            self.add_log("日志为空，无需保存")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存日志", "bend_manager_log.txt", "Text Files (*.txt)"
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                self.add_log(f"日志已保存到: {file_path}")
            except Exception as e:
                self.add_log(f"保存日志失败: {str(e)}")

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
        self.add_log("日志已清空")

    def on_refresh_clicked(self):
        """刷新按钮点击事件——在同一条工作线程中完成连接+数据收集（避免COM跨线程问题）"""
        self.add_log("正在连接SolidWorks...")
        self.btn_refresh.setEnabled(False)

        self.worker_thread = WorkerThread("connect_and_collect")
        self.worker_thread.task_completed.connect(self.on_refresh_all_completed)
        self.worker_thread.start()

    def on_refresh_all_completed(self, success, message, data, debug_info=None, is_sheet_metal=False):
        """连接+收集数据完成的统一回调"""
        self.btn_refresh.setEnabled(True)

        if success:
            # 连接成功
            self.is_connected = True
            self.label_connection.setText("连接状态: 已连接")
            self.label_connection.setStyleSheet("color: green; font-weight: bold;")
            self.add_log("已成功连接到SolidWorks")

            # 显示调试信息
            if debug_info:
                for info in debug_info:
                    self.add_log(f"  调试: {info}")

            if is_sheet_metal:
                # 是钣金件
                self.current_filename = message
                self.label_filename.setText(f"文件名: {message}")
                self.add_log(f"文件: {message}")
                self.add_log("当前文件是钣金件")

                if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                    # 有折弯数据
                    self.bend_data_list = data
                    self.populate_table()
                    self.add_log(f"共找到 {len(self.bend_data_list)} 个折弯特征")
                else:
                    # 钣金件但无折弯数据
                    self.add_log("未找到折弯特征（可能该钣金件没有折弯）")
                    self.table_widget.setRowCount(0)
                    self.bend_data_list = []
            else:
                # 非钣金件
                self.label_filename.setText(f"文件名: {message}")
                self.add_log(f"文件: {message}")

                if isinstance(data, list) and data:
                    types_str = ", ".join(str(t) for t in data[:20])
                    self.add_log(f"检测到的特征类型: {types_str}")

                self.add_log("当前文件不是钣金件")
                self.table_widget.setRowCount(0)
                self.bend_data_list = []
        else:
            # 连接失败
            self.is_connected = False
            self.label_connection.setText("连接状态: 未连接")
            self.label_connection.setStyleSheet("color: red; font-weight: bold;")
            self.add_log(f"连接失败: {message}")
            if debug_info:
                for info in debug_info:
                    self.add_log(f"  调试: {info}")

    def collect_bend_data(self):
        """收集钣金折弯数据（用于重新加载按钮）"""
        self.add_log("正在收集钣金数据...")
        self.btn_refresh.setEnabled(False)

        self.worker_thread = WorkerThread("collect")
        self.worker_thread.task_completed.connect(self.on_collect_completed)
        self.worker_thread.start()

    def on_collect_completed(self, success, message, data, debug_info=None, is_sheet_metal=False):
        """数据收集完成回调"""
        self.btn_refresh.setEnabled(True)

        if success:
            self.current_filename = message
            self.label_filename.setText(f"文件名: {message}")
            self.add_log(f"当前文件: {message}")

            if debug_info:
                for info in debug_info:
                    self.add_log(f"  调试: {info}")

            if is_sheet_metal:
                self.add_log("当前文件是钣金件")
                self.bend_data_list = data if isinstance(data, list) and data and isinstance(data[0], dict) else []
                self.populate_table()
                self.add_log(f"共找到 {len(self.bend_data_list)} 个折弯特征")
            else:
                self.add_log("当前文件不是钣金件")
                self.table_widget.setRowCount(0)
                self.bend_data_list = []
        else:
            self.add_log(f"收集数据失败: {message}")
            if debug_info:
                for info in debug_info:
                    self.add_log(f"  调试: {info}")
            self.table_widget.setRowCount(0)
            self.bend_data_list = []

    def populate_table(self):
        """填充折弯数据表格"""
        self.table_widget.setRowCount(0)

        for bend_data in self.bend_data_list:
            row = self.table_widget.rowCount()
            self.table_widget.insertRow(row)

            # 折弯特征名称（只读）
            item_name = QTableWidgetItem(bend_data["feature_name"])
            item_name.setFlags(item_name.flags() & ~Qt.ItemIsEditable)
            self.table_widget.setItem(row, 0, item_name)

            # 基体名称（只读）
            item_sm_name = QTableWidgetItem(bend_data.get("sheet_metal_name", ""))
            item_sm_name.setFlags(item_sm_name.flags() & ~Qt.ItemIsEditable)
            self.table_widget.setItem(row, 1, item_sm_name)

            # 折弯角度（只读，BendAngle是SW API只读属性）
            item_angle = QTableWidgetItem(f"{bend_data['bend_angle']:.2f}")
            item_angle.setFlags(item_angle.flags() & ~Qt.ItemIsEditable)
            self.table_widget.setItem(row, 2, item_angle)

            # 厚度（只读）
            item_thickness = QTableWidgetItem(f"{bend_data['thickness']:.2f}")
            item_thickness.setFlags(item_thickness.flags() & ~Qt.ItemIsEditable)
            self.table_widget.setItem(row, 3, item_thickness)

            # 折弯半径（只读）
            item_radius = QTableWidgetItem(f"{bend_data['bend_radius']:.2f}")
            item_radius.setFlags(item_radius.flags() & ~Qt.ItemIsEditable)
            self.table_widget.setItem(row, 4, item_radius)

            # 折弯系数类型 - 使用下拉框（可修改）
            combo_type = QComboBox()
            combo_type.addItems(["K因子", "折弯扣除"])
            combo_type.setCurrentText(bend_data["allowance_type"])
            self.table_widget.setCellWidget(row, 5, combo_type)

            # 折弯系数数值（可修改）
            item_value = QTableWidgetItem(f"{bend_data['allowance_value']:.4f}")
            self.table_widget.setItem(row, 6, item_value)

            # 修改按钮
            btn_modify = QPushButton("修改")
            btn_modify.clicked.connect(lambda checked, r=row: self.on_modify_clicked(r))
            self.table_widget.setCellWidget(row, 7, btn_modify)

    def on_modify_clicked(self, row):
        """修改按钮点击事件 - 在工作线程中执行COM修改操作"""
        if row >= len(self.bend_data_list):
            return

        # 禁用修改按钮防止重复点击
        self._set_buttons_enabled(False)

        # 获取表格中的值
        allowance_type = self.table_widget.cellWidget(row, 5).currentText()
        try:
            allowance_value = float(self.table_widget.item(row, 6).text())
        except ValueError:
            self.add_log(f"第 {row + 1} 行数值格式错误，请输入有效数字")
            self._set_buttons_enabled(True)
            return

        feature_name = self.table_widget.item(row, 0).text()
        self.add_log(f"正在修改: {feature_name}...")

        # 创建修改工作线程（传递数据列表和行号，在工作线程中重新连接SW获取COM指针）
        self.modify_worker = ModifyBendWorker(
            task_type="single",
            bend_data_list=self.bend_data_list,
            row=row,
            new_type=allowance_type,
            new_value=allowance_value
        )
        self.modify_worker.single_modify_done.connect(self.on_single_modify_done)
        self.modify_worker.start()

    def on_modify_all_clicked(self):
        """统一修改按钮点击事件 - 按每行用户选择的系数类型和数值分别回填（等价分别点击每行修改按钮）"""
        if not self.bend_data_list:
            self.add_log("没有可修改的数据")
            return

        # 禁用按钮防止重复点击
        self._set_buttons_enabled(False)

        # 逐行从表格中读取用户选择的系数类型和数值，写入bend_data_list
        valid_rows = 0
        for row in range(len(self.bend_data_list)):
            combo = self.table_widget.cellWidget(row, 5)
            item_val = self.table_widget.item(row, 6)
            if combo is None or item_val is None:
                self.add_log(f"第 {row + 1} 行缺少控件，跳过")
                continue

            row_type = combo.currentText()
            try:
                row_value = float(item_val.text())
            except ValueError:
                self.add_log(f"第 {row + 1} 行数值格式错误，跳过")
                continue

            # 将UI层的选择写入bend_data，供线程读取
            self.bend_data_list[row]["ui_allowance_type"] = row_type
            self.bend_data_list[row]["ui_allowance_value"] = row_value
            valid_rows += 1

        if valid_rows == 0:
            self.add_log("没有有效的行数据可以修改")
            self._set_buttons_enabled(True)
            return

        self.add_log(f"开始统一修改: 共 {valid_rows} 行，按各行选择的类型和数值分别回填...")

        # 创建批量修改工作线程（不再传统一的new_type/new_value，由线程从bend_data_list逐行读取）
        self.modify_worker = ModifyBendWorker(
            task_type="batch",
            bend_data_list=self.bend_data_list,
            new_type="",       # 批量模式不再使用统一值
            new_value=0.0
        )
        self.modify_worker.progress_updated.connect(self.on_modify_progress)
        self.modify_worker.modify_completed.connect(self.on_modify_all_done)
        self.modify_worker.start()

    def on_modify_progress(self, row, message):
        """修改进度回调"""
        self.add_log(f"  {message}")

    def on_modify_all_done(self, success, success_count, message):
        """批量修改完成回调"""
        self._set_buttons_enabled(True)
        self.add_log(message)
        # 修改完成后不再自动重新加载，由用户自行点击"重新加载"按钮

    def on_single_modify_done(self, success, row, message):
        """单行修改完成回调"""
        self._set_buttons_enabled(True)
        self.add_log(message)
        # 修改完成后不再自动重新加载，由用户自行点击"重新加载"按钮

    def _set_buttons_enabled(self, enabled):
        """设置所有操作按钮的启用状态"""
        self.btn_modify_all.setEnabled(enabled)
        self.btn_refresh.setEnabled(enabled)
        self.btn_reload.setEnabled(enabled)

    def on_reload_clicked(self):
        """重新加载按钮点击事件"""
        if not self.is_connected:
            self.add_log("请先连接SolidWorks")
            return

        self.collect_bend_data()

    def closeEvent(self, event):
        """窗口关闭事件"""
        event.accept()


# ==================== 主程序入口 ====================
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 设置Metro风格
    app.setStyle("Fusion")

    window = BendManagerWindow()
    window.show()

    sys.exit(app.exec())