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
        # 使用InvokeTypes精确控制参数类型，解决某些COM对象PROPERTYPUT失败问题
        oleobj.InvokeTypes(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, (pythoncom.VT_EMPTY, 0), arg_descs, value)
        return True
    except Exception:
        # 回退到普通Invoke
        try:
            oleobj.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, value)
            return True
        except Exception:
            return False


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
SHEET_METAL_FEATURE_TYPES = [
    "EdgeFlange", "FlattenBends", "SMBaseFlange", "SMMiteredFlange",
    "Hem", "Jog", "SM3dBend", "LoftedBend", "SolidToSheetMetal"
]

# 折弯子特征类型
BEND_FEATURE_TYPES = ["OneBend", "SketchBend"]

# 钣金判断特征类型（参考VBA IsSheetMetalByFeature函数）
SHEET_METAL_IDENTIFY_TYPES = ["BaseFlange", "SMBaseFlange", "SheetMetal"]

# 折弯系数类型映射
BEND_ALLOWANCE_TYPE_MAP = {
    2: "KFactor",      # K因子
    4: "折弯扣除"       # 折弯扣除
}

BEND_ALLOWANCE_TYPE_REVERSE_MAP = {
    "KFactor": 2,
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
            import traceback
            self.debug_info.append(f"connect_to_sw详细错误: {traceback.format_exc()}")
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
                                allowance_type_str = BEND_ALLOWANCE_TYPE_MAP.get(allowance_type, "KFactor")

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


# ==================== 工作线程 ====================
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

                    at_str = BEND_ALLOWANCE_TYPE_MAP.get(allowance_type, "KFactor")
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
            debug_info.append(f"详细错误: {_tb.format_exc()}")
            self.task_completed.emit(False, f"线程执行出错: {str(e)}", [], debug_info, False)



# ==================== 主窗口类 ====================
class BendManagerWindow(QMainWindow):
    """折弯管理器主窗口"""

    def __init__(self):
        super().__init__()
        self.bend_data_list = []
        self.worker_thread = None
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
            combo_type.addItems(["KFactor", "折弯扣除"])
            combo_type.setCurrentText(bend_data["allowance_type"])
            self.table_widget.setCellWidget(row, 5, combo_type)

            # 折弯系数数值（可修改）
            item_value = QTableWidgetItem(f"{bend_data['allowance_value']:.4f}")
            self.table_widget.setItem(row, 6, item_value)

            # 修改按钮
            btn_modify = QPushButton("修改")
            btn_modify.clicked.connect(lambda checked, r=row: self.on_modify_clicked(r))
            self.table_widget.setCellWidget(row, 7, btn_modify)

    def on_restore_clicked(self, row):
        """还原按钮点击事件"""
        if row >= len(self.bend_data_list):
            return

        bend_data = self.bend_data_list[row]

        try:
            obfd = bend_data["obfd_ptr"]
            cba = bend_data["cba_ptr"]

            # 还原折弯半径（BendAngle是SW API只读属性，无需还原）
            if not _com_set(obfd, "BendRadius", bend_data["original_radius"]):
                raise Exception("无法设置折弯半径")

            # 还原折弯系数
            original_type = bend_data["original_allowance_type"]
            if not _com_set(cba, "Type", original_type):
                raise Exception("无法设置折弯系数类型")

            if original_type == 2:  # K因子
                if not _com_set(cba, "KFactor", bend_data["original_allowance_value"]):
                    raise Exception("无法设置K因子")
            elif original_type == 4:  # 折弯扣除
                if not _com_set(cba, "BendDeduction", bend_data["original_allowance_value"]):
                    raise Exception("无法设置折弯扣除")

            # 应用修改到特征
            model = self._get_sw_model()
            feature = bend_data.get("feature_ptr")
            if model and feature:
                feature.ModifyDefinition(obfd, model, None)

            self.add_log(f"已还原第 {row + 1} 行数据")

            # 刷新表格显示
            self.populate_table()

        except Exception as e:
            self.add_log(f"还原失败: {str(e)}")

    def _get_sw_model(self):
        """获取当前SolidWorks活动文档"""
        try:
            sw_app = win32com.client.Dispatch("SldWorks.Application")
            return sw_app.ActiveDoc
        except Exception:
            return None

    def _rename_sheet_metal_feature(self, old_name, new_name):
        """参考VBA逻辑修改钣金特征名称 (SelectByID2 + SelectedFeatureProperties)"""
        try:
            sw_app = win32com.client.Dispatch("SldWorks.Application")
            model = sw_app.ActiveDoc
            if model is None:
                raise Exception("没有激活的SolidWorks文档")

            # 选中旧的钣金特征 (BODYFEATURE)
            bool_status = model.Extension.SelectByID2(
                old_name, "BODYFEATURE", 0, 0, 0, False, 0, None, 0
            )
            if not bool_status:
                raise Exception(f"无法选中钣金特征 '{old_name}'，请确认名称正确")

            # 修改名称 (参考VBA: SelectedFeatureProperties)
            model.SelectedFeatureProperties(0, 0, 0, 0, 0, 0, 0, True, False, new_name)

            self.add_log(f"基体名称已修改: '{old_name}' -> '{new_name}'")
            return True
        except Exception as e:
            self.add_log(f"修改基体名称失败: {str(e)}")
            raise

    def on_modify_clicked(self, row):
        """修改按钮点击事件 - 只修改折弯系数类型和折弯系数数值"""
        if row >= len(self.bend_data_list):
            return

        bend_data = self.bend_data_list[row]
        feature_name = self.table_widget.item(row, 0).text()

        try:
            # 获取表格中的值（只获取可修改的字段）
            allowance_type = self.table_widget.cellWidget(row, 5).currentText()
            allowance_value = float(self.table_widget.item(row, 6).text())

            # 修改折弯数据
            obfd = bend_data["obfd_ptr"]
            cba = bend_data["cba_ptr"]

            # 只修改折弯系数类型和数值
            allowance_type_enum = BEND_ALLOWANCE_TYPE_REVERSE_MAP.get(allowance_type, 2)
            if not _com_set(cba, "Type", allowance_type_enum):
                raise Exception("无法设置折弯系数类型")

            if allowance_type_enum == 2:  # K因子
                if not _com_set(cba, "KFactor", allowance_value):
                    raise Exception("无法设置K因子")
            elif allowance_type_enum == 4:  # 折弯扣除
                if not _com_set(cba, "BendDeduction", allowance_value / 1000):
                    raise Exception("无法设置折弯扣除")

            # 应用修改到特征
            model = self._get_sw_model()
            feature = bend_data.get("feature_ptr")
            if model and feature:
                feature.ModifyDefinition(obfd, model, None)

            self.add_log(f"已修改第 {row + 1} 行: {feature_name} - 折弯系数类型: {allowance_type}, 数值: {allowance_value}")

        except Exception as e:
            self.add_log(f"修改失败: {str(e)}")

    def on_modify_all_clicked(self):
        """统一修改按钮点击事件 - 只修改折弯系数类型和折弯系数数值"""
        if not self.bend_data_list:
            self.add_log("没有可修改的数据")
            return

        success_count = 0
        fail_count = 0

        for row in range(len(self.bend_data_list)):
            try:
                # 获取表格中的值（只获取可修改的字段）
                allowance_type = self.table_widget.cellWidget(row, 5).currentText()
                allowance_value = float(self.table_widget.item(row, 6).text())

                bend_data = self.bend_data_list[row]
                obfd = bend_data["obfd_ptr"]
                cba = bend_data["cba_ptr"]

                # 只修改折弯系数类型和数值
                allowance_type_enum = BEND_ALLOWANCE_TYPE_REVERSE_MAP.get(allowance_type, 2)
                if not _com_set(cba, "Type", allowance_type_enum):
                    raise Exception("无法设置折弯系数类型")

                if allowance_type_enum == 2:  # K因子
                    if not _com_set(cba, "KFactor", allowance_value):
                        raise Exception("无法设置K因子")
                elif allowance_type_enum == 4:  # 折弯扣除
                    if not _com_set(cba, "BendDeduction", allowance_value / 1000):
                        raise Exception("无法设置折弯扣除")

                # 应用修改到特征
                model = self._get_sw_model()
                feature = bend_data.get("feature_ptr")
                if model and feature:
                    feature.ModifyDefinition(obfd, model, None)

                success_count += 1

            except Exception as e:
                fail_count += 1
                self.add_log(f"修改第 {row + 1} 行失败: {str(e)}")

        self.add_log(f"统一修改完成: 成功 {success_count} 条, 失败 {fail_count} 条")

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