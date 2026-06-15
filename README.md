# SW折弯管理器 (SW Bend Manager)

## 项目简介

SW折弯管理器是一个用于读取和修改SolidWorks钣金件折弯参数的Python工具。该工具可以帮助工程师高效地管理和批量修改钣金件中的折弯系数（K因子/折弯扣除）。

## 功能特性

- ✅ **连接SolidWorks** - 自动连接到运行中的SolidWorks实例
- ✅ **加载折弯数据** - 自动识别并读取钣金件中的所有折弯特征
- ✅ **数据展示** - 以表格形式展示折弯特征的详细信息
- ✅ **单行修改** - 修改选中行的折弯系数类型和数值
- ✅ **批量修改** - 一键批量修改所有折弯的系数类型和数值
- ✅ **日志记录** - 完整的操作日志记录和保存功能
- ✅ **进度显示** - 实时显示操作进度和耗时

## 技术参数

| 参数 | 说明 |
|------|------|
| 支持的SolidWorks版本 | 2021及以上 |
| 支持的折弯类型 | OneBend, SketchBend |
| 支持的系数类型 | KFactor, 折弯扣除 |
| 开发语言 | Python 3.x |
| GUI框架 | PySide6 |
| COM接口 | pywin32 |

## 支持的钣金特征类型

- EdgeFlange（边线法兰）
- FlattenBends（插入折弯）
- SMBaseFlange（基体法兰）
- SMMiteredFlange（斜接法兰）
- Hem（褶边）
- Jog（转折）
- SM3dBend（绘制的折弯）
- LoftedBend（放样折弯）
- SolidToSheetMetal（转换到钣金）

## 使用方法

### 环境要求

1. 安装Python 3.8+
2. 安装依赖库：
```bash
pip install pywin32 pyside6
```

### 运行方式

**方式一：直接运行Python脚本**
```bash
python sw_bend_manager-V2.3.py
```

**方式二：运行打包后的EXE**
```bash
./dist/sw_bend_manager-V2.3.exe
```

### 操作步骤

1. **连接SolidWorks**
   - 确保SolidWorks已启动并打开了一个钣金零件
   - 点击"连接SolidWorks"按钮

2. **加载折弯数据**
   - 连接成功后，点击"加载折弯数据"按钮
   - 程序会自动读取当前激活零件的所有折弯特征

3. **修改折弯参数**
   - **单行修改**：选中表格中的某一行，修改"系数类型"下拉框和"系数值"，然后点击"修改选中行"
   - **批量修改**：在"批量修改设置"区域设置目标类型和数值，然后点击"批量修改"

4. **保存日志**
   - 点击"保存日志"按钮可将操作日志保存到文件

## 项目结构

```
sw_bend_manager/
├── sw_bend_manager-V2.3.py    # 主程序（最新版）
├── sw_bend_manager-V2.2.py    # 旧版本
├── sw_bend_manager-V2.1.py    # 旧版本
├── sw_bend_manager-V2.0.py    # 旧版本
├── dist/                      # 打包后的可执行文件
│   └── sw_bend_manager-V2.2.exe
├── build/                     # 打包中间文件
├── 参考.txt                   # VBA参考代码
├── vba修改折弯参数的宏         # VBA宏参考
├── 异常报告.txt               # 异常报告
└── README.md                  # 项目说明文档
```

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| v2.3 | 2026-06-15 | 修复ModifyDefinition调用问题，使用三参数版本 |
| v2.2 | 2026-06-14 | 添加自定义折弯系数勾选步骤 |
| v2.1 | 2026-06-13 | 修改为抓取上一层法兰 |
| v2.0 | 2026-06-12 | 调整为只修改折弯系数 |
| v1.9 | 2026-06-11 | 数据抓取功能完善 |

## 常见问题

### Q: 连接SolidWorks失败？

**A:** 请确保：
1. SolidWorks已正确安装并启动
2. 当前有一个零件文档处于激活状态
3. 以管理员身份运行程序

### Q: 加载折弯数据时显示"当前文件不是钣金件"？

**A:** 请确保当前打开的是一个钣金零件（包含BaseFlange或SMBaseFlange特征）

### Q: 修改折弯参数失败？

**A:** 可能的原因：
1. 特征名称在模型中已被修改，请重新加载数据
2. SolidWorks模型被锁定，请检查模型状态
3. 权限不足，请以管理员身份运行

## 开发说明

### 核心代码结构

```python
# COM工具函数（位于文件开头）
- _get_oleobj()      # 获取底层OLE对象
- _invoke()          # 调用COM方法
- _com_attr()        # 获取COM属性
- _com_set()         # 设置COM属性
- _com_method()      # 调用COM方法并传参

# 业务逻辑类
- BendDataProcessor  # 折弯数据处理器（数据抓取）
- ModifyBendWorker   # 修改折弯参数工作线程
- WorkerThread       # 数据收集工作线程

# GUI组件
- LogAreaWidget      # 日志区域组件
- MainWindow         # 主窗口
```

### 修改折弯参数的流程

1. 获取折弯特征的Definition对象
2. 获取CustomBendAllowance对象
3. 设置UseBendTable=False（禁用折弯表）
4. 根据类型设置UseKFactor=True或UseBendDeduction=True
5. 设置Type属性（2=KFactor, 4=折弯扣除）
6. 设置对应的系数数值
7. 调用ModifyDefinition(obfd, model, None)应用修改
8. 调用EditRebuild3重建模型

## 许可证

本项目仅供内部使用和学习参考。

## 联系方式

如有问题或建议，请联系开发者。

---

*制作人: wsl*
*最后更新: 2026-06-15*