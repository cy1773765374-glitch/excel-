# Excel 图片批量导出模块（Xbot / Python）

用于从 Excel 指定工作表中，按“行”批量导出图片到本地目录。  
支持两类 Excel 图片来源：

- **传统浮动图片（Drawing / Shape 图片）**：openpyxl 可直接读取并导出
- **嵌入单元格图片 / 单元格图片类型**：openpyxl 可能拿不全，自动切换 **Excel COM** 导出（更稳）

> 适配 Xbot 可视化流程「调用模块」方式，也可作为独立 Python 模块调用。

---

## 功能概览

- ✅ 按行导出图片：`row_10_1.jpg / row_10_2.png ...`
- ✅ 一行多图：同一行多张图片自动编号 `_1/_2/...`
- ✅ 图片命名来自指定列（默认 B 列）
- ✅ 图片所在列可过滤（默认 A 列），并支持“列容差”（A/B/C 锚点偏移也能匹配）
- ✅ 双引擎导出：
  - `openpyxl`：速度快，适合浮动图片
  - `Excel COM`：兼容性强，适合嵌入单元格图片 / openpyxl 识别不全的情况
  - `auto`：先 openpyxl，再 COM 补齐未导出行（推荐）

---

## 运行环境与依赖

### 1) Python 依赖
- `openpyxl`
- `pillow`（用于保存图片）
- `pywin32`（**仅 Windows + COM 导出需要**）
- `xbot`（如果你在 Xbot 里运行）

安装示例：

```bash
pip install openpyxl pillow
pip install pywin32   # 需要 Excel COM 功能时（Windows）



## 为什么需要两套导出方式？

Excel 里的“图片”在内部并不只有一种形态，不同形态对应不同读取方式：

### 1) 浮动图片 / Shape / Drawing 图片
- 通常是**插入后漂浮在单元格上**（或“随单元格移动/缩放”的传统图片对象）
- 这类图片一般能被 `openpyxl` 识别到，出现在 `ws._images` 中
- 因此可以直接用 **openpyxl** 导出（速度快、依赖少）

### 2) 嵌入单元格图片 / 单元格图片类型
- Excel 新式“图片在单元格里”的数据结构（肉眼看起来像在格子里）
- `openpyxl` 往往**拿不全或拿不到**（有时只识别到一小部分）
- 这时用 **Excel COM**（让 Excel 本体来复制/导出）最稳

---

## `auto` 模式如何工作？

`auto` 模式的目标是：**尽可能完整导出**，同时避免重复。

执行流程如下：

1. 先用 **openpyxl** 导出能识别到的图片行  
2. 记录 openpyxl 已经导出的**行号集合**
3. 再用 **Excel COM** 只导出“openpyxl 未导出的行”  
4. 达到“补齐”的效果，并避免重复导出

---

## 使用方式

### 方式 A：Xbot 可视化流程（推荐）

1. 将本模块放入 Xbot 项目中  
2. 在可视化流程里使用「调用模块」  
3. 选择函数：`export_images_by_row`  
4. 填写参数并运行  

---

## 必填参数说明

| 参数名 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `xlsx_path` | `str` | 无 | Excel 文件路径（建议 `.xlsx`） |
| `imgSavePath` | `str` | 无 | 图片导出目录（必须存在） |
| `sheetName` | `str` | `Sheet2` | 工作表名称 |
| `nameCol` | `str` | `B` | 图片命名来源列（每行用这列的文本命名） |
| `imgCol` | `str` | `A` | 图片锚点所在列（一般是图片列） |
| `startRow` | `int` | `1` | 从第几行开始处理 |

---

## 推荐可选参数

| 参数名 | 类型 | 默认值 | 说明 |
|---|---|---|---|
| `engine` | `str` | `auto` | `auto / openpyxl / com` 三选一 |
| `colTolerance` | `int` | `2` | 锚点列允许偏移量（`A±2 => A/B/C`） |
| `debug` | `int` | `0` | `1` 输出更详细日志 |

---

## 推荐参数组合

### 默认推荐（兼容性最好）
- `engine="auto"`
- `colTolerance=5`
- `debug=1`（首次排查用，跑通后改 `0`）

### 你确认都是“嵌入单元格图片”
- `engine="com"`

---

## 方式 B：Python 直接调用

```python
from your_module import export_images_by_row

export_images_by_row(
    xlsx_path=r"C:\xxx\test.xlsx",
    imgSavePath=r"C:\xxx\out",
    sheetName="Sheet2",
    nameCol="B",
    imgCol="A",
    startRow=1,
    colTolerance=5,
    debug=1,
    engine="auto",
)
