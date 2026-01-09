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
