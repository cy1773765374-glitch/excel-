# 使用提醒:
# 1. xbot包提供软件自动化、数据表格、Excel、日志、AI等功能
# 2. package包提供访问当前应用数据的功能，如获取元素、访问全局变量、获取资源文件等功能
# 3. 当此模块作为流程独立运行时执行main函数
# 4. 可视化流程中可以通过"调用模块"的指令使用此模块

import os
import re
from xbot import print as xprint, sleep


# =========================
# 安全打印：永远只传 1 个字符串参数给 xbot.print
# =========================
def xlog(msg):
    try:
        xprint(str(msg))
    except Exception:
        try:
            import builtins
            builtins.print(str(msg))
        except Exception:
            pass


# =========================
# 列字母 <-> 列号
# =========================
def _col_to_index(col):
    """
    'A'->1, 'B'->2, 'AA'->27
    也支持直接传入数字/数字字符串：'2'->2
    """
    if col is None:
        raise ValueError("列标为空")

    # 允许直接传数字
    try:
        if isinstance(col, (int, float)):
            return int(col)
        s = str(col).strip()
        if s.isdigit():
            return int(s)
    except Exception:
        pass

    s = str(col).strip().upper()
    if not s or not s.isalpha():
        raise ValueError("列标不合法: {}".format(col))

    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)  # 'A'->1
    return n


def _index_to_col(n):
    n = int(n)
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# =========================
# 文件名清洗 + 防覆盖路径
# =========================
def _safe_filename(name, default="img"):
    if name is None:
        name = default
    name = str(name).strip()
    if not name:
        name = default
    return re.sub(r'[\\/:*?"<>|]+', "_", name)


def _unique_path(dir_path, filename):
    base, ext = os.path.splitext(filename)
    p = os.path.join(dir_path, filename)
    if not os.path.exists(p):
        return p
    for n in range(2, 9999):
        p2 = os.path.join(dir_path, "{}_{}{}".format(base, n, ext))
        if not os.path.exists(p2):
            return p2
    return p


# =========================
# Chart 粘贴导出（剪贴板 -> Chart -> Export）
# =========================
def _chart_export_from_clipboard(xl_ws, width, height, save_path):
    co = xl_ws.ChartObjects().Add(0, 0, max(20, int(width)), max(20, int(height)))
    try:
        chart = co.Chart
        try:
            chart.ChartArea.Select()
        except Exception:
            pass
        sleep(0.1)
        chart.Paste()
        sleep(0.1)
        chart.Export(save_path)
    finally:
        try:
            co.Delete()
        except Exception:
            pass


def _looks_like_blank_file(p, min_kb):
    try:
        size = os.path.getsize(p)
        return size < int(min_kb) * 1024
    except Exception:
        return True


def _try_export_cell_picture(xl_app, xl_ws, rng, save_path, debug=0, min_kb=8, retries=3):
    """
    对单元格执行 CopyPicture 并导出 PNG
    - 若单元格合并，取 MergeArea 顶格
    - Appearance 两种模式都试：Screen(1)/Printer(2)
    - 重试 + 空白文件检测（太小就删掉重试）
    """
    try:
        if bool(rng.MergeCells):
            rng = rng.MergeArea.Cells(1, 1)
    except Exception:
        pass

    try:
        w = rng.Width
        h = rng.Height
    except Exception:
        w, h = 300, 300

    appearances = [1, 2]  # xlScreen, xlPrinter

    for attempt in range(1, int(retries) + 1):
        for ap in appearances:
            try:
                xl_app.CutCopyMode = False
            except Exception:
                pass

            try:
                rng.CopyPicture(Appearance=ap, Format=2)  # xlBitmap
                sleep(0.15)

                _chart_export_from_clipboard(xl_ws, w, h, save_path)

                if _looks_like_blank_file(save_path, min_kb=min_kb):
                    try:
                        os.remove(save_path)
                    except Exception:
                        pass
                    if int(debug) == 1:
                        xlog("EMPTY -> retry (attempt={}, appearance={}, cell={})".format(
                            attempt, ap, getattr(rng, "Address", "?")
                        ))
                    continue

                return True

            except Exception as e:
                if int(debug) == 1:
                    xlog("Copy/Export failed (attempt={}, appearance={}, cell={}): {}".format(
                        attempt, ap, getattr(rng, "Address", "?"), repr(e)
                    ))
                continue

    return False


def _get_last_row_strong(xl_ws, fallback_col_idx):
    """
    比 End(xlUp) 更可靠：LastCell/UsedRange 能覆盖“只有图片/格式但没值”的行
    """
    try:
        last_cell = xl_ws.Cells.SpecialCells(11)  # xlCellTypeLastCell = 11
        return int(last_cell.Row)
    except Exception:
        pass

    try:
        ur = xl_ws.UsedRange
        return int(ur.Row + ur.Rows.Count - 1)
    except Exception:
        pass

    try:
        return int(xl_ws.Cells(xl_ws.Rows.Count, int(fallback_col_idx)).End(-4162).Row)  # xlUp
    except Exception:
        return 1


def _get_merged_name_in_col(xl_ws, row, name_col_idx, last_name):
    """
    A列可能合并：取 MergeArea 左上角值
    若仍为空，用 last_name 兜底
    """
    raw = None
    try:
        cell = xl_ws.Cells(int(row), int(name_col_idx))
        if bool(cell.MergeCells):
            raw = cell.MergeArea.Cells(1, 1).Value
        else:
            raw = cell.Value
    except Exception:
        raw = None

    raw = _safe_filename(raw, default="img")
    if raw == "img" and last_name not in [None, "", "img"]:
        return last_name
    return raw


# =========================
# 核心导出（你的最新结构）
# - Sheet1
# - A列名称（可能合并）
# - B列图片（嵌入单元格）
# 命名：{A列名称}_{B2}.png 例如：水表规格与位置_B2.png
# =========================
def export_images_by_row(
        xlsx_path,
        imgSavePath,
        sheetName="Sheet1",
        nameCol="A",
        imgCol="B",
        startRow=1,
        debug=0,
        min_kb=8,
        retries=3
):
    if not os.path.isfile(xlsx_path):
        xlog("xlsx_path 不存在: {}".format(xlsx_path))
        return False

    if not os.path.isdir(imgSavePath):
        xlog("imgSavePath 目录不存在: {}".format(imgSavePath))
        return False

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        xlog("缺少 pywin32，无法使用 Excel COM。err={}".format(repr(e)))
        return False

    name_col_idx = _col_to_index(nameCol)
    img_col_idx = _col_to_index(imgCol)
    img_col_letter = _index_to_col(img_col_idx)

    xl = None
    wb = None
    exported = 0

    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False

        wb = xl.Workbooks.Open(xlsx_path)
        ws = wb.Worksheets(sheetName)

        # 用 LastCell/UsedRange 推 last_row，避免“图片没值导致行数偏小”
        last_row = _get_last_row_strong(ws, fallback_col_idx=name_col_idx)
        xlog("COM last_row: {}".format(last_row))

        last_name = None

        for r in range(int(startRow), int(last_row) + 1):
            base_name = _get_merged_name_in_col(ws, r, name_col_idx, last_name)
            last_name = base_name

            cell_coord = "{}{}".format(img_col_letter, r)
            filename = "{}_{}.png".format(base_name, cell_coord)
            save_path = _unique_path(imgSavePath, filename)

            try:
                rng = ws.Cells(int(r), int(img_col_idx))
            except Exception:
                continue

            ok = _try_export_cell_picture(
                xl_app=xl,
                xl_ws=ws,
                rng=rng,
                save_path=save_path,
                debug=debug,
                min_kb=min_kb,
                retries=retries
            )

            if ok:
                exported += 1
                xlog("OK: {}".format(save_path))
            else:
                if int(debug) == 1:
                    xlog("SKIP(no picture): {}".format(cell_coord))

        xlog("导出数量: {}".format(exported))
        return exported > 0

    except Exception as e:
        xlog("执行异常: {}".format(repr(e)))
        return False

    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if xl is not None:
                xl.Quit()
        except Exception:
            pass


# =========================
# xbot 入口
# =========================
def main(args):
    try:
        xlsx_path = args.get("xlsx_path", "")
        imgSavePath = args.get("imgSavePath", "")
        sheetName = args.get("sheetName", "Sheet1")
        nameCol = args.get("nameCol", "A")
        imgCol = args.get("imgCol", "B")
        startRow = args.get("startRow", 1)

        debug = args.get("debug", 0)
        min_kb = args.get("min_kb", 8)
        retries = args.get("retries", 3)

        return export_images_by_row(
            xlsx_path=xlsx_path,
            imgSavePath=imgSavePath,
            sheetName=sheetName,
            nameCol=nameCol,
            imgCol=imgCol,
            startRow=startRow,
            debug=debug,
            min_kb=min_kb,
            retries=retries
        )
    except Exception as e:
        xlog("main 执行异常: {}".format(repr(e)))
        return False
