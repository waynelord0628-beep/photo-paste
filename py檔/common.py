"""
共用工具模組 - 供所有圖片彙整腳本使用
"""

import sys
import os
import re
import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.oxml.ns import qn

# ── HEIC 支援（選用）────────────────────────────────────────────
# 若已安裝 pillow-heif，自動啟用 HEIC/HEIF 支援
try:
    import pillow_heif

    pillow_heif.register_heif_opener()
    _HEIC_SUPPORTED = True
except ImportError:
    _HEIC_SUPPORTED = False

from PIL import Image


# ── 路徑工具 ────────────────────────────────────────────────────


def get_base_path():
    """取得執行檔或 .py 所在資料夾路徑（相容 PyInstaller .exe）"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_unique_filename(folder, base_name, ext):
    """若檔名已存在，自動在後方加序號避免覆蓋"""
    filename = f"{base_name}{ext}"
    counter = 1
    while os.path.exists(os.path.join(folder, filename)):
        filename = f"{base_name}_{counter}{ext}"
        counter += 1
    return filename


# ── 圖片清單 ────────────────────────────────────────────────────

# python-docx 可直接插入的格式
_DOCX_NATIVE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp"}

# 需要透過 Pillow 轉換才能插入的格式
_PILLOW_CONVERT_EXTS = {".webp", ".tiff", ".tif"}

# HEIC/HEIF（需要 pillow-heif）
_HEIC_EXTS = {".heic", ".heif"}


def _build_valid_exts():
    """動態建立支援的副檔名集合（不分大小寫）"""
    exts = _DOCX_NATIVE_EXTS | _PILLOW_CONVERT_EXTS
    if _HEIC_SUPPORTED:
        exts = exts | _HEIC_EXTS
    # 同時加入大寫版本
    return exts | {e.upper() for e in exts}


VALID_IMAGE_EXTS = _build_valid_exts()


def load_images(path_captures):
    """
    從指定資料夾讀取合法圖片，回傳 (完整路徑列表, 無副檔名名稱列表)。
    若資料夾不存在會 raise FileNotFoundError。
    """
    if not os.path.exists(path_captures):
        raise FileNotFoundError(f"找不到 照片放這 資料夾，請確認路徑：{path_captures}")
    names = sorted(
        fn
        for fn in os.listdir(path_captures)
        if os.path.splitext(fn)[1].lower() in {e.lower() for e in VALID_IMAGE_EXTS}
    )
    paths = [os.path.join(path_captures, fn) for fn in names]
    names_noext = [os.path.splitext(fn)[0] for fn in names]
    return paths, names_noext


# ── 圖片插入輔助 ─────────────────────────────────────────────────


def open_image_as_stream(img_path):
    """
    將圖片轉成 python-docx 可插入的 bytes stream（PNG 格式）。
    - JPEG/PNG：直接讀取原始檔案（保持品質）
    - HEIC、WEBP、TIFF 等其他格式：用 Pillow 轉成 PNG bytes
    """
    ext = os.path.splitext(img_path)[1].lower()
    if ext in (".jpg", ".jpeg", ".png"):
        # 原生支援，直接回傳檔案路徑即可（讓 add_picture 直接讀檔）
        return img_path
    # 其他格式：用 Pillow 讀取後輸出成 PNG bytes
    with Image.open(img_path) as img:
        # RGBA 轉 RGB（Word 不支援帶透明通道的 PNG 在某些情況下）
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGBA")
        else:
            img = img.convert("RGB")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf


# ── Word 文件操作 ────────────────────────────────────────────────


def open_template(template_path):
    """開啟 Word 模板，若找不到則 raise FileNotFoundError"""
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板檔，請確認路徑：{template_path}")
    return docx.Document(template_path)


def setup_header(document, section, title_text):
    """
    設定頁首：距離、標題文字（標楷體 20pt 置中）、下方空行。
    """
    section.header_distance = Cm(1.1)
    section.footer_distance = Cm(0.3)

    header = section.header
    if header.paragraphs:
        p_header = header.paragraphs[0]
        p_header.clear()
    else:
        p_header = header.add_paragraph()

    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.style = document.styles["Normal"]

    fmt = p_header.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = Pt(20)

    run = p_header.add_run(title_text)
    run.font.size = Pt(20)
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "標楷體")

    header.add_paragraph("")


def set_run_font(
    run, size_pt=14, western_font="Times New Roman", east_asia_font="標楷體"
):
    """設定 run 的字體大小與中西文字型"""
    run.font.size = Pt(size_pt)
    run.font.name = western_font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), east_asia_font)


def delete_first_paragraph_if_empty(doc):
    """刪除文件開頭的空白段落（Word 模板常見殘留）"""
    if not doc.paragraphs:
        return
    first_p = doc.paragraphs[0]
    if not first_p.text.strip():
        elem = first_p._element
        elem.getparent().remove(elem)


def delete_trailing_empty_paragraphs(doc):
    """
    刪除文件結尾所有空白段落。

    python-docx 的 add_section() 會在 body 尾端插入一個帶有
    <w:sectPr type="nextPage"> 的空段落，Word 開啟時會把這個
    段落視為「下一頁的起點」，造成多出一頁空白。
    將這些段落移除後，分節符會自動併入文件最終的 <w:sectPr>，
    空白頁就消失了。
    """
    from docx.oxml.ns import qn as _qn

    body = doc.element.body
    # 從後往前掃，遇到非空段落就停止
    while True:
        paras = doc.paragraphs
        if not paras:
            break
        last_p = paras[-1]
        # 若段落含有文字或圖片等內容，停止刪除
        if last_p.text.strip():
            break
        # 段落沒有文字，但若其 <w:pPr><w:sectPr> 是文件最終 sectPr，不能刪
        # （最終 sectPr 直接掛在 <w:body> 下，不在段落內）
        elem = last_p._element
        # 只刪除「段落本身帶有 sectPr（即分節段落）」或「完全空白的段落」
        elem.getparent().remove(elem)


# ── 儲存文件 ────────────────────────────────────────────────────


def save_document(document, path_now, title_text):
    """
    將文件儲存到 path_now 目錄，檔名依標題產生（去除非法字元）。
    回傳最終輸出的完整路徑。
    """
    safe_title = re.sub(r'[\\/*?:"<>|]', "_", title_text)
    filename = get_unique_filename(path_now, safe_title, ".docx")
    output_path = os.path.join(path_now, filename)
    document.save(output_path)
    print(f"完成！輸出檔案：{output_path}")
    return output_path
