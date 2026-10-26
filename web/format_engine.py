# -*- coding: utf-8 -*-
"""
Engine chuan hoa do an / bao cao Word (.docx) cho sinh vien Viet Nam.

Doc file .docx "lon xon" (font co chu lung tung, tieu de go tay, khong muc luc,
khong so trang) va tra ve file da chuan hoa theo mot bo quy dinh trinh bay:
  - Kho giay A4 + can le chuan
  - Font Times New Roman, co chu, gian dong, thut dau dong, can deu 2 ben
  - Nhan dien tieu de chuong/muc (theo style Heading hoac theo mau danh so
    "1.", "1.1", "CHUONG 1", "MO DAU", "TAI LIEU THAM KHAO"...)
  - Chuan hoa chu thich Hinh/Bang
  - Chen so trang o chan trang
  - Chen Muc luc tu dong (truong TOC cua Word, tu cap nhat khi mo file)

Dung: python format_engine.py input.docx output.docx [preset]
"""
import re
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

BLACK = RGBColor(0, 0, 0)

# ---------------------------------------------------------------- presets ----
PRESETS = {
    "chuan_vn": {
        "label": "Chuẩn phổ thông VN (đồ án/báo cáo)",
        "page": {"width_cm": 21.0, "height_cm": 29.7,
                 "top_cm": 2.5, "bottom_cm": 2.5, "left_cm": 3.5, "right_cm": 2.0},
        "font_name": "Times New Roman",
        "body": {"size": 13, "line_spacing": 1.3, "space_before_pt": 0,
                 "space_after_pt": 6, "first_line_indent_cm": 1.0,
                 "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY},
        "chapter": {"size": 16, "bold": True, "caps": True,
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before_pt": 12, "space_after_pt": 12},
        "h1": {"size": 14, "bold": True, "space_before_pt": 12, "space_after_pt": 6,
               "left_indent_cm": 0.0},
        "h2": {"size": 13, "bold": True, "space_before_pt": 6, "space_after_pt": 4,
               "left_indent_cm": 0.5},
        "h3": {"size": 13, "bold": True, "italic": True,
               "space_before_pt": 6, "space_after_pt": 4, "left_indent_cm": 1.0},
        "h4": {"size": 13, "bold": False, "italic": True,
               "space_before_pt": 6, "space_after_pt": 4, "left_indent_cm": 1.5},
        "caption": {"size": 12, "italic": True,
                    "alignment": WD_ALIGN_PARAGRAPH.CENTER,
                    "space_before_pt": 4, "space_after_pt": 8},
        "table_text": {"size": 12, "line_spacing": 1.0},
        "reference": {"size": 13, "line_spacing": 1.3, "space_after_pt": 4,
                      "hanging_indent_cm": 1.0},
        "page_number": {"size": 12, "alignment": WD_ALIGN_PARAGRAPH.CENTER},
        "desc": "Quy định trình bày phổ biến của các trường đại học Việt Nam.",
    }
}


def _variant(label, desc, **over):
    """Tao preset moi tu preset goc, ghi de cac muc theo duong dan cham (vd body.size=14)."""
    import copy
    p = copy.deepcopy(PRESETS["chuan_vn"])
    p["label"] = label
    p["desc"] = desc
    for path, val in over.items():
        keys = path.split(".")
        node = p
        for k in keys[:-1]:
            node = node[k]
        node[keys[-1]] = val
    return p


PRESETS["do_an"] = _variant(
    "Đồ án tốt nghiệp",
    "Chuẩn đồ án: lề trái 3,5 cm để đóng gáy, chương in hoa canh giữa.",
)
PRESETS["khoa_luan"] = _variant(
    "Khóa luận tốt nghiệp",
    "Bản dài: giãn dòng 1.5 dễ đọc, trình bày trang trọng như đồ án.",
    **{"body.line_spacing": 1.5})
PRESETS["nckh"] = _variant(
    "Nghiên cứu khoa học",
    "Báo cáo NCKH: tiêu đề mục đánh số kiểu bài báo, canh trái, gọn gàng.",
    **{"page.left_cm": 3.0, "page.right_cm": 2.0,
       "chapter.alignment": WD_ALIGN_PARAGRAPH.LEFT, "chapter.size": 14,
       "chapter.space_before_pt": 10, "chapter.space_after_pt": 6})
PRESETS["thuc_tap"] = _variant(
    "Báo cáo thực tập",
    "Báo cáo thực tập tại doanh nghiệp: chuẩn phổ thông, chương 15pt.",
    **{"chapter.size": 15})
# Tieu luan mon hoc — do truc tiep tu bai mau chuan anh Kord chi dinh (nhom 7,
# DHXDGT21A, "Hien tuong truot dat", 26/10/2026): A4; phan noi dung le 2,5/2/3/2
# (phan dau tach section rieng le 2,54 — tai lieu nhieu section thi tool giu nguyen le);
# than van TNR 13 thuong, can deu. NHIP THAN VAN CHU DAO cua bai mau (101 doan noi
# dung): gian dong 1.3, Before 6 / After 0, thut dau dong 1,0 cm — trung khop chuan
# thuyet minh IUH. (Reng cum Loi Mo Dau dung 1.5/6/6/1,25 — lech cuc bo, khong theo.)
# Chuong 13 dam canh giua (ten chuong in hoa); muc 1. = 13 dam; 1.1 = 13 dam;
# 1.1.1 = 13 dam nghieng thut 0,5; 1.1.1.1 = 12 nghieng; chu thich hinh/bang 13
# nghieng (khong dam) canh giua; so trang o goc duoi-phai.
PRESETS["tieu_luan"] = _variant(
    "Tiểu luận môn học",
    "Theo bài tiểu luận mẫu chuẩn: TNR 13, giãn dòng 1.3, lề 2,5/2/3/2, số trang góc dưới-phải.",
    **{"page.top_cm": 2.5, "page.bottom_cm": 2.0, "page.left_cm": 3.0, "page.right_cm": 2.0,
       "body.size": 13, "body.line_spacing": 1.3,
       "body.space_before_pt": 6, "body.space_after_pt": 0,
       "body.first_line_indent_cm": 1.0,
       "chapter.size": 13, "chapter.space_before_pt": 6, "chapter.space_after_pt": 6,
       "h1.size": 13, "h1.space_before_pt": 6, "h1.space_after_pt": 0,
       "h2.size": 13, "h2.italic": False, "h2.left_indent_cm": 0.0,
       "h2.space_before_pt": 6, "h2.space_after_pt": 0,
       "h3.size": 13, "h3.left_indent_cm": 0.5,
       "h3.space_before_pt": 6, "h3.space_after_pt": 0,
       "h4": {"size": 12, "bold": False, "italic": True,
              "space_before_pt": 6, "space_after_pt": 0, "left_indent_cm": 1.0},
       "caption.size": 13, "caption.bold": False, "caption.italic": True,
       "page_number.alignment": WD_ALIGN_PARAGRAPH.RIGHT})

# Thuyet minh khoa luan/do an — trich nguyen van tu mot file mau thuyet minh
# (04_Thuyetminh_Doan_BTL.docx): A4; le tren 2,5 / duoi 2 / trai 3 / phai 2 cm;
# Times New Roman 13; gian dong 1.3; doan cach nhau Before 6pt / After 0pt;
# doan dau moi muc khong thut dau dong, cac doan sau thut 1 cm;
# chuong 14 dam canh giua; muc 1.1 dam dung; 1.1.1 dam nghieng; 1.1.1.1 nghieng thuong;
# chu thich Hinh/Bang 13 dam nghieng canh giua; thut le muc: 0 / 0,5 / 1,0 cm.
PRESETS["iuh"] = _variant(
    "Thuyết minh khóa luận/đồ án",
    "Đúng quy định trong file mẫu thuyết minh: TNR 13, giãn dòng 1.3, lề 2,5/2/3/2.",
    **{"page.top_cm": 2.5, "page.bottom_cm": 2.0, "page.left_cm": 3.0, "page.right_cm": 2.0,
       "body.size": 13, "body.line_spacing": 1.3,
       "body.space_before_pt": 6, "body.space_after_pt": 0,
       "body.first_line_indent_cm": 1.0, "body.first_para_no_indent": True,
       "chapter.size": 14, "chapter.space_after_pt": 6,
       "h1.size": 13, "h2.size": 13, "h2.italic": False, "h2.left_indent_cm": 0.0,
       "h3.size": 13, "h3.left_indent_cm": 0.5,
       "h4": {"size": 13, "bold": False, "italic": True,
              "space_before_pt": 6, "space_after_pt": 4, "left_indent_cm": 1.0},
       "caption.size": 13, "caption.bold": True, "caption.italic": True})

FRONT_MATTER = {
    "MỞ ĐẦU", "KẾT LUẬN", "KẾT LUẬN VÀ KIẾN NGHỊ", "TÀI LIỆU THAM KHẢO",
    "PHỤ LỤC", "DANH MỤC BẢNG", "DANH MỤC HÌNH", "DANH MỤC HÌNH ẢNH",
    "DANH MỤC BẢNG BIỂU", "DANH MỤC TỪ VIẾT TẮT", "DANH MỤC CÁC KÝ HIỆU",
    "LỜI CẢM ƠN", "LỜI CAM ĐOAN", "LỜI MỞ ĐẦU", "TÓM TẮT",
}

RE_CHAPTER = re.compile(r"^\s*CHƯƠNG\s+(\d+|[IVXLCDM]+)\b", re.IGNORECASE)
RE_NUMBERED = re.compile(r"^\s*(\d+(?:\.\d+){0,3})\.?[\s:]+(\S.*)$")
RE_CAPTION = re.compile(r"^\s*(Hình|Bảng|Biểu đồ|Sơ đồ|Đồ thị)\s+\d+([.\-]\d+)?\s*[:.\-]", re.IGNORECASE)
RE_REF_ITEM = re.compile(r"^\s*\[\d+\]")


# ---------------------------------------------------------------- helpers ----
def set_run_font(run, name, size=None, bold=None, italic=None):
    f = run.font
    f.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), name)
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    f.color.rgb = BLACK  # chuan hoa: chu den, het mau xanh cua style Heading mac dinh


def apply_para(p, cfg, preset, *, heading=False, first_indent=True):
    pf = p.paragraph_format
    body = preset["body"]
    if heading:
        # numPr (danh so kieu danh sach) chi la "thua" khi doan DA co so GO TAY
        # trong van ban — luc do de ca hai se trung so. Neu van ban khong co so
        # go tay thi numPr la nguon so DUY NHAT (bai chuan dung kieu nay: text
        # chi la "Hien tuong truot dat.", so "1.1." do numPr sinh khi hien thi)
        # — go numPr la cat mat so muc cua bai.
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None:
            numPr = pPr.find(qn("w:numPr"))
            if numPr is not None and re.match(
                    r"^\s*(\d+(\.\d+)*|[IVX]{1,4})[.)]?\s", p.text or ""):
                pPr.remove(numPr)
    if "alignment" in cfg:
        p.alignment = cfg["alignment"]
    elif not heading:
        p.alignment = body["alignment"]
    # Tieu de phai di kem doan sau (khong nam tro troi o cuoi trang), than van
    # tranh de 1 dong le loi sang trang moi.
    pf.keep_with_next = bool(heading)
    pf.keep_together = bool(heading)
    pf.widow_control = True
    pf.line_spacing = cfg.get("line_spacing", body["line_spacing"])
    pf.space_before = Pt(cfg.get("space_before_pt", 0 if not heading else 6))
    pf.space_after = Pt(cfg.get("space_after_pt", body["space_after_pt"]))
    if "left_indent_cm" in cfg:
        pf.left_indent = Cm(cfg["left_indent_cm"])
    elif not cfg.get("hanging_indent_cm"):
        pf.left_indent = Cm(0)
    if cfg.get("hanging_indent_cm"):
        pf.left_indent = Cm(cfg["hanging_indent_cm"])
        pf.first_line_indent = Cm(-cfg["hanging_indent_cm"])
    elif first_indent and not heading:
        pf.first_line_indent = Cm(body["first_line_indent_cm"])
    else:
        pf.first_line_indent = Cm(0)
    for run in p.runs:
        set_run_font(run, preset["font_name"], cfg.get("size"),
                     cfg.get("bold"), cfg.get("italic"))


RE_DOT_LEADER = re.compile(r"[.．…]{4,}")
RE_ROMAN = re.compile(r"^(I{1,3}|IV|V|VI{0,3}|IX|X)\.\s+\S")
RE_ROMAN_PREFIX = re.compile(r"^(I{1,3}|IV|V|VI{0,3}|IX|X)\.\s+")


def is_leader_line(p):
    """Dong muc luc/danh muc go tay: hoac chua day cham that (.....), hoac dung TAB
    co leader cham (dau cham chi la cach hien thi cua tab, khong co trong van ban),
    hoac co tab + so trang o cuoi. Cac dong nay phai duoc bo qua hoan toan."""
    text = p.text
    if RE_DOT_LEADER.search(text):
        return True
    try:
        if 'leader="dot"' in p._p.xml or "leader='dot'" in p._p.xml:
            return True
    except Exception:
        pass
    if "\t" in text:
        head, _, tail = text.rpartition("\t")
        if head.strip() and len(head) <= 120 and tail.strip().rstrip(".").isdigit():
            return True
    return False


def classify(p):
    """Tra ve: chapter | h1 | h2 | h3 | h4 | caption | ref | body | skip"""
    style = (p.style.name or "")
    text = p.text.strip()
    if not text:
        return "body"
    if is_leader_line(p):
        return "skip"  # dong muc luc go tay — khong dung vao
    upper = text.upper()
    # Chuong/phan mo dau nhan dien theo NOI DUNG truoc style: nhieu tai lieu gan
    # style Heading 1 cho dong "Chuong 1" — van phai xu ly nhu tieu de chuong.
    if RE_CHAPTER.match(text) or upper in FRONT_MATTER or re.match(r"^\s*CHƯƠNG\s", upper):
        return "chapter"
    # Muc lon danh so La Ma (I. MO DAU, II. NOI DUNG, III. TAI LIEU THAM KHAO...)
    rm = RE_ROMAN.match(text)
    if rm and len(text) <= 120 and len(text.split()) <= 18:
        after = text[rm.end():].lstrip()
        if after[:1].isupper() and not text.rstrip().endswith((",", ";")):
            return "h1"
    if style.startswith("Heading 1") or style == "Title":
        return "h1"
    if style.startswith("Heading 2"):
        return "h2"
    if style.startswith("Heading 3"):
        return "h3"
    if RE_CAPTION.match(text) and len(text) < 200:
        return "caption"
    m = RE_NUMBERED.match(text)
    if m and len(text) <= 120 and len(text.split()) <= 24 and not text.rstrip().endswith((",", ";")):
        depth = m.group(1).count(".") + 1
        return {1: "h1", 2: "h2", 3: "h3", 4: "h4"}[depth]
    if RE_REF_ITEM.match(text):
        return "ref"
    return "body"


def add_page_number_footer(doc, preset):
    cfg = preset["page_number"]
    for section in doc.sections:
        # Ca 3 loai footer (thuong / trang dau / trang chan-le) — file goc hay co
        # footer rieng cho trang dau, chi sua footer thuong thi so trang cu van con.
        footers = [section.footer]
        try:
            footers += [section.first_page_footer, section.even_page_footer]
        except Exception:
            pass
        for footer in footers:
            footer.is_linked_to_previous = False
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            p.text = ""
            p.alignment = cfg["alignment"]
            run = p.add_run()
            set_run_font(run, preset["font_name"], cfg["size"])
            fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
            instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
            instr.text = "PAGE"
            fld_sep = OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
            fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
            run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_sep); run._r.append(fld_end)


def _field_paragraph(instr_text, placeholder):
    p = OxmlElement("w:p")
    def mk_run():
        return OxmlElement("w:r")
    r = mk_run()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin"); r.append(b); p.append(r)
    r = mk_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr_text
    r.append(it); p.append(r)
    r = mk_run()
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate"); r.append(s); p.append(r)
    r = mk_run()
    t = OxmlElement("w:t"); t.text = placeholder; r.append(t); p.append(r)
    r = mk_run()
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end"); r.append(e); p.append(r)
    return p


def _toc_title_para():
    title = OxmlElement("w:p")
    jc = OxmlElement("w:pPr"); al = OxmlElement("w:jc"); al.set(qn("w:val"), "center")
    jc.append(al); title.append(jc)
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rpr.append(OxmlElement("w:b"))
    sz = OxmlElement("w:sz"); sz.set(qn("w:val"), "32"); rpr.append(sz)
    r.append(rpr)
    t = OxmlElement("w:t"); t.text = "MỤC LỤC"; r.append(t); title.append(r)
    return title


def _page_break_para():
    brk = OxmlElement("w:p")
    r = OxmlElement("w:r"); br = OxmlElement("w:br"); br.set(qn("w:type"), "page")
    r.append(br); brk.append(r)
    return brk


def _classify_head(text):
    """Phan loai chi theo van ban cua MOT segment tieu de (dung cho doan gop
    tieu de + than van: classify() tren ca doan se bi tu choi vi qua dai)."""
    t = text.strip()
    if not t:
        return "body"
    if RE_CHAPTER.match(t) or re.match(r"^\s*CHƯƠNG\s", t.upper()):
        return "chapter"
    m = RE_NUMBERED.match(t)
    if m and len(t) <= 120 and len(t.split()) <= 24:
        depth = m.group(1).count(".") + 1
        return {1: "h1", 2: "h2", 3: "h3", 4: "h4"}[depth]
    return "body"


def _collect_toc_entries(doc, anchor_para):
    """Thu cac muc luc mau tu tieu de that cua bai (phan loai theo van ban) de
    lam 'ket qua tam' cho truong TOC — mo bang trinh xem nao cung thay muc luc,
    khong phai trang trang cho den khi Word cap nhat so trang."""
    entries = []
    paras = doc.paragraphs
    start = None
    for i, p in enumerate(paras):
        if p._p is anchor_para._p:
            start = i
            break
    if start is None:
        return entries
    level_map = {"chapter": 1, "h1": 1, "h2": 2, "h3": 3, "h4": 3}
    i = start
    while i < len(paras) and len(entries) < 120:
        p = paras[i]
        t = p.text.strip()
        i += 1
        if not t or is_leader_line(p):
            continue
        kind = classify(p)
        text = t
        if kind not in level_map and "\n" in t:
            # Doan gop 'tieu de + than van': thu phan loai rieng segment dau
            head = t.split("\n")[0].strip()
            hk = _classify_head(head)
            if hk in level_map:
                kind, text = hk, head
        if kind not in level_map:
            continue
        text = " ".join(seg.strip() for seg in text.split("\n") if seg.strip())
        if kind == "chapter" and len(text) <= 22 and i < len(paras):
            qt = " ".join(seg.strip()
                          for seg in paras[i].text.strip().split("\n") if seg.strip())
            caps = sum(1 for c in qt if c.isupper())
            letters = sum(1 for c in qt if c.isalpha())
            if (qt and len(qt) <= 90 and not qt.endswith(".")
                    and letters and caps >= 0.7 * letters):
                text = f"{text}: {qt}"
                i += 1
        entries.append((level_map[kind], text))
    return entries


def _toc_field_paras(doc, preset, anchor_para):
    """Cac doan cua truong TOC: mo truong + cac muc mau + dong huong dan + dong
    truong. Khong co muc mau thi dung ban mot doan nhu cu."""
    entries = _collect_toc_entries(doc, anchor_para)
    if not entries:
        return [_field_paragraph(' TOC \\o "1-3" \\h \\z \\u ',
                                 "(Nhấp chuột phải vào đây → Update Field để hiện mục lục)")]
    style_names = {s.name for s in doc.styles}
    for lvl in (1, 2, 3):
        if f"TOC {lvl}" not in style_names:
            doc.styles.add_style(f"TOC {lvl}", WD_STYLE_TYPE.PARAGRAPH)

    def mk_para(style_name, text=None, field_runs=()):
        p = OxmlElement("w:p")
        if style_name:
            pPr = OxmlElement("w:pPr")
            ps = OxmlElement("w:pStyle")
            ps.set(qn("w:val"), style_name.replace(" ", ""))
            pPr.append(ps)
            p.append(pPr)
        for fr in field_runs:
            p.append(fr)
        if text is not None:
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = text
            r.append(t)
            p.append(r)
        return p

    def fld(kind):
        r = OxmlElement("w:r")
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        r.append(f)
        return r

    def instr():
        r = OxmlElement("w:r")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = ' TOC \\o "1-3" \\h \\z \\u '
        r.append(it)
        return r

    paras = []
    lvl0, txt0 = entries[0]
    paras.append(mk_para(f"TOC {lvl0}", txt0,
                         (fld("begin"), instr(), fld("separate"))))
    for lvl, txt in entries[1:]:
        paras.append(mk_para(f"TOC {lvl}", txt))
    paras.append(mk_para(None, "(Trong Word: nhấp chuột phải vào mục lục → "
                               "Update Field để hiện số trang)"))
    paras.append(mk_para(None, None, (fld("end"),)))
    return paras


def insert_toc(doc, preset, anchor_para=None):
    """Chen trang MUC LUC tu dong ngay TRUOC doan neo noi dung (sau toan bo phan
    dau: bia, nhiem vu, loi cam on, ca khoi muc luc go tay neu co). Chi khi khong
    co doan neo moi dung phuong an du phong tim doan tieu de 'MUC LUC' co san —
    ban cu tim 'MUC LUC' truoc nen chen nham vao giua khoi muc luc go tay."""
    toc_field = _field_paragraph(' TOC \\o "1-3" \\h \\z \\u ',
                                 "(Nhấp chuột phải vào đây → Update Field để hiện mục lục)")
    if anchor_para is not None:
        anchor = anchor_para._p
        # addprevious chen ngay truoc doan neo, theo thu tu: tieu de -> truong -> ngat trang
        anchor.addprevious(_toc_title_para())
        for fp in _toc_field_paras(doc, preset, anchor_para):
            anchor.addprevious(fp)
        anchor.addprevious(_page_break_para())
        return True
    for p in doc.paragraphs:
        t = p.text.strip().upper()
        if t == "MỤC LỤC" or (t.startswith("MỤC LỤC") and len(t) <= 12):
            p._p.addnext(toc_field)
            return True
    if not doc.paragraphs:
        return False
    anchor = doc.paragraphs[0]._p
    anchor.addprevious(_toc_title_para())
    anchor.addprevious(toc_field)
    anchor.addprevious(_page_break_para())
    return True


_CITATION_STARTS = ("Luật", "Nghị định", "Thông tư", "Quyết định", "Chỉ thị",
                    "Công văn", "Hiến pháp")


def _dash_grouped(paras, idx):
    """Doan ke truoc/sau (bo qua doan trong) co phai muc liet ke bat dau bang
    gach ngang khong? Mot khoi liet ke dung gach la PHONG CACH cua bai (bai mau
    nhom 7 dung '- ' lam dau muc) — chi gach dinh o doan van xuoi DON DOC moi
    la loi go thua can xoa."""
    for step in (-1, 1):
        seen = 0
        j = idx + step
        while 0 <= j < len(paras) and seen < 2:
            t = paras[j].text.strip()
            if t:
                if len(t) >= 2 and t[0] in "-–—" and t[1] == " ":
                    return True
                break
            seen += 1
            j += step
    return False


def _strip_prose_dash(p, grouped=False):
    """Doan VAN XUOI dai bi go thua dau gach dau dong ('-Trong moi du an...',
    '- Hoc phan Nhap mon...') -> bo dau gach, giu nguyen noi dung. Chi ap dung khi
    doan dai (>=150 ky tu), ket thuc bang dau cham — cac muc LIET KE ngan va danh
    sach van ban phap luat (Luat/Nghi dinh/Thong tu...) giu nguyen kieu gach.
    Doan nam trong KHOI liet ke gach (grouped) thi giu nguyen dau gach."""
    if grouped:
        return False
    t = p.text.strip()
    if len(t) < 150 or not t.rstrip().endswith("."):
        return False
    if t[0] not in "-–—":
        return False
    if t[1:].lstrip().startswith(_CITATION_STARTS):
        return False
    for run in p.runs:
        if not run.text.strip():
            continue
        s = run.text
        i = 0
        while i < len(s) and s[i] in " \t":
            i += 1
        if i < len(s) and s[i] in "-–—":
            run.text = s[:i] + s[i + 1:].lstrip(" ")
            return True
        return False
    return False


def _looks_columnar(text):
    """Doan co nhieu khoang trang dai can cot (phieu nhiem vu 2 cot) — khong phai
    van xuoi, khong duoc ep dinh dang than van o vung dau."""
    return len(re.findall(r" {3,}", text)) >= 2


def enable_update_fields_on_open(doc):
    settings = doc.settings.element
    el = OxmlElement("w:updateFields")
    el.set(qn("w:val"), "true")
    settings.append(el)


# ------------------------------------------------------------------ main ----
def _split_at_first_break(p):
    """Tach doan tai dau ngat dong (w:br) dau tien: phan truoc tra ve trong doan goc,
    phan sau thanh doan moi chen ngay sau. Tra ve Paragraph phan sau hoac None."""
    import copy as _copy
    from docx.text.paragraph import Paragraph
    p_el = p._p
    children = [c for c in p_el if c.tag != qn("w:pPr")]
    for idx, child in enumerate(children):
        if child.tag != qn("w:r"):
            continue
        br = child.find(qn("w:br"))
        if br is None:
            continue
        new_el = OxmlElement("w:p")
        pPr = p_el.find(qn("w:pPr"))
        if pPr is not None:
            new_el.append(_copy.deepcopy(pPr))
        tail_run = OxmlElement("w:r")
        rPr = child.find(qn("w:rPr"))
        if rPr is not None:
            tail_run.append(_copy.deepcopy(rPr))
        seen = False
        for sub in list(child):
            if sub is br:
                seen = True
                child.remove(sub)
                continue
            if seen:
                tail_run.append(sub)
        if len(tail_run):
            new_el.append(tail_run)
        for sib in children[idx + 1:]:
            new_el.append(sib)
        p_el.addnext(new_el)
        return Paragraph(new_el, p._parent)
    return None


def _head_segment_is_heading(text, prev_was_chapter=False):
    """Doan dau tien (truoc dau ngat dong) co phai tieu de khong."""
    t = text.strip()
    if not t or len(t) > 120 or RE_DOT_LEADER.search(t):
        return False
    if RE_CHAPTER.match(t) or t.upper() in FRONT_MATTER:
        return True
    if prev_was_chapter and len(t) <= 90 and t == t.upper() and any(c.isalpha() for c in t):
        return True  # dong ten chuong viet hoa dung ngay sau dong "Chuong X"
    m = RE_NUMBERED.match(t)
    return bool(m and len(t.split()) <= 16)


def style_toc(doc, preset):
    """Chuan hoa style TOC 1/2/3 cua tai lieu de khi Word cap nhat muc luc, cac dong
    muc luc sat le trai, dung co chu/dam nhat theo mau (cap 1 dam, cac cap sau thuong).
    Bo qua am tham neu tai lieu khong dinh nghia san style TOC."""
    toc_cfgs = (("TOC 1", preset["h1"]["size"], True),
                ("TOC 2", preset["body"]["size"], False),
                ("TOC 3", preset["body"]["size"], False))
    for name, size, bold in toc_cfgs:
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        st.font.name = preset["font_name"]
        st.font.size = Pt(size)
        st.font.bold = bold
        st.font.italic = False
        rpr = st.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:eastAsia"), preset["font_name"])
        pf = st.paragraph_format
        pf.left_indent = Cm(0)
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = preset["body"]["line_spacing"]


def ensure_heading_styles(doc, preset):
    """Mot so file (xuat tu Google Docs/WPS/trinh soan thao truc tuyen) KHONG co san
    style Heading 1-4. Khong co chung, doan tieu de chi 'trong giong' tieu de nho
    dinh dang truc tiep, nhung Word khong coi la tieu de: muc luc bao 'No table of
    contents entries found', navigation trong xuat. Thieu thi tao moi, co kem
    outline level de Word/LibreOffice quet vao muc luc duoc."""
    from docx.enum.style import WD_STYLE_TYPE
    from docx.shared import RGBColor
    specs = (("Heading 1", 0, preset["h1"]), ("Heading 2", 1, preset["h2"]),
             ("Heading 3", 2, preset["h3"]), ("Heading 4", 3, preset["h4"]))
    for name, lvl, cfg in specs:
        try:
            doc.styles[name]
            continue
        except KeyError:
            pass
        st = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        try:
            st.base_style = doc.styles["Normal"]
        except KeyError:
            pass
        pPr = st.element.get_or_add_pPr()
        ol = OxmlElement("w:outlineLvl")
        ol.set(qn("w:val"), str(lvl))
        pPr.append(ol)
        st.font.name = preset["font_name"]
        st.font.size = Pt(cfg["size"])
        st.font.bold = bool(cfg.get("bold"))
        st.font.italic = bool(cfg.get("italic"))
        st.font.color.rgb = RGBColor(0, 0, 0)


_LINE_END_OK = (".", "!", "?", ":", ";", "…", '"', "”", ")", "]")


def _is_mergeable_line(p):
    """Mot DONG roi trong doan van bi xé nhỏ: dai vua, nhieu tu, khong phai tieu de /
    dong muc luc / muc liet ke / dong co tab hay ngat dong san."""
    t = p.text.strip()
    if len(t) < 35 or len(t.split()) < 6:
        return False
    if "\t" in p.text or "\n" in p.text or is_leader_line(p):
        return False
    if re.search(r" {3,}", p.text):
        return False
    if t[0] in "-–—•*":
        return False
    if p._p.xpath(".//a:blip"):
        return False
    return classify(p) == "body"


def _is_mergeable_tail(p):
    """Dong CUOI cua mot doan bi xe nho thuong rat ngan ('tap va phat trien.') —
    chi can khong phai tieu de / dong muc luc / muc liet ke / dong co cau truc."""
    t = p.text.strip()
    if not t:
        return False
    if "\t" in p.text or "\n" in p.text or is_leader_line(p):
        return False
    if re.search(r" {3,}", p.text):
        return False
    if t[0] in "-–—•*":
        return False
    if p._p.xpath(".//a:blip"):
        return False
    return classify(p) == "body"


def _merge_shredded_lines(doc, protected=frozenset()):
    """Gop cac dong roi thanh doan van hoan chinh. Nhieu bai sinh vien dan van ban
    (tu PDF/Word khac) bi xe moi dong thanh mot 'doan' rieng — dinh dang tung dong
    se tao khoang cach doan khap noi va loi dam/nghieng tung dong khong sua duoc.
    Nguyen tac gop: dong truoc chua ket cau (khong cham/hoi/cham than/hai cham/
    cham phay/dong ngoac) thi noi dong sau vao, cach nhau mot dau cach."""
    merged = 0
    paras = doc.paragraphs
    i = 0
    while i < len(paras) - 1:
        p, q = paras[i], paras[i + 1]
        t = p.text.strip()
        if (p._p not in protected and q._p not in protected
                and _is_mergeable_line(p) and _is_mergeable_tail(q)
                and not t.endswith(_LINE_END_OK)):
            # Noi mot dau cach vao cuoi doan goc neu chua co
            ts = [n for n in p._p.iter() if n.tag == qn("w:t")]
            if ts and ts[-1].text and not ts[-1].text.endswith(" "):
                ts[-1].text += " "
            for child in list(q._p):
                if child.tag != qn("w:pPr"):
                    q._p.remove(child)
                    p._p.append(child)
            q._p.getparent().remove(q._p)
            merged += 1
            paras = doc.paragraphs
        else:
            i += 1
    return merged


def _split_columns(line):
    """Tach mot dong 2 cot viet bang dau cach thanh (cot trai, cot phai)."""
    parts = re.split(r" {3,}", line.strip())
    parts = [x.strip() for x in parts if x.strip()]
    if len(parts) >= 2:
        return parts[0], " ".join(parts[1:])
    return None


def _rebuild_quoc_hieu(doc):
    """Khoi quoc hieu 2 cot viet bang dau cach (TRUONG... | CONG HOA...) rat mong
    manh: doi font/do le la cot phai vo xuong dong. Dung lai bang 2 cot khong
    vien thi moi trinh doc deu hien giong nhau. Tra ve True neu da dung lai."""
    # Quoc hieu da la BANG san (tai lieu hoan thien) thi giu nguyen, khong dung lai.
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if "CỘNG HÒA" in cell.text.upper():
                    return False
    for i, p in enumerate(doc.paragraphs[:4]):
        if "CỘNG HÒA" not in p.text.upper():
            continue
        lefts, rights = [], []
        for seg in p.text.split("\n"):
            if not seg.strip():
                continue
            cols = _split_columns(seg)
            if cols is None:
                return False
            lefts.append(cols[0])
            rights.append(cols[1])
        if len(lefts) < 2:
            return False
        # Dong "KHOA ..." dung rieng o doan lien sau cung thuoc cot trai
        extra = None
        if i + 1 < len(doc.paragraphs):
            q = doc.paragraphs[i + 1]
            qt = q.text.strip()
            if (qt and len(qt) <= 45 and "\n" not in q.text
                    and not re.search(r" {3,}", qt)
                    and sum(1 for c in qt if c.isupper()) >= 0.6 * max(1, sum(1 for c in qt if c.isalpha()))):
                extra = qt
        table = doc.add_table(rows=1, cols=2)
        tbl = table._tbl
        p._p.addprevious(tbl)
        # Khong vien bang + tat autofit, cot 8/8cm (vung chu rong 16cm)
        tblPr = tbl.find(qn("w:tblPr"))
        borders = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "none")
            borders.append(el)
        tblPr.append(borders)
        for idx_w, w in enumerate((8.0, 8.0)):
            table.columns[idx_w].width = Cm(w)
        if extra:
            lefts.append(extra)
        cell_l, cell_r = table.rows[0].cells
        for cell, lines, center, bold_last in (
                (cell_l, lefts, False, True), (cell_r, rights, True, False)):
            for j, line in enumerate(lines):
                para = cell.paragraphs[0] if j == 0 else cell.add_paragraph()
                para.alignment = (WD_ALIGN_PARAGRAPH.CENTER if center
                                  else WD_ALIGN_PARAGRAPH.LEFT)
                para.paragraph_format.space_after = Pt(0)
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.line_spacing = 1.15
                run = para.add_run(line)
                bold = center or (bold_last and j == len(lines) - 1)
                set_run_font(run, "Times New Roman", 13.0 if not center else 13.5,
                             bold=bold)
        p._p.getparent().remove(p._p)
        if extra:
            q._p.getparent().remove(q._p)
        return True
    return False


def _spaces_to_tabs_in_para(p, stop_cm=9.0):
    """Doi cac khoang dau cach dai O GIUA dong (canh cot gia) thanh ky tu tab that
    + dat mot moc tab co dinh: cot khong con vo khi doi font/may doc."""
    changed = False
    for r in list(p._p.findall(qn("w:r"))):
        for t in list(r.findall(qn("w:t"))):
            txt = t.text or ""
            if not re.search(r"\S {3,}\S", txt):
                continue
            parts = re.split(r"(?<=\S) {3,}(?=\S)", txt)
            if len(parts) < 2:
                continue
            changed = True
            t.text = parts[0]
            anchor = t
            for part in parts[1:]:
                tab = OxmlElement("w:tab")
                anchor.addnext(tab)
                nt = OxmlElement("w:t")
                nt.set(qn("xml:space"), "preserve")
                nt.text = part
                tab.addnext(nt)
                anchor = nt
    if changed:
        p.paragraph_format.tab_stops.add_tab_stop(Cm(stop_cm))
    return changed


def _repair_front_matter(doc, content_start):
    """Sua cac khoi phan dau viet bang dau cach: quoc hieu (bang 2 cot) va cac
    dong bieu mau canh cot (ten | MSSV...) doi sang tab stop co dinh."""
    n_tab = 0
    for i, p in enumerate(doc.paragraphs):
        if i >= content_start:
            break
        if is_leader_line(p):
            continue
        if re.search(r" {3,}", p.text) and _spaces_to_tabs_in_para(p):
            n_tab += 1
    return n_tab


def _replace_hand_toc(doc, preset, content_start):
    """Thay khoi muc luc GO TAY (dong cham dan, so trang sai cu) bang muc luc
    TU DONG dung tieu de that cua bai. Tra ve True neu da thay."""
    leader_idx = [i for i, p in enumerate(doc.paragraphs)
                  if re.search(r"[.．…]{4,}", p.text)]
    if len(leader_idx) < 3 or content_start is None:
        return False
    first = leader_idx[0]
    if first >= content_start:
        return False
    anchor_el = doc.paragraphs[content_start]._p
    # Tieu de muc luc co the la segment cuoi cua doan lien truoc cum
    if first > 0:
        prev = doc.paragraphs[first - 1]
        segs = prev.text.split("\n")
        if segs and segs[-1].strip().upper() in ("MỤC LỤC", "PHỤ LỤC"):
            _drop_last_segment(prev)
    # Xoa toan bo doan trong khoi [first, content_start)
    for p in list(doc.paragraphs[first:content_start]):
        p._p.getparent().remove(p._p)
    anchor = None
    for p in doc.paragraphs:
        if p._p is anchor_el:
            anchor = p
            break
    if anchor is None:
        return False
    ok = insert_toc(doc, preset, anchor)
    if ok:
        style_toc(doc, preset)
    return ok


def _drop_last_segment(p):
    """Bo segment cuoi (sau w:br cuoi) cua mot doan gop — dung khi segment do la
    tieu de MUC LUC/PHU LUC dinh vao doan truoc khoi muc luc tay."""
    last_br = None
    for child in p._p:
        if child.tag == qn("w:r") and child.find(qn("w:br")) is not None:
            last_br = child
    if last_br is None:
        # Khong co br: ca doan la tieu de — xoa sach run
        for child in list(p._p):
            if child.tag != qn("w:pPr"):
                p._p.remove(child)
        return
    seen = False
    for child in list(p._p):
        if child is last_br:
            seen = True
            p._p.remove(child)
            continue
        if seen:
            p._p.remove(child)


_TITLE_PHRASES = ("NHIỆM VỤ", "MỤC LỤC", "PHỤ LỤC", "LỜI MỞ ĐẦU", "LỜI CẢM ƠN",
                  "LỜI GIỚI THIỆU", "DANH MỤC")


def _cap_front_fonts(doc, content_start, cap_pt=13.0):
    """Dong bieu mau o phan dau thuong bi go 16pt dam nghieng khac han chuan 13
    cua bai — ha co chu xuong 13 cho dong bo, giu nguyen dam/nghieng/bocuc.
    Bo qua: dong muc luc, dong tieu de lon (NHIEM VU / MUC LUC / LOI MO DAU...)."""
    n = 0
    for i, p in enumerate(doc.paragraphs):
        if i >= content_start:
            break
        t = p.text.strip()
        if not t or is_leader_line(p):
            continue
        if any(ph in t.upper() for ph in _TITLE_PHRASES):
            continue
        changed = False
        for run in p.runs:
            sz = run.font.size
            if sz is not None and sz.pt > cap_pt:
                run.font.size = Pt(cap_pt)
                changed = True
        if changed:
            n += 1
    return n


def _fix_glued_case(doc, protected=frozenset()):
    """Sua chu dinh do dan van ban: 'ChíMinh' -> 'Chí Minh' (ranh gioi chu
    thuong -> chu hoa o giua tu). Chi ap doan than van, khong dung tieu de /
    dong muc luc. Chu dinh ca hai phia thuong (trungphân) khong tu sua duoc."""
    n = 0
    _lower = "a-zàáâãèéêìíòóôõùúăđĩũơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹ"
    _upper = "A-ZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠƯẠẢẤẦẨẪẬẮẰẲẴẶẸẺẼẾỀỂỄỆỈỊỌỎỐỒỔỖỘỚỜỞỠỢỤỦỨỪỬỮỰỲỴỶỸ"
    pat = re.compile(f"(?<=[{_lower}])(?=[{_upper}])")
    for p in doc.paragraphs:
        if p._p in protected:
            continue
        if p.style.name.startswith("Heading") or is_leader_line(p):
            continue
        if classify(p) != "body":
            continue
        for t in p._p.iter(qn("w:t")):
            if t.text and pat.search(t.text):
                t.text, k = pat.subn(" ", t.text)
                n += k
    return n


def _find_toc_field(doc):
    """Tim field TOC that (Word tu sinh) trong tai lieu. Tra ve (set element cac
    doan thuoc vung field, chi so doan cuoi cua field). Tai lieu da co muc luc
    tu dong that la tai lieu DA hoan thien: dot leader cua no do TAB dan sinh ra
    (khong co ky tu cham trong van ban) nen cac co che nhan dien muc luc go tay
    khong thay — vong lap chinh tung format lai cac dong muc luc nay nhu van
    thuong, pha hong muc luc dep san co (bai hoc tu bai mau nhom 7, 26/10)."""
    protected = set()
    end_idx = -1
    # TOC duoc Word boc trong khoi sdt (gallery) — cac doan nam ben trong sdt,
    # khong xuat hien trong doc.paragraphs nen phai quet rieng o muc body.
    from lxml import etree
    for child in doc.element.body:
        if child.tag == qn("w:sdt") and re.search(
                r"<w:instrText[^>]*>[^<]*\bTOC\b",
                etree.tostring(child, encoding="unicode")):
            for pel in child.iter(qn("w:p")):
                protected.add(pel)
    if protected:
        return protected, end_idx
    depth = 0
    in_toc = False
    toc_depth = 0
    paras = doc.paragraphs
    for i, p in enumerate(paras):
        xml = p._p.xml
        begins = xml.count('fldCharType="begin"')
        ends = xml.count('fldCharType="end"')
        has_toc_instr = bool(re.search(r"<w:instrText[^>]*>[^<]*\bTOC\b", xml))
        if has_toc_instr and not in_toc:
            in_toc = True
            toc_depth = depth
            if i > 0 and paras[i - 1].text.strip().upper() == "MỤC LỤC":
                protected.add(paras[i - 1]._p)
        if in_toc:
            protected.add(p._p)
            end_idx = i
        depth += begins - ends
        if in_toc and ends and depth <= toc_depth:
            in_toc = False
    return protected, end_idx


def _has_page_field(doc):
    """Tai lieu da co field so trang (PAGE/NUMPAGES) trong footer chua?"""
    for section in doc.sections:
        for part in (section.footer, section.first_page_footer,
                     section.even_page_footer):
            try:
                if "PAGE" in part._element.xml:
                    return True
            except Exception:
                continue
    return False


def format_document(input_path, output_path, preset="chuan_vn",
                    add_page_numbers=True, add_toc=True):
    if isinstance(preset, str):
        preset = PRESETS[preset]
    doc = Document(input_path)
    toc_protected, _toc_end = _find_toc_field(doc)
    polished = bool(toc_protected)
    n_merged = _merge_shredded_lines(doc, protected=toc_protected)
    n_glued = _fix_glued_case(doc, protected=toc_protected)
    ensure_heading_styles(doc, preset)
    quoc_hieu_ok = _rebuild_quoc_hieu(doc)
    report = {"preset": preset["label"], "chuong": 0, "muc_cap_1": 0,
              "muc_cap_2": 0, "muc_cap_3": 0, "muc_cap_4": 0, "chu_thich": 0,
              "doan_van": 0, "doan_bo_qua": 0, "gach_da_xoa": 0, "doan_da_gop": 0,
              "bang": 0, "so_trang_da_them": False, "muc_luc_da_chen": False,
              "muc_luc_tay_co_san": False, "dong_bieu_mau_da_sua": 0,
              "quoc_hieu_da_dung_lai": False, "chu_dinh_da_sua": 0}
    report["doan_da_gop"] = n_merged
    report["chu_dinh_da_sua"] = n_glued

    report["muc_luc_tu_dong_co_san"] = polished
    page = preset["page"]
    # Tai lieu nhieu section thuong da tach le co chu dich (vd phan dau 2,54 /
    # noi dung 2,5-2-3-2 nhu bai mau nhom 7) — ep mot le chung se pha hinh hoc
    # trang dau. Chi chuan hoa le khi tai lieu co dung 1 section.
    if len(doc.sections) == 1:
        for s in doc.sections:
            s.page_width = Cm(page["width_cm"]); s.page_height = Cm(page["height_cm"])
            s.top_margin = Cm(page["top_cm"]); s.bottom_margin = Cm(page["bottom_cm"])
            s.left_margin = Cm(page["left_cm"]); s.right_margin = Cm(page["right_cm"])

    # Vung bao ve phan dau: moi doan TRUOC diem bat dau noi dung that su (bia, nhiem vu,
    # loi cam on, muc luc go tay...) duoc GIU NGUYEN, khong ep dinh dang. Diem bat dau
    # la doan chuong / muc La Ma / tieu de phan mo dau dau tien KHONG PHAI dong muc luc
    # va PHAI co doan than van theo sau (dong tieu de trong khoi muc luc go tay trong
    # y het tieu de that, chi khac o cho sau chung khong co than van).
    def _followed_by_body(idx, window=8):
        seen = 0
        for q in doc.paragraphs[idx + 1:]:
            qt = q.text.strip()
            if not qt:
                continue
            seen += 1
            if seen > window:
                break
            if is_leader_line(q):
                continue
            if len(qt) > 60 and classify(q) == "body":
                return True
        return False

    content_start = 0
    # Truong hop co khoi MUC LUC go tay: cac dong trong khoi trong y het tieu de
    # that (ke ca dong cuoi khoi nam sat noi dung) nen khong tin tung dong duoc.
    # Neo dung: tieu de dung ngay truoc doan THAN VAN dau tien sau khoi muc luc.
    ml_idx = None
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if t and (t.upper() == "MỤC LỤC" or t.upper().startswith("MỤC LỤC\n")):
            ml_idx = i
            break
    if ml_idx is not None:
        body_idx = None
        for j in range(ml_idx + 1, len(doc.paragraphs)):
            q = doc.paragraphs[j]
            qt = q.text.strip()
            if qt and len(qt) > 60 and not is_leader_line(q) and classify(q) == "body":
                body_idx = j
                break
        if body_idx is not None:
            for k in range(body_idx - 1, ml_idx, -1):
                q = doc.paragraphs[k]
                qt = q.text.strip()
                if qt and not is_leader_line(q) and classify(q) in ("chapter", "h1"):
                    content_start = k
                    break
    first_candidate = None
    for i, p in enumerate(doc.paragraphs):
        if content_start:
            break  # da co neo tu khoi muc luc go tay — neo do dang tin hon
        t = p.text.strip()
        if not t or is_leader_line(p):
            continue
        if (RE_CHAPTER.match(t) or RE_ROMAN.match(t)
                or t.upper() in ("MỞ ĐẦU", "LỜI MỞ ĐẦU", "PHẦN MỞ ĐẦU")):
            if first_candidate is None:
                first_candidate = i
            if _followed_by_body(i):
                content_start = i
                break
    else:
        if first_candidate is not None:
            content_start = first_candidate

    # Chuan hoa co chu phan dau: cac dong bieu mau go 16pt -> 13pt cho dong bo
    # (giu nguyen dam/nghieng/canh cot; bo qua dong tieu de lon)
    report["dong_bieu_mau_da_sua"] = _cap_front_fonts(doc, content_start)
    report["quoc_hieu_da_dung_lai"] = quoc_hieu_ok
    # Thay khoi muc luc go tay bang muc luc tu dong ngay tai vi tri cua no
    toc_replaced = False
    if add_toc and content_start and not polished:
        anchor_el = doc.paragraphs[content_start]._p
        toc_replaced = _replace_hand_toc(doc, preset, content_start)
        if toc_replaced:
            report["muc_luc_tay_co_san"] = True
            report["muc_luc_da_chen"] = True
            for j, q in enumerate(doc.paragraphs):
                if q._p is anchor_el:
                    content_start = j
                    break

    in_refs = False
    after_heading = False
    prev_was_chapter = False
    first_para_free = bool(preset["body"].get("first_para_no_indent"))
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if p._p in toc_protected:
            report["doan_bo_qua"] += 1
            prev_was_chapter = False
            continue
        if idx < content_start:
            if polished:
                # Tai lieu da hoan thien: phan dau giu nguyen tung doan.
                if text:
                    report["doan_bo_qua"] += 1
                prev_was_chapter = False
                continue
            # Vung dau: bia / phieu nhiem vu / dong cau truc ngan duoc GIU NGUYEN,
            # nhung doan VAN XUOI (loi cam on, loi gioi thieu...) bi canh giua, dinh
            # gach dau dong trong ban thao thi PHAI chuan hoa nhu than van — bao ve
            # ca khoi la ly do ban cu de nguyen cac loi nay.
            handled = False
            if (text and len(text) > 60 and not is_leader_line(p)
                    and "\t" not in p.text and not re.search(r" {3,}", p.text)):
                tail = None
                if "\n" in p.text:
                    head = p.text.split("\n")[0].strip()
                    if len(head) <= 40 and len(text) - len(head) > 60:
                        tail = _split_at_first_break(p)
                targets = []
                if tail is not None:
                    targets = [tail]  # doan goc gio la tieu de ngan — giu nguyen
                    handled = True
                elif ("\t" not in p.text and not re.search(r" {3,}", p.text)
                        and classify(p) == "body"):
                    targets = [p]
                for q in targets:
                    if not q.text.strip():
                        continue
                    apply_para(q, preset["body"], preset, first_indent=True)
                    for run in q.runs:
                        run.font.bold = False
                        run.font.italic = False
                    if _strip_prose_dash(q, grouped=_dash_grouped(doc.paragraphs, idx)):
                        report["gach_da_xoa"] += 1
                    report["doan_van"] += 1
                    handled = True
            if not handled and text:
                report["doan_bo_qua"] += 1
            prev_was_chapter = False
            continue
        kind = classify(p)
        if kind == "skip":
            report["doan_bo_qua"] += 1
            continue
        # Doan bi gon: tieu de + noi dung nam chung mot doan, cach nhau bang ngat dong
        # -> tach ra de tieu de duoc dinh dang dung cap, phan con lai thanh than van.
        tail = None
        if "\n" in p.text and _head_segment_is_heading(p.text.split("\n")[0], prev_was_chapter):
            tail = _split_at_first_break(p)
            if tail is not None:
                text = p.text.strip()
                kind = classify(p)
        if RE_ROMAN_PREFIX.sub("", text).strip().upper() == "TÀI LIỆU THAM KHẢO":
            in_refs = True
        # Dong ten chuong viet rieng ngay sau dong "Chuong X" (kieu file mau IUH)
        if (prev_was_chapter and kind == "body" and text and len(text) <= 90
                and text == text.upper() and any(c.isalpha() for c in text)):
            kind = "chapter_title"
        if kind in ("chapter", "chapter_title"):
            try:
                p.style = doc.styles["Heading 1"]
            except KeyError:
                pass
            apply_para(p, preset["chapter"], preset, heading=True)
            for run in p.runs:
                run.font.all_caps = bool(preset["chapter"].get("caps", True))
                run.font.bold = True
                run.font.italic = False
            if kind == "chapter":
                report["chuong"] += 1
            prev_was_chapter = (kind == "chapter")
            after_heading = True
        elif kind in ("h1", "h2", "h3", "h4"):
            style_map = {"h1": "Heading 1", "h2": "Heading 2",
                         "h3": "Heading 3", "h4": "Heading 4"}
            try:
                p.style = doc.styles[style_map[kind]]
            except KeyError:
                pass
            apply_para(p, preset[kind], preset, heading=True)
            report[{"h1": "muc_cap_1", "h2": "muc_cap_2",
                    "h3": "muc_cap_3", "h4": "muc_cap_4"}[kind]] += 1
            prev_was_chapter = False
            after_heading = True
        elif kind == "caption":
            apply_para(p, preset["caption"], preset, heading=True)
            report["chu_thich"] += 1
            prev_was_chapter = False
        elif kind == "ref" or in_refs:
            apply_para(p, preset["reference"], preset, first_indent=False)
            for run in p.runs:
                run.font.bold = False
                run.font.italic = False
            report["doan_van"] += 1
            prev_was_chapter = False
            after_heading = False
        else:
            has_image = bool(p._p.xpath(".//a:blip"))
            if has_image:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                for run in p.runs:
                    run.font.bold = False
                    run.font.italic = False
            else:
                # Quy dinh IUH: doan dau tien cua moi muc KHONG thut dau dong
                no_indent = first_para_free and after_heading and bool(text)
                apply_para(p, preset["body"], preset, first_indent=not no_indent)
                if "\n" in p.text:
                    # Doan chua ngat dong tay (w:br): can trai de cac dong khong
                    # bi keo gian chu khi can deu 2 ben.
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                # Than van chuan la chu THUONG: xoa dam/nghieng thua tu ban thao
                for run in p.runs:
                    run.font.bold = False
                    run.font.italic = False
                if _strip_prose_dash(p, grouped=_dash_grouped(doc.paragraphs, idx)):
                    report["gach_da_xoa"] += 1
            if text:
                after_heading = False
            prev_was_chapter = False
            report["doan_van"] += 1
        # Phan duoi cua doan bi gon (da tach o tren): dinh dang nhu than van thuong,
        # va vi no la doan dau ngay sau tieu de nen ap dung quy tac khong thut dau dong.
        # NGOAI LE: doan chuong dang "Chuong n\nTEN CHUONG" — phan duoi la TEN
        # chuong, phai dinh dang chuong (dam, in hoa, canh giua), khong phai than van.
        if tail is not None and tail.text.strip():
            if kind in ("chapter", "chapter_title"):
                apply_para(tail, preset["chapter"], preset, heading=True)
                for run in tail.runs:
                    run.font.all_caps = bool(preset["chapter"].get("caps", True))
                    run.font.bold = True
                    run.font.italic = False
                after_heading = True
                prev_was_chapter = False
            else:
                no_indent = first_para_free and after_heading
                apply_para(tail, preset["body"], preset, first_indent=not no_indent)
                if "\n" in tail.text:
                    tail.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in tail.runs:
                    run.font.bold = False
                    run.font.italic = False
                report["doan_van"] += 1
                after_heading = False
                prev_was_chapter = False

    for table in doc.tables:
        report["bang"] += 1
        if polished:
            continue  # bang trong tai lieu hoan thien: giu nguyen dinh dang goc
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = preset["table_text"]["line_spacing"]
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        set_run_font(run, preset["font_name"],
                                     preset["table_text"]["size"])

    if add_page_numbers:
        if _has_page_field(doc):
            report["so_trang_co_san"] = True
        else:
            add_page_number_footer(doc, preset)
            report["so_trang_da_them"] = True
    if add_toc and not toc_replaced and not polished:
        # Bai da co san muc luc tay (cac dong co CHAM DAN trong van ban) thi
        # KHONG chen them muc luc tu dong — tranh trung lap 2 muc luc lien nhau.
        # (Chi dem dong cham that trong text; dong bieu mau co tab + so o cuoi
        # khong tinh — is_leader_line bat ca nhung dong do.)
        n_leader = sum(1 for p in doc.paragraphs
                       if re.search(r"[.．…]{4,}", p.text))
        report["muc_luc_tay_co_san"] = n_leader >= 3
        if not report["muc_luc_tay_co_san"]:
            anchor = doc.paragraphs[content_start] if 0 < content_start < len(doc.paragraphs) else None
            report["muc_luc_da_chen"] = insert_toc(doc, preset, anchor)
            if report["muc_luc_da_chen"]:
                style_toc(doc, preset)
    enable_update_fields_on_open(doc)

    doc.save(output_path)
    return report


# ------------------------------------------------- trich xuat mau dinh dang ----
def _common(values):
    from collections import Counter
    vals = [v for v in values if v is not None]
    return Counter(vals).most_common(1)[0][0] if vals else None


def _style_cfg(style, fallback):
    """Doc cau hinh tu mot style doan van (Heading/Normal) cua file mau."""
    cfg = dict(fallback)
    if style is None:
        return cfg
    f, pf = style.font, style.paragraph_format
    if f.size is not None:
        cfg["size"] = f.size.pt
    if f.bold is not None:
        cfg["bold"] = f.bold
    if f.italic is not None:
        cfg["italic"] = f.italic
    if pf.alignment is not None:
        cfg["alignment"] = pf.alignment
    if pf.left_indent is not None:
        cfg["left_indent_cm"] = round(pf.left_indent.cm, 1)
    if pf.space_before is not None:
        cfg["space_before_pt"] = pf.space_before.pt
    if pf.space_after is not None:
        cfg["space_after_pt"] = pf.space_after.pt
    if pf.line_spacing is not None and isinstance(pf.line_spacing, float):
        cfg["line_spacing"] = pf.line_spacing
    return cfg


def extract_preset_from_sample(sample_path):
    """Doc file MAU dinh dang (cua truong/giang vien) va rut ra bo quy tac tuong duong.

    Tra ve (preset_dict, summary) — summary de hien cho nguoi dung xem da hoc duoc gi.
    Moi thong tin thieu deu roi ve gia tri cua preset chuan, khong bao gio loi.
    """
    import copy
    preset = copy.deepcopy(PRESETS["chuan_vn"])
    preset["label"] = "Mẫu tùy chỉnh (từ file mẫu của bạn)"
    preset["desc"] = "Bộ quy tắc được trích xuất tự động từ file mẫu bạn tải lên."
    sdoc = Document(sample_path)
    learned = []

    # 1) Kho giay + can le tu section dau tien cua file mau
    try:
        s = sdoc.sections[0]
        pg = preset["page"]
        for attr, key in (("page_width", "width_cm"), ("page_height", "height_cm"),
                          ("top_margin", "top_cm"), ("bottom_margin", "bottom_cm"),
                          ("left_margin", "left_cm"), ("right_margin", "right_cm")):
            v = getattr(s, attr, None)
            if v is not None:
                pg[key] = round(v.cm, 1)
        learned.append(f"lề {pg['top_cm']}/{pg['bottom_cm']}/{pg['left_cm']}/{pg['right_cm']} cm")
    except Exception:
        pass

    # 2) Than van: STYLE Normal la nguon dang tin nhat — file mau chuan thuong dinh
    #    nghia co chu o style, cac run ke thua nen khong ghi co chu truc tiep.
    #    Chi quet thong ke o VUNG NOI DUNG (tu chuong dau tien), tranh bia/nhiem vu.
    paras = list(sdoc.paragraphs)
    start = 0
    for i, p in enumerate(paras):
        t = p.text.strip()
        if t and RE_CHAPTER.match(t) and not is_leader_line(p):
            start = i
            break
    content_paras = paras[start:]
    body_paras = [p for p in content_paras
                  if len(p.text.strip()) > 40 and classify(p) == "body"]
    names = [r.font.name for p in body_paras for r in p.runs if r.font.name]
    sizes = [r.font.size.pt for p in body_paras for r in p.runs if r.font.size]
    spacings = [p.paragraph_format.line_spacing for p in body_paras
                if isinstance(p.paragraph_format.line_spacing, float)]
    try:
        normal = sdoc.styles["Normal"]
    except KeyError:
        normal = None
    # Chon nguon dang tin hon giua STYLE va dinh dang truc tiep tren run:
    # - File mau dung style chuan (run ke thua, khong ghi co chu) -> style thang.
    # - File mau dinh dang truc tiep (da so run ghi ro co chu) -> gia tri pho bien thang.
    total_runs = sum(len(p.runs) for p in body_paras)
    explicit_ratio = (len(sizes) / total_runs) if total_runs else 0.0
    style_name = normal.font.name if normal else None
    if style_name and not style_name.startswith(".") and explicit_ratio < 0.5:
        preset["font_name"] = style_name  # bo qua font legacy kieu .VnTime
    elif _common(names):
        preset["font_name"] = _common(names)
    elif style_name and not style_name.startswith("."):
        preset["font_name"] = style_name
    body = preset["body"]
    if explicit_ratio >= 0.5 and _common(sizes):
        body["size"] = _common(sizes)
    elif normal is not None and normal.font.size is not None:
        body["size"] = normal.font.size.pt
    elif _common(sizes):
        body["size"] = _common(sizes)
    nls = normal.paragraph_format.line_spacing if normal else None
    modal_spacing = _common(spacings)
    explicit_ls = sum(1 for p in body_paras
                      if isinstance(p.paragraph_format.line_spacing, float))
    if modal_spacing and body_paras and explicit_ls >= max(3, len(body_paras) // 2):
        body["line_spacing"] = modal_spacing
    elif isinstance(nls, float):
        body["line_spacing"] = nls
    elif modal_spacing:
        body["line_spacing"] = modal_spacing
    if normal is not None:
        if normal.paragraph_format.space_before is not None:
            body["space_before_pt"] = normal.paragraph_format.space_before.pt
        if normal.paragraph_format.space_after is not None:
            body["space_after_pt"] = normal.paragraph_format.space_after.pt
    # Khoang cach doan: lay gia tri pho bien nhat tu chinh cac doan than van cua mau
    # (dinh dang truc tiep tren doan la bang chung ro nhat, de len tren style)
    from collections import Counter

    def _modal_pt(vals):
        if not vals:
            return None
        return Counter(v if v is not None else 0.0 for v in vals).most_common(1)[0][0]

    sb = _modal_pt([p.paragraph_format.space_before.pt if p.paragraph_format.space_before else None
                    for p in body_paras])
    sa = _modal_pt([p.paragraph_format.space_after.pt if p.paragraph_format.space_after else None
                    for p in body_paras])
    if sb is not None:
        body["space_before_pt"] = sb
    if sa is not None:
        body["space_after_pt"] = sa
    # Phat hien quy tac "doan dau moi muc khong thut, doan sau thut X cm" cua file mau
    headings_seen = 0
    for j, p in enumerate(content_paras):
        if classify(p) in ("chapter", "h1", "h2", "h3", "h4"):
            headings_seen += 1
            nxt = [q for q in content_paras[j + 1:j + 8]
                   if len(q.text.strip()) > 40 and classify(q) == "body"]
            if len(nxt) >= 2:
                i0 = nxt[0].paragraph_format.first_line_indent
                i1 = nxt[1].paragraph_format.first_line_indent
                v0 = i0.cm if i0 else 0.0
                v1 = i1.cm if i1 else 0.0
                if v0 == 0.0 and v1 > 0:
                    body["first_para_no_indent"] = True
                    body["first_line_indent_cm"] = round(v1, 1)
                    break
                if v1 > 0:
                    body["first_line_indent_cm"] = round(v1, 1)
                    break
            if headings_seen >= 5:
                break
    learned.append(f"font {preset['font_name']} {body['size']}pt, giãn dòng {body['line_spacing']}")

    # 3) Tieu de: chuong doc tu doan CHUONG that; cac cap muc doc tu chinh cac doan
    #    danh so trong file mau (file mau chuan thuong dinh dang truc tiep, khong dung
    #    style Heading) — style Heading chi la nguon du phong.
    def get_style(name):
        try:
            return sdoc.styles[name]
        except KeyError:
            return None

    chapter_para = next((p for p in content_paras
                         if RE_CHAPTER.match(p.text.strip()) and p.runs
                         and not is_leader_line(p)), None)
    if chapter_para is not None:
        r = chapter_para.runs[0]
        ch = preset["chapter"]
        if r.font.size:
            ch["size"] = r.font.size.pt
        ch["bold"] = bool(r.font.bold)
        if chapter_para.alignment is not None:
            ch["alignment"] = chapter_para.alignment
        learned.append(f"chương {ch['size']}pt")
    else:
        preset["chapter"] = _style_cfg(get_style("Heading 1"), preset["chapter"])
        preset["chapter"].setdefault("caps", True)
    preset["h1"] = _style_cfg(get_style("Heading 1"), preset["h1"])

    def numbered_para(depth):
        for p in content_paras:
            t = p.text.strip()
            if not t or is_leader_line(p):
                continue
            m = RE_NUMBERED.match(t)
            if m and m.group(1).count(".") + 1 == depth and len(t) <= 120:
                return p
        return None

    for depth, slot, style_name_h in ((2, "h2", "Heading 2"), (3, "h3", "Heading 3"),
                                      (4, "h4", "Heading 4")):
        cfg = _style_cfg(get_style(style_name_h), preset[slot])
        hp = numbered_para(depth)
        if hp is not None:
            rs = [r for r in hp.runs if r.text.strip()]
            if rs:
                sz = next((r.font.size.pt for r in rs if r.font.size), None)
                cfg["size"] = sz if sz else body["size"]  # ke thua = co chu than van
                cfg["bold"] = any(bool(r.font.bold) for r in rs)
                cfg["italic"] = any(bool(r.font.italic) for r in rs)
            pf = hp.paragraph_format
            # File mau VN thuong dung thut dong dau lam bac thang tieu de (0 / 0,5 / 1,0).
            # Doan mau khong ghi thut le nghia la muc do sat le (0 cm) — ghi ro 0,
            # khong thua huong gia tri mac dinh cua preset goc.
            iv = pf.first_line_indent if pf.first_line_indent is not None else pf.left_indent
            cfg["left_indent_cm"] = round(iv.cm, 1) if iv is not None else 0.0
        preset[slot] = cfg

    # 4) Chu thich Hinh/Bang dau tien tim thay trong file mau (doc ca dam/nghieng)
    cap_para = next((p for p in content_paras
                     if RE_CAPTION.match(p.text.strip()) and p.runs), None)
    if cap_para is not None:
        r = cap_para.runs[0]
        cap = preset["caption"]
        cap["size"] = r.font.size.pt if r.font.size else body["size"]
        cap["italic"] = bool(r.font.italic)
        cap["bold"] = bool(r.font.bold)
        if cap_para.alignment is not None:
            cap["alignment"] = cap_para.alignment

    # 5) Co chu trong bang
    try:
        cell_run = next(r for t in sdoc.tables for row in t.rows for c in row.cells
                        for p in c.paragraphs for r in p.runs if r.font.size)
        preset["table_text"]["size"] = cell_run.font.size.pt
    except StopIteration:
        pass

    summary = {"font": preset["font_name"], "co_chu_than_van": preset["body"]["size"],
               "gian_dong": preset["body"]["line_spacing"],
               "le_cm": [preset["page"]["top_cm"], preset["page"]["bottom_cm"],
                         preset["page"]["left_cm"], preset["page"]["right_cm"]],
               "da_hoc": learned}
    return preset, summary


def extract_preset_from_samples(sample_paths):
    """Hoc quy tac tu NHIEU file mau (toi da 4 file): trich xuat tung file roi gop —
    moi thong so lay gia tri xuat hien NHIEU NHAT trong cac file, de mot file lech
    khong keo sai ca bo quy tac."""
    if len(sample_paths) == 1:
        return extract_preset_from_sample(sample_paths[0])
    presets = [extract_preset_from_sample(sp)[0] for sp in sample_paths]

    def merge(dicts):
        out = {}
        keys = []
        for d in dicts:
            for k in d:
                if k not in keys:
                    keys.append(k)
        for k in keys:
            # Co thong so chi xuat hien o mot so file (vd first_para_no_indent chi
            # set khi phat hien duoc) — chi gop tren cac file CO thong so do.
            vals = [d[k] for d in dicts if k in d]
            if isinstance(vals[0], dict):
                out[k] = merge(vals)
            else:
                v = _common(vals)
                if v is None:
                    v = next((x for x in vals if x is not None), vals[0])
                out[k] = v
        return out

    merged = merge(presets)
    merged["label"] = f"Mẫu tùy chỉnh (học từ {len(sample_paths)} file mẫu)"
    summary = {"font": merged["font_name"],
               "co_chu_than_van": merged["body"]["size"],
               "gian_dong": merged["body"]["line_spacing"],
               "le_cm": [merged["page"]["top_cm"], merged["page"]["bottom_cm"],
                         merged["page"]["left_cm"], merged["page"]["right_cm"]],
               "da_hoc": [f"Gộp từ {len(sample_paths)} file mẫu — mỗi thông số lấy "
                          "giá trị phổ biến nhất"],
               "so_file_mau": len(sample_paths)}
    return merged, summary


if __name__ == "__main__":
    inp, outp = sys.argv[1], sys.argv[2]
    preset = sys.argv[3] if len(sys.argv) > 3 else "chuan_vn"
    rep = format_document(inp, outp, preset)
    print("Bao cao chuan hoa:")
    for k, v in rep.items():
        print(f"  {k}: {v}")
    print("Da luu:", outp)
