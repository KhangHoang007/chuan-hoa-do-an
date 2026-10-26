# -*- coding: utf-8 -*-
"""Web app: Chuan hoa do an Word cho sinh vien — chay local tai 127.0.0.1:7862.

- Chon bo mau co san (do an, khoa luan, NCKH, thuc tap, tieu luan...)
- Hoac tai len FILE MAU dinh dang cua truong/giang vien -> tool tu hoc quy tac tu file mau
- Dat ten file ket qua tuy y
File xu ly xong tu xoa sau 1 gio.
"""
import glob
import os
import re
import threading
import time
import unicodedata
import uuid

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from format_engine import PRESETS, extract_preset_from_samples, format_document

import sys
import tempfile

# Khi dong goi thanh .exe (PyInstaller): tai nguyen (static) nam trong thu muc
# giai nen _MEIPASS, con file xu ly tam ghi ra thu muc temp cua he dieu hanh.
_FROZEN = getattr(sys, "frozen", False)
_RES_BASE = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
STATIC_DIR = os.path.join(_RES_BASE, "static")
if _FROZEN:
    WORK = os.path.join(tempfile.gettempdir(), "ChuanHoaDoAn")
else:
    WORK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "work")
os.makedirs(WORK, exist_ok=True)

app = FastAPI(title="Chuan hoa do an Word")
DISPLAY_NAMES = {}  # ten file luu tren dia -> ten hien thi nguoi dung dat


def cleanup_old():
    while True:
        now = time.time()
        for f in glob.glob(os.path.join(WORK, "*")):
            try:
                if now - os.path.getmtime(f) > 3600:
                    os.remove(f)
            except OSError:
                pass
        time.sleep(600)


threading.Thread(target=cleanup_old, daemon=True).start()


def sanitize_output_name(name: str, fallback: str) -> str:
    name = unicodedata.normalize("NFC", (name or "").strip())
    name = re.sub(r"[\\/:*?\"<>|]", "", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    if name.lower().endswith(".docx"):
        name = name[:-5]
    if not name:
        name = fallback
    return name[:80] + ".docx"


@app.get("/api/presets")
def presets():
    return [{"id": k, "label": v["label"], "desc": v.get("desc", "")}
            for k, v in PRESETS.items()]


@app.get("/api/health")
def health():
    return {"ok": True}


@app.post("/api/format")
async def do_format(file: UploadFile = File(...),
                    preset: str = Form("chuan_vn"),
                    samples: list[UploadFile] = File(default=[]),
                    output_name: str = Form(""),
                    page_numbers: bool = Form(True),
                    toc: bool = Form(True)):
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(400, "Chi nhan file .docx")
    data = await file.read()
    if len(data) > 30 * 1024 * 1024:
        raise HTTPException(400, "File qua lon (toi da 30 MB)")
    uid = uuid.uuid4().hex[:10]
    stem = os.path.splitext(os.path.basename(file.filename))[0]
    inp = os.path.join(WORK, f"{uid}_in.docx")
    outp = os.path.join(WORK, f"{uid}_out.docx")
    with open(inp, "wb") as f:
        f.write(data)

    used_preset = preset
    sample_summary = None
    sample_paths = []
    try:
        real_samples = [s for s in samples if s is not None and s.filename]
        if real_samples:
            if len(real_samples) > 4:
                raise HTTPException(400, "Toi da 4 file mau")
            for n, s in enumerate(real_samples):
                if not s.filename.lower().endswith(".docx"):
                    raise HTTPException(400, "File mau phai la .docx")
                sp = os.path.join(WORK, f"{uid}_sample{n}.docx")
                with open(sp, "wb") as f:
                    f.write(await s.read())
                sample_paths.append(sp)
            try:
                preset_obj, sample_summary = extract_preset_from_samples(sample_paths)
            except Exception as e:
                raise HTTPException(400, f"Khong doc duoc file mau: {e}")
            used_preset = "custom"
        else:
            if preset not in PRESETS:
                raise HTTPException(400, "Preset khong hop le")
            preset_obj = preset
        report = format_document(inp, outp, preset_obj,
                                 add_page_numbers=page_numbers, add_toc=toc)
    finally:
        for pth in [inp] + sample_paths:
            if pth:
                try:
                    os.remove(pth)
                except OSError:
                    pass
    display_name = sanitize_output_name(output_name, f"{stem}_da-chuan-hoa")
    DISPLAY_NAMES[os.path.basename(outp)] = display_name
    return {"download_url": f"/api/download/{os.path.basename(outp)}",
            "filename": display_name, "report": report,
            "preset_used": used_preset, "sample_summary": sample_summary}


@app.get("/api/download/{name}")
def download(name: str):
    base = os.path.basename(name)
    path = os.path.join(WORK, base)
    if not os.path.isfile(path):
        raise HTTPException(404, "File khong con ton tai (da het han 1 gio)")
    return FileResponse(path, filename=DISPLAY_NAMES.get(base, base),
                        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
