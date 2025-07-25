# -*- coding: utf-8 -*-
"""
hotkey_ocr_fast_tuned2.py

- 速さを落とさず微細な誤字をさらに減らす版
  * アンシャープマスクで軽くシャープ化
  * psm=6/7 をスコア( conf + length + jp_ratio )で比較
  * jpn → 必要時だけ jpn+eng に 1 回だけ挑戦
  * conf フィルタ(65)で短すぎたら 60 に緩めて再構成（OCRはやり直さない）
  * ヒューリスティック補正（“かなの間の1文字漢字”や連続記号など）
  * （任意）低 conf 行のみ再OCR（デフォルトOFF）

Ctrl+Alt+S : Snipping → OCR
Ctrl+Alt+Q / Esc : Exit
"""

import os
import time
import re
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Optional, Union, Tuple

import keyboard
import pyperclip
from PIL import Image, ImageGrab
import pytesseract

import cv2
import numpy as np
import pandas as pd
from pytesseract import Output

# ======== 設定 ========
TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

LANG_PRIMARY   = "jpn"        # まずこれ
LANG_SECONDARY = "jpn+eng"    # 英字が多そうなら一回だけ試す

PSMS = [6, 7]                 # 最大2回
CONF_TH_INIT  = 65
CONF_TH_RELAX = 60
EARLY_ACCEPT_CONF = 86.0      # これ以上なら即決
MIN_TEXT_LEN  = 10

RE_OCR_LOWCONF = False        # 低conf行再OCR（必要時だけ True に）
LINE_CONF_TH   = 70

TESSDATA_DIR = ""             # 固定したい場合だけ指定
EXTRA_FLAGS = f'--tessdata-dir "{TESSDATA_DIR}"' if TESSDATA_DIR else ""

OUT_DIR = Path(r"D:\Python\OCR\Hotkey_ocr")
TRIGGER_SNIP = True

OPEN_AFTER_SAVE   = True
OPEN_WITH_NOTEPAD = False

DEBUG = False

# ======== 初期化 ========
pytesseract.pytesseract.tesseract_cmd = TESSERACT
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ======== ユーティリティ ========
def launch_snipping_tool() -> None:
    try:
        subprocess.Popen(["explorer", "ms-screenclip:"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return
    except Exception:
        pass
    try:
        subprocess.Popen([r"C:\Windows\system32\SnippingTool.exe", "/clip"],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def wait_clipboard_image(timeout: float = 30.0) -> Optional[Image.Image]:
    end = time.time() + timeout
    while time.time() < end:
        data: Union[None, Image.Image, list] = ImageGrab.grabclipboard()
        if isinstance(data, Image.Image):
            return data
        if isinstance(data, list):
            try:
                if data and isinstance(data[0], str):
                    return Image.open(data[0])
            except Exception:
                pass
        time.sleep(0.1)
    return None

def auto_invert_if_needed(gray: np.ndarray) -> np.ndarray:
    return 255 - gray if gray.mean() < 128 else gray

def unsharp(gray: np.ndarray) -> np.ndarray:
    # 軽いアンシャープマスク
    blur = cv2.GaussianBlur(gray, (0, 0), 1.0)
    sharp = cv2.addWeighted(gray, 1.5, blur, -0.5, 0)
    return np.clip(sharp, 0, 255).astype(np.uint8)

def light_preprocess(pil_im: Image.Image) -> Image.Image:
    g = np.array(pil_im.convert("L"))
    g = cv2.resize(g, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    g = auto_invert_if_needed(g)
    g = unsharp(g)
    return Image.fromarray(g)

def df_to_text(df: pd.DataFrame, conf_th: Optional[int], jpn: bool) -> str:
    if df.empty:
        return ""
    df = df[df.conf.astype(float) >= 0].copy()
    if conf_th is not None:
        df = df[df.conf >= conf_th]
    lines = []
    for (_, _, _), line_df in df.groupby(["block_num", "par_num", "line_num"], sort=True):
        words = [w for w in line_df.sort_values("word_num")["text"].dropna().tolist() if w.strip()]
        if not words:
            continue
        joiner = "" if jpn else " "
        lines.append(joiner.join(words))
    return "\n".join(lines)

def normalize_ws(text: str) -> str:
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = "\n".join(line.strip() for line in text.splitlines())
    return text

KANA = r"ぁ-んァ-ヶー"
def heuristic_fix(text: str) -> str:
    # かなの間に 1 文字だけ漢字が挟まった場合に削除
    text = re.sub(fr'(?<=[{KANA}])[一-龥々〆ヵヶ](?=[{KANA}])', '', text)
    # 具体的に気になるパターンはここに追加
    text = text.replace("で和複製", "で複製")
    # 重複記号の削減
    text = re.sub(r'([=、。．．…])\1+', r'\1', text)
    # 句読点/括弧まわりのスペース整理
    text = re.sub(r'\s+([、。．，））])', r'\1', text)
    text = re.sub(r'（\s+', '（', text)
    return normalize_ws(text)

def jp_ratio(text: str) -> float:
    if not text:
        return 0.0
    jp = sum(("ぁ" <= c <= "ん") or ("ァ" <= c <= "ヶ") or ("一" <= c <= "龥") for c in text)
    return jp / len(text)

def need_eng(text: str, ratio=0.1) -> bool:
    if not text:
        return False
    ascii_cnt = sum(ord(c) < 128 for c in text)
    return (ascii_cnt / len(text)) > ratio

def score_text(text: str, conf: float) -> float:
    # conf を主、長さと日本語率で微調整
    return conf + min(len(text.strip()) / 500.0, 1.0) + jp_ratio(text) * 0.5

def ocr_df(pil_im: Image.Image, lang: str, psm: int) -> Tuple[pd.DataFrame, float]:
    config = fr"--psm {psm} --oem 3 -c user_defined_dpi=300 preserve_interword_spaces=1 {EXTRA_FLAGS}"
    df = pytesseract.image_to_data(pil_im, lang=lang, config=config, output_type=Output.DATAFRAME)
    if df.empty:
        return df, 0.0
    valid = df[df.conf.astype(float) >= 0]
    mean_conf = valid.conf.mean() if not valid.empty else 0.0
    return df, float(mean_conf)

def reconstruct_text_from_df(df: pd.DataFrame, lang: str) -> str:
    text = df_to_text(df, CONF_TH_INIT, jpn=("jpn" in lang))
    if len(text.strip()) < MIN_TEXT_LEN:
        text = df_to_text(df, CONF_TH_RELAX, jpn=("jpn" in lang))
    return text

def tsv_line_mean_conf(df: pd.DataFrame) -> pd.Series:
    valid = df[df.conf.astype(float) >= 0]
    if valid.empty:
        return pd.Series([], dtype=float)
    return valid.groupby(["block_num", "par_num", "line_num"]).conf.mean()

def reocr_low_conf_lines(pil_im: Image.Image, base_df: pd.DataFrame, lang: str) -> str:
    if base_df.empty:
        return ""
    groups = base_df.groupby(["block_num", "par_num", "line_num"])
    mean_conf_per_line = tsv_line_mean_conf(base_df)
    out_lines = []
    rgb = np.array(pil_im.convert("RGB"))
    for key, line_df in groups:
        line_text = " ".join(w for w in line_df["text"].dropna().tolist() if w.strip())
        if mean_conf_per_line.get(key, 100.0) >= LINE_CONF_TH or not line_text.strip():
            out_lines.append(line_text)
            continue
        x, y = line_df.left.min(), line_df.top.min()
        w = (line_df.left + line_df.width).max() - x
        h = (line_df.top + line_df.height).max() - y
        crop = Image.fromarray(rgb[y:y+h, x:x+w])
        config = fr"--psm 7 --oem 3 -c user_defined_dpi=300 preserve_interword_spaces=1 {EXTRA_FLAGS}"
        improved = pytesseract.image_to_string(crop, lang=lang, config=config)
        out_lines.append(improved.strip() if improved.strip() else line_text)
    return "\n".join(out_lines)

def fast_best_ocr(img: Image.Image) -> Tuple[str, float, int, str]:
    pil = light_preprocess(img)

    # 1) psm6/7 @ jpn
    best_text, best_conf, best_psm, best_lang = "", -1.0, 6, LANG_PRIMARY

    for psm in PSMS:
        df, conf = ocr_df(pil, LANG_PRIMARY, psm)
        txt = reconstruct_text_from_df(df, LANG_PRIMARY)

        if RE_OCR_LOWCONF and not df.empty:
            txt_alt = reocr_low_conf_lines(pil, df, LANG_PRIMARY)
            if len(txt_alt.strip()) > len(txt.strip()):
                txt = txt_alt

        sc = score_text(txt, conf)
        if sc > score_text(best_text, best_conf):
            best_text, best_conf, best_psm, best_lang = txt, conf, psm, LANG_PRIMARY

    # 早期 accept
    if (best_conf >= EARLY_ACCEPT_CONF and len(best_text.strip()) >= MIN_TEXT_LEN and jp_ratio(best_text) > 0.6):
        return heuristic_fix(best_text), best_conf, best_psm, best_lang

    # 2) 英字が多そうなら jpn+eng を 1回だけ試す
    if need_eng(best_text):
        for psm in PSMS:
            df, conf = ocr_df(pil, LANG_SECONDARY, psm)
            txt = reconstruct_text_from_df(df, LANG_SECONDARY)

            if RE_OCR_LOWCONF and not df.empty:
                txt_alt = reocr_low_conf_lines(pil, df, LANG_SECONDARY)
                if len(txt_alt.strip()) > len(txt.strip()):
                    txt = txt_alt

            sc = score_text(txt, conf)
            if sc > score_text(best_text, best_conf):
                best_text, best_conf, best_psm, best_lang = txt, conf, psm, LANG_SECONDARY

    return heuristic_fix(best_text), best_conf, best_psm, best_lang

def open_with_notepad(path: Path) -> None:
    if OPEN_WITH_NOTEPAD:
        subprocess.Popen(["notepad.exe", str(path)],
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    else:
        os.startfile(path)

def ocr_and_output(img: Image.Image) -> Path:
    text, conf, psm, lang = fast_best_ocr(img)

    pyperclip.copy(text)

    out = OUT_DIR / f"{datetime.now():%Y%m%d_%H%M%S}.txt"
    out.write_text(text, encoding="utf-8-sig")

    if OPEN_AFTER_SAVE:
        open_with_notepad(out)

    print(f"  conf={conf:.1f}, psm={psm}, lang={lang}, len={len(text)}")
    return out

# ======== Hotkey ========
def do_flow():
    if TRIGGER_SNIP:
        launch_snipping_tool()
        print("…範囲選択してください（タイムアウト: 30秒）")

    img = wait_clipboard_image(timeout=30)
    if img is None:
        print("✖ 画像が取得できませんでした")
        return

    out = ocr_and_output(img)
    print(f"✔ OCR 完了 → クリップボードへコピー / {out}")

def main():
    print("=== Hotkey OCR Launcher (fast tuned2) ===")
    print("Ctrl+Alt+S : Snipping → OCR")
    print("Ctrl+Alt+Q / Esc : Exit")
    print("OUT_DIR   :", OUT_DIR.resolve())
    print("Tesseract :", pytesseract.get_tesseract_version())
    print("LANG_PRIMARY   :", LANG_PRIMARY)
    print("LANG_SECONDARY :", LANG_SECONDARY)
    print("PSMS      :", PSMS)
    print("CONF_TH   :", CONF_TH_INIT, " (relax ->", CONF_TH_RELAX, ")")
    print("EARLY_ACCEPT_CONF:", EARLY_ACCEPT_CONF)
    print("RE_OCR_LOWCONF   :", RE_OCR_LOWCONF, "(line_conf_th =", LINE_CONF_TH, ")")

    keyboard.add_hotkey("ctrl+alt+s", do_flow)
    keyboard.add_hotkey("ctrl+alt+q", lambda: (_ for _ in ()).throw(SystemExit))
    keyboard.add_hotkey("esc",        lambda: (_ for _ in ()).throw(SystemExit))

    try:
        while True:
            time.sleep(1.0)
    except SystemExit:
        print("Bye!")

if __name__ == "__main__":
    main()
