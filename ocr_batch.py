from pathlib import Path
from PIL import Image, ImageOps
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

IN_DIR   = Path(r"D:\Python\OCR")
EXTS     = {".jpg", ".jpeg", ".png"}
LANG     = "jpn"          # 英数字も多いなら "jpn+eng"
CFG      = r"--psm 6 --oem 3"
ENCODING = "utf-8-sig"

def preprocess(im: Image.Image) -> Image.Image:
    # 必要なければ外してください（小さいUI文字のとき精度が上がります）
    im = im.convert("L")                         # グレースケール
    im = im.resize((im.width * 2, im.height * 2))# 2倍拡大
    im = im.point(lambda x: 0 if x < 180 else 255, '1')  # 単純二値化
    return im

def main():
    files = [p for p in sorted(IN_DIR.iterdir()) if p.suffix.lower() in EXTS]
    if not files:
        print("対象画像がありません")
        return

    ok = ng = 0
    for img in files:
        try:
            print(f"[OCR] {img.name} ... ", end="")
            with Image.open(img) as im:
                im   = preprocess(im)
                text = pytesseract.image_to_string(im, lang=LANG, config=CFG)
            img.with_suffix(".txt").write_text(text, encoding=ENCODING)
            print("OK")
            ok += 1
        except Exception as e:
            print("FAILED:", e)
            ng += 1

    print(f"\n=== 完了 === 成功: {ok} / 失敗: {ng}")

if __name__ == "__main__":
    main()
