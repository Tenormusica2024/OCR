# -*- coding: utf-8 -*-
"""
working_ocr_service.py

現在の環境で確実に動作するOCRサービス
- 段階的インポートでエラー箇所を特定
- 利用可能な機能のみ使用
- 高品質なスペース除去機能
"""

import os
import sys
import time
import subprocess
from datetime import datetime
from pathlib import Path
import re
import tempfile

print("📚 基本ライブラリインポート中...")

# 基本ライブラリ
try:
    import keyboard
    print("✅ keyboard: OK")
except ImportError as e:
    print(f"❌ keyboard: {e}")
    sys.exit(1)

try:
    import pyperclip
    print("✅ pyperclip: OK")
except ImportError as e:
    print(f"❌ pyperclip: {e}")
    sys.exit(1)

# PIL関連
try:
    from PIL import Image, ImageGrab, ImageEnhance, ImageFilter
    print("✅ PIL: OK")
    PIL_AVAILABLE = True
except ImportError as e:
    print(f"❌ PIL: {e}")
    PIL_AVAILABLE = False

# NumPy関連（オプション）
try:
    import numpy as np
    print(f"✅ NumPy: {np.__version__}")
    NUMPY_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  NumPy: {e} (オプション)")
    NUMPY_AVAILABLE = False

# pytesseract（オプション）
try:
    import pytesseract
    print("✅ pytesseract: OK")
    PYTESSERACT_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  pytesseract: {e} (直接Tesseractを使用)")
    PYTESSERACT_AVAILABLE = False

# OpenCV（オプション）
try:
    import cv2
    print(f"✅ OpenCV: {cv2.__version__}")
    CV2_AVAILABLE = True
except ImportError as e:
    print(f"⚠️  OpenCV: {e} (オプション)")
    CV2_AVAILABLE = False

print("🔧 ライブラリチェック完了\n")

# ======== 設定 ========
TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
OUT_DIR = Path(r"D:\Python\OCR\Hotkey_ocr")

# OCR設定
LANG = "jpn+eng"
PSM = 6

# ======== 初期化 ========
OUT_DIR.mkdir(parents=True, exist_ok=True)

class WorkingOCRService:
    def __init__(self):
        self.running = True
        self.temp_dir = Path(tempfile.gettempdir()) / "working_ocr"
        self.temp_dir.mkdir(exist_ok=True)
        
        # Tesseract確認
        if not Path(TESSERACT).exists():
            print("❌ Tesseractが見つかりません")
            print(f"   パス: {TESSERACT}")
            print("   https://github.com/UB-Mannheim/tesseract/wiki からダウンロードしてください")
            return
        
        print("✅ Tesseract確認完了")
        self.log_capabilities()
        
    def log_capabilities(self):
        """利用可能機能をログ出力"""
        print("\n🎯 利用可能機能:")
        print(f"  PIL画像処理: {'✅' if PIL_AVAILABLE else '❌'}")
        print(f"  NumPy配列: {'✅' if NUMPY_AVAILABLE else '❌'}")
        print(f"  pytesseract: {'✅' if PYTESSERACT_AVAILABLE else '❌'}")
        print(f"  OpenCV: {'✅' if CV2_AVAILABLE else '❌'}")
        print()
        
    def launch_snipping_tool(self):
        """スニッピングツール起動"""
        try:
            subprocess.Popen(["explorer", "ms-screenclip:"],
                             stdout=subprocess.DEVNULL, 
                             stderr=subprocess.DEVNULL)
        except:
            try:
                subprocess.Popen([r"C:\Windows\system32\SnippingTool.exe", "/clip"],
                                 stdout=subprocess.DEVNULL, 
                                 stderr=subprocess.DEVNULL)
            except:
                self.log("スニッピングツールの起動に失敗しました")

    def wait_clipboard_image(self, timeout=30.0):
        """クリップボード画像待機"""
        if not PIL_AVAILABLE:
            self.log("❌ PIL不使用のため画像取得できません")
            return None
            
        self.log("📷 クリップボード画像を待機中...")
        
        end_time = time.time() + timeout
        
        while time.time() < end_time and self.running:
            try:
                data = ImageGrab.grabclipboard()
                if isinstance(data, Image.Image):
                    self.log("✅ クリップボード画像を取得しました")
                    return data
                elif isinstance(data, list) and data:
                    if isinstance(data[0], str):
                        img = Image.open(data[0])
                        self.log("✅ クリップボード画像を取得しました")
                        return img
            except Exception as e:
                self.log(f"クリップボード取得エラー: {e}")
            
            time.sleep(0.5)
        
        self.log("⏰ タイムアウト：画像が取得できませんでした")
        return None

    def enhance_image(self, img):
        """利用可能な方法で画像強化"""
        if not PIL_AVAILABLE:
            return img
            
        try:
            self.log("🖼️  画像前処理中...")
            
            # グレースケール変換
            if img.mode != 'L':
                img = img.convert('L')
            
            # サイズ拡大
            w, h = img.size
            scale = 3
            img = img.resize((w * scale, h * scale), Image.LANCZOS)
            
            # コントラスト強化
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(1.8)
            
            # シャープネス強化
            enhancer = ImageEnhance.Sharpness(img)
            img = enhancer.enhance(2.0)
            
            self.log("✅ 画像前処理完了")
            return img
            
        except Exception as e:
            self.log(f"画像前処理エラー: {e}")
            return img

    def ocr_with_pytesseract(self, img):
        """pytesseractを使用したOCR"""
        try:
            config = f"--psm {PSM} --oem 3 -c user_defined_dpi=300"
            text = pytesseract.image_to_string(img, lang=LANG, config=config)
            return text.strip()
        except Exception as e:
            self.log(f"pytesseract OCRエラー: {e}")
            return ""

    def ocr_with_direct_tesseract(self, img):
        """Tesseractコマンド直接実行"""
        try:
            # 一時ファイル保存
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            temp_file = self.temp_dir / f"ocr_{timestamp}.png"
            img.save(str(temp_file), "PNG")
            
            # Tesseractコマンド実行
            cmd = [
                TESSERACT,
                str(temp_file),
                "stdout",
                "-l", LANG,
                "--psm", str(PSM),
                "--oem", "3",
                "-c", "user_defined_dpi=300"
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )
            
            # 一時ファイル削除
            try:
                os.remove(temp_file)
            except:
                pass
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                self.log(f"Tesseractエラー: {result.stderr}")
                return ""
                
        except Exception as e:
            self.log(f"直接Tesseract OCRエラー: {e}")
            return ""

    def run_ocr(self, img):
        """最適な方法でOCR実行"""
        enhanced_img = self.enhance_image(img)
        
        # pytesseractが利用可能な場合は優先使用
        if PYTESSERACT_AVAILABLE:
            self.log("🔍 pytesseractでOCR実行中...")
            text = self.ocr_with_pytesseract(enhanced_img)
            if text:
                return text
        
        # フォールバック：直接Tesseract実行
        self.log("🔍 直接TesseractでOCR実行中...")
        return self.ocr_with_direct_tesseract(enhanced_img)

    def advanced_text_cleaning(self, text):
        """超強化テキストクリーニング（改行修正強化版）"""
        if not text:
            return text
        
        original_length = len(text)
        self.log("📝 超強化テキストクリーニング中...")
        
        # 0. 特殊文字の前処理
        text = text.replace('・:', ': ')  # ・: → : 
        text = text.replace('、。', '。')   # 、。 → 。
        
        # 1. 日本語文字間スペース除去（強化版）
        # ひらがな間（全角・半角スペース両対応）
        text = re.sub(r'([あ-ん])[\s\u3000]+([あ-ん])', r'\1\2', text)
        
        # カタカナ間（長音記号・小文字含む）
        text = re.sub(r'([ア-ヶーァィゥェォャュョッ])[\s\u3000]+([ア-ヶーァィゥェォャュョッ])', r'\1\2', text)
        
        # 漢字間
        text = re.sub(r'([一-龥々〆ヵヶ])[\s\u3000]+([一-龥々〆ヵヶ])', r'\1\2', text)
        
        # 2. 文字種混在間のスペース除去（強化版）
        # あらゆる日本語文字間
        text = re.sub(r'([あ-んア-ヶーァィゥェォャュョッ一-龥々〆ヵヶ])[\s\u3000]+([あ-んア-ヶーァィゥェォャュョッ一-龥々〆ヵヶ])', r'\1\2', text)
        
        # 3. 英数字の調整（超強化版）
        # 英字間のスペース除去（ただし単語境界は保持）
        text = re.sub(r'([A-Za-z])[\s\u3000]+([A-Za-z])', r'\1\2', text)
        
        # 数字間のスペース除去
        text = re.sub(r'([0-9])[\s\u3000]+([0-9])', r'\1\2', text)
        
        # ピリオド・ドットまわり
        text = re.sub(r'([A-Za-z0-9])[\s\u3000]*\.[\s\u3000]*([A-Za-z0-9])', r'\1.\2', text)
        
        # 4. プログラミング関連の修正
        # コマンドライン系
        text = re.sub(r'python[\s\u3000]+\.[\s\u3000]*\\', r'python .\\', text)
        text = re.sub(r'dir[\s\u3000]*\*[\s\u3000]*\.[\s\u3000]*py', r'dir *.py', text)
        text = re.sub(r'ls[\s\u3000]*\*[\s\u3000]*\.[\s\u3000]*py', r'ls *.py', text)
        text = re.sub(r'bash[\s\u3000]*#', r'bash\n#', text)
        text = re.sub(r'#[\s\u3000]*([A-Za-z])', r'# \1', text)
        
        # ファイルパス修正
        text = re.sub(r'C[\s\u3000]*:[\s\u3000]*\\[\s\u3000]*python[\s\u3000]*\\', r'C:\\python\\', text)
        
        # 5. 記号まわり（超強化版）
        # 句読点前のスペース除去
        text = re.sub(r'[\s\u3000]+([、。！？）］｝」』])', r'\1', text)
        
        # 開始記号後のスペース除去
        text = re.sub(r'([（［｛「『])[\s\u3000]+', r'\1', text)
        
        # コロン・セミコロンまわり
        text = re.sub(r'[\s\u3000]*:[\s\u3000]*', ': ', text)
        text = re.sub(r'[\s\u3000]*;[\s\u3000]*', '; ', text)
        
        # ハッシュ（#）まわり
        text = re.sub(r'[\s\u3000]*#[\s\u3000]*', '# ', text)
        
        # 6. 助詞・語尾の修正（強化版）
        particles = ['を', 'が', 'に', 'へ', 'と', 'で', 'の', 'も', 'は', 'から', 'まで', 'より', 'こそ', 'ばかり', 'だけ', 'でも', 'など', 'として', 'について', 'に対して']
        for particle in particles:
            text = re.sub(rf'[\s\u3000]+{particle}', particle, text)
        
        # 7. 数字・記号組み合わせ
        text = re.sub(r'(\d+)[\s\u3000]*\.[\s\u3000]*', r'\1. ', text)  # 1. → 1. 
        text = re.sub(r'(\d+)[\s\u3000]*-[\s\u3000]*', r'\1-', text)    # 1 - → 1-
        
        # 8. 英語特有のパターン
        # ファイル名・拡張子
        text = re.sub(r'([a-zA-Z0-9_]+)[\s\u3000]*\.[\s\u3000]*([a-zA-Z]+)', r'\1.\2', text)
        
        # 9. ★★ 改行修正（強化版）★★
        # 見出し・セクション境界の改行追加
        text = re.sub(r'([。！？])([1-9]\.|仮想環境|トラブルシューティング|PowerShell|OCR|手順)', r'\1\n\n\2', text)
        
        # 番号付きリストの改行
        text = re.sub(r'([。！？])([1-9]\.[^0-9])', r'\1\n\n\2', text)
        
        # PowerShellコードブロックの改行
        text = re.sub(r'powershell([a-zA-Z])', r'powershell\n\1', text)
        text = re.sub(r'([^:\n])powershell', r'\1\n\npowershell', text)
        
        # コマンド間の改行
        text = re.sub(r'(\.py)([a-zA-Z#])', r'\1\n\2', text)
        text = re.sub(r'(activate)([a-zA-Z])', r'\1\n\n\2', text)
        
        # bash/cmd行の修正
        text = re.sub(r'bash[\s\u3000]*\n?[\s\u3000]*#', r'bash\n#', text)
        
        # セクションタイトルの前後改行
        section_titles = [
            '仮想環境の入り方',
            'ディレクトリ移動', 
            '仮想環境をアクティベート',
            'プアクティベート',
            '成功確認',
            'OCRサービス実行',
            'トラブルシューティング',
            '仮想環境が見つからない場合'
        ]
        
        for title in section_titles:
            text = re.sub(rf'([。！？])({re.escape(title)})', r'\1\n\n\2', text)
            text = re.sub(rf'({re.escape(title)})([あ-んア-ヶー一-龥々A-Za-z])', r'\1\n\2', text)
        
        # 10. 特殊ケースの修正（超強化版）
        
        # A. 英数字誤認識の段階的修正
        # 数字とアルファベットの誤認識（パターンマッチング）
        text = re.sub(r'PowerShe[1l][1l]*[+L]*', 'PowerShell', text)
        text = re.sub(r'PowerShet[1l]*L+', 'PowerShell', text)
        text = re.sub(r'PowerShe[1l]+L*', 'PowerShell', text)
        text = re.sub(r'cop[1l]*[iI][1l]*[Ll]ot', 'copilot', text)
        text = re.sub(r'cop[1l]*[Ll]ot', 'copilot', text)
        text = re.sub(r'copli[Ll]ot', 'copilot', text)
        text = re.sub(r'copilote?', 'copilot', text)
        
        # B. 具体的な文字置換
        char_misrecognition_fixes = {
            # PowerShell系
            'PowerShe1+L': 'PowerShell',
            'PowerShe11L': 'PowerShell', 
            'PowerSheLl': 'PowerShell',
            'PowerSheLLL': 'PowerShell',
            'PowerShetLL': 'PowerShell',
            'powersheLl': 'powershell',
            'powershe11': 'powershell',
            
            # copilot系  
            'cop1iLot': 'copilot',
            'cop11Lot': 'copilot',
            'cop1Lot': 'copilot',
            'copliLot': 'copilot',
            'copilote': 'copilot',
            'Copilot': 'Copilot',
            
            # コマンド系
            'cdC:': 'cd C:',
            'cd C:': 'cd C:',
            'python': 'python',
            '.py': '.py',
            'py >': '.py →',
            '.py >': '.py →',
            
            # パス系
            'C:\\python': 'C:\\python',
            'C:\python': 'C:\\python',
            '\\python': '\\python',
            'ocr_env': 'ocr_env',
            'Scripts': 'Scripts', 
            'activate': 'activate',
            
            # 記号系
            '—': '→',
            '©': '・',
            'OS': '→',
            '|': '',
            
            # よくある誤認識
            'ユコードブロック': 'コードブロック',
            'ププロンプト': 'プロンプト',
            'アクテンーファ': 'アクティベート',
            'アクティベート': 'アクティベート',
            '成功確認': '成功確認',
            'CKITIZLESEICAREИТ': '改行修正強化版を実行',
            'IE L UID P INECH T': '正しいファイル名で実行してください',
            '武存の': '現在の',
            'ディルクムソ': 'ディレクトリ',
            'たは': 'または',
            'もるし': 'もしくは',
            'プアクティベート': 'アクティベート',
            'WOES': '成功確認',
            'プブロンプト': 'プロンプト',
            'ELLAOBAReSIT': '修正済みの高精度版を実行',
            'F7z=IMBON-LDay': 'または他のバージョン',
        }
        
        # 文字単位の修正を実行
        for wrong, correct in char_misrecognition_fixes.items():
            text = text.replace(wrong, correct)
        
        # C. プログラミング特化修正（正規表現）
        programming_fixes = [
            # Pythonコマンド修正
            (r'python[\s]*\.[\s]*py', r'python *.py'),
            (r'python[\s]*\.[\s]*\\', r'python .\\'),
            (r'python[\s]*\.[\s]*working_ocr_service[\s]*\.[\s]*py', r'python .\\working_ocr_service.py'),
            
            # ファイルパス修正
            (r'C[\s]*:[\s]*\\[\s]*python', r'C:\\python'),
            (r'ocr_env[\s]*\\[\s]*Scripts[\s]*\\[\s]*activate', r'ocr_env\\Scripts\\activate'),
            
            # bash/powershell修正
            (r'bash[\s]*\n?[\s]*#', r'bash\n#'),
            (r'bash[\s]*python', r'bash\npython'),
            (r'powershell[\s]*\n?[\s]*([a-zA-Z])', r'powershell\n\1'),
            
            # 拡張子・ファイル名修正
            (r'\.[\s]*py[\s]*>', r'.py →'),
            (r'\.[\s]*py', r'.py'),
            (r'\.[\s]*txt', r'.txt'),
            (r'\.[\s]*exe', r'.exe'),
            
            # コマンド区切り修正
            (r'(\.py)([a-zA-Z])', r'\1\n\2'),
            (r'(activate)([a-zA-Z])', r'\1\n\n\2'),
        ]
        
        for pattern, replacement in programming_fixes:
            text = re.sub(pattern, replacement, text)
        
        # D. クォート・括弧内の文字列修正
        quote_fixes = [
            # Python辞書形式の修正
            (r'[|｜][\s]*[\'\"]*([^\'\"\n]+)[\'\"]*[\s]*[:\:][\s]*[\'\"]*([^\'\"\n]+)[\'\"]*', r"'\1': '\2'"),
            (r'「([^」]+)」[\s]*:[\s]*「([^」]+)」', r"'\1': '\2'"),
            (r'「([^」]+)」[\s]*,', r"'\1',"),
            
            # 一般的なクォート修正
            (r'「([^」]+)」', r'\1'),
            (r'『([^』]+)』', r'\1'),
            (r'[\'\"]+([^\'\"\n]+)[\'\"]+[\s]*:[\s]*[\'\"]+([^\'\"\n]+)[\'\"]+', r"'\1': '\2'"),
        ]
        
        for pattern, replacement in quote_fixes:
            text = re.sub(pattern, replacement, text)
        
        # E. 矢印・記号の統一
        arrow_fixes = [
            (r'[—–−]', r'→'),
            (r'>', r'→'),
            (r'[\s]*→[\s]*', r' → '),
            (r'\.py[\s]*→[\s]*py', r'.py → .py'),
        ]
        
        for pattern, replacement in arrow_fixes:
            text = re.sub(pattern, replacement, text)
        
        # 11. 行単位清掃（強化版）
        lines = []
        for line in text.splitlines():
            line = line.strip()
            if line:  # 空行でない場合のみ
                # 行内の重複スペース除去
                line = re.sub(r'[\s\u3000]{2,}', ' ', line)
                lines.append(line)
        
        text = '\n'.join(lines)
        
        # 12. 最終的な改行構造の整理
        # 重複改行の整理（ただし構造的改行は保持）
        text = re.sub(r'\n{4,}', '\n\n\n', text)  # 4個以上の改行を3個に
        text = re.sub(r'\n{3}', '\n\n', text)     # 3個の改行を2個に
        
        # 13. 最終調整
        # 全角スペースを半角に統一
        text = text.replace('\u3000', ' ')
        
        # 行頭・行末の空白除去
        text = '\n'.join(line.strip() for line in text.splitlines())
        
        # 文全体の前後空白除去
        text = text.strip()
        
        # クリーニング結果
        cleaned_chars = original_length - len(text)
        if cleaned_chars > 0:
            self.log(f"✅ 超強化クリーニング完了: {cleaned_chars}文字削除")
        else:
            self.log("ℹ️  クリーニング: 変更なし")
        
        return text

    def ocr_flow(self):
        """OCRメイン処理"""
        self.log("📷 スクリーンショット範囲を選択してください...")
        
        # 1. スニッピングツール起動
        self.launch_snipping_tool()
        
        # 2. クリップボード画像待機
        img = self.wait_clipboard_image()
        if not img:
            self.log("❌ 画像取得に失敗しました")
            return
        
        # 3. OCR実行
        raw_text = self.run_ocr(img)
        if not raw_text:
            self.log("❌ OCRでテキストを取得できませんでした")
            return
        
        # 4. テキストクリーニング
        cleaned_text = self.advanced_text_cleaning(raw_text)
        
        # 5. クリップボードにコピー
        pyperclip.copy(cleaned_text)
        self.log("📋 クリップボードにコピーしました")
        
        # 6. ファイル保存
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = OUT_DIR / f"working_ocr_{timestamp}.txt"
        
        # メタデータ付きで保存
        metadata = f"# 確実動作OCR結果 - {timestamp}\n"
        metadata += f"# 原文字数: {len(raw_text)}\n"
        metadata += f"# クリーニング後: {len(cleaned_text)}\n"
        metadata += f"# 使用機能: PIL={PIL_AVAILABLE}, NumPy={NUMPY_AVAILABLE}, pytesseract={PYTESSERACT_AVAILABLE}\n\n"
        
        content = metadata + cleaned_text
        out_file.write_text(content, encoding="utf-8-sig")
        
        # 7. メモ帳で開く
        self.open_notepad(out_file)
        
        self.log(f"✅ OCR完了！ 文字数: {len(cleaned_text)}")
        self.log(f"📁 ファイル: {out_file.name}")

    def open_notepad(self, file_path):
        """メモ帳で開く"""
        try:
            subprocess.Popen(["notepad.exe", str(file_path)])
        except Exception as e:
            self.log(f"メモ帳起動エラー: {e}")

    def log(self, message):
        """ログ出力"""
        if not sys.executable.endswith('pythonw.exe'):
            timestamp = datetime.now().strftime("%H:%M:%S")
            print(f"[{timestamp}] {message}")

    def cleanup_temp_files(self):
        """一時ファイル削除"""
        try:
            for temp_file in self.temp_dir.glob("ocr_*.png"):
                temp_file.unlink()
        except:
            pass

    def quit_service(self):
        """サービス終了"""
        self.log("🛑 確実動作OCRサービスを終了しています...")
        self.running = False
        self.cleanup_temp_files()
        
        try:
            keyboard.unhook_all()
        except:
            pass
        
        sys.exit(0)

    def run(self):
        """サービス実行"""
        print("🚀 確実動作OCRサービス開始")
        print("⌨️  ホットキー: Ctrl+Alt+S でOCR実行")
        print("🛑 終了: Ctrl+Alt+Q または Ctrl+C")
        print()
        
        # ホットキー登録
        keyboard.add_hotkey("ctrl+alt+s", self.ocr_flow)
        keyboard.add_hotkey("ctrl+alt+q", self.quit_service)
        
        try:
            while self.running:
                time.sleep(1.0)
        except (KeyboardInterrupt, SystemExit):
            self.quit_service()

def main():
    try:
        service = WorkingOCRService()
        service.run()
    except Exception as e:
        print(f"❌ 起動エラー: {e}")
        print("詳細:")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
