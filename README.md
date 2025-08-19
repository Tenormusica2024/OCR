高精度OCRツール - 開発ポートフォリオ
【製品概要】
ワンクリック文字起こしツール
スニッピングツールとOCR技術を組み合わせ、画像・映像内のテキストを瞬時にデジタル化するWindowsデスクトップアプリケーション。生成AIとGoogle Vision APIの併用により、従来のOCRツールを大幅に上回る認識精度を実現。
主要機能

Windows起動時自動常駐
Ctrl+Alt+Sによるワンクリック文字起こし
高精度OCR（日本語・英語対応）
自動ログ保存・履歴管理
クリップボード連携
- **🆕 バックグラウンド実行モード（CMD画面非表示）**

【開発背景・課題設定】
解決すべき課題

手動文字入力による業務効率の低下
既存OCRツールの認識精度不足
文字起こし作業の煩雑な工程

技術的課題
タイピング熟練者にとって、既存のOCRツールは精度・速度の面で手打ちに劣ることが多く、導入メリットが薄い。真の効率化には「撮影→認識→テキスト化」のワンストップ処理が必要。
【技術仕様・実装】
アーキテクチャ

言語：Python 3.9
OCRエンジン：Tesseract + Google Vision API
画像処理：PIL/OpenCV
UI：キーボードフック（keyboard）
自動起動：Windowsレジストリ操作

最適化施策

マルチスケール画像前処理による認識精度向上
日本語特化の後処理アルゴリズム実装
非同期処理によるレスポンス速度改善

## 🆕 【バックグラウンド実行機能】

### 新機能概要
OCRサービスをCMD画面を表示せずにバックグラウンドで実行する機能を実装。ユーザーの作業画面を妨げることなく常駐サービスとして動作。

### 実装方法

#### 1. VBScript版 (`run_ocr_hidden.vbs`)
```vbs
' CMD画面を非表示でOCRサービスを起動
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c python background_ocr_service.py", 0, False
```

#### 2. PowerShell版 (`run_ocr_hidden.ps1`)
```powershell
# 非表示プロセスでOCRサービスを起動
$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
$psi.CreateNoWindow = $true
```

#### 3. Python Wrapper版 (`run_ocr_background.pyw`)
```python
# .pyw拡張子によりコンソール非表示で実行
subprocess.Popen([sys.executable, "ocr_service.py"], 
                creationflags=subprocess.CREATE_NO_WINDOW)
```

### 自動起動設定
- **スタートアップフォルダ**: 非表示版ショートカットを自動生成
- **起動方式**: `wscript.exe run_ocr_hidden.vbs`
- **ウィンドウ状態**: 完全非表示（バックグラウンド常駐）

### 技術的メリット
1. **ユーザビリティ向上**: 作業画面にCMD画面が表示されない
2. **システムリソース効率化**: 不要なUI描画処理を排除
3. **安定性向上**: ユーザーの誤操作によるサービス停止を防止

【今後の開発ロードマップ】
Phase 1: インテリジェント分類機能

機械学習による自動カテゴリ分類
ユーザー辞書学習機能
LLM APIを活用したセカンドオピニオン分析

Phase 2: パーソナライゼーション機能

ログデータ分析による行動パターン可視化
AIエージェントによるタスク整理・優先度提案
目標設定支援・進捗管理機能

Phase 3: マルチプラットフォーム展開

モバイルアプリケーション開発
クラウドベース処理基盤構築
リアルタイム同期・共有機能

【技術的価値・差別化要素】

精度: 複数OCRエンジンの統合による高精度認識
UX: ワンクリック操作による圧倒的な操作性
拡張性: AIエージェント連携による付加価値創出
実用性: 実業務での継続使用を前提とした設計
- **🆕 バックグラウンド**: 非表示常駐による作業効率化

【開発成果・学習要素】

Windows APIとPythonの統合技術習得
画像処理・OCR技術の実装経験
ユーザビリティを重視したツール設計
自動化による業務効率化の定量的実証
- **🆕 プロセス管理**: バックグラウンドサービス制御技術

## 【ファイル構成】

- `working_ocr_service.py` - メインOCRサービス
- `hotkey_ocr.py` - ホットキー制御
- `run_ocr_hidden.vbs` - 非表示起動VBScript
- `run_ocr_hidden.ps1` - 非表示起動PowerShell
- `run_ocr_background.pyw` - Python非表示Wrapper
- `start_ocr_hidden.bat` - 簡易起動バッチ
