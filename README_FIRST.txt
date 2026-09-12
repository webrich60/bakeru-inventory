BAKERU v2.7.3 FULL BACKUP / FOLLOW-UP PACKAGE
=================================================

このZIPは差分ファイルではありません。
BAKERU 商品・在庫管理 v2.7.3 を今後フォローアップ・復旧するための
最新完全ソース一式です。

【実際に更新するファイル】
1. GitHub Pages
   index.html をGitHubリポジトリの index.html に上書き

2. Google Apps Script
   Code.gs の全内容で既存の Code.gs を全置換して保存
   GASを変更した後は、既存Webアプリを「新しいバージョン」で再デプロイ
   （通常はWebアプリURLを変えない運用）

【v2.7.3の主な変更】
・領収書/納品書の読取を「標準OCR → ルール整理 → 商品マスター照合」に変更
・標準OCRだけならGemini APIキーなしで使用可能
・標準OCRはブラウザ側のTesseract.js（日本語+英語）を使用
・日付、金額、数量などをJavaScriptのルールで整理
・商品名をBAKERUの商品マスターと照合
・判定が難しい場合だけ、Gemini設定済みならOCR済み文字をAI補正
・領収書画像そのものはGeminiへ送らない
・Gemini未設定でもOCR結果を確認・修正し、商品マスターへ紐付けて登録可能
・OCR結果の生テキストと信頼度を確認可能
・明細の手動追加/修正/削除に対応

【Gemini Script Properties】
AI補正・AI分析を使う場合のみ設定してください。

プロパティ名: GEMINI_API_KEY
値: Gemini APIキー

任意:
プロパティ名: GEMINI_MODEL
値: 使用するGeminiモデル名

※ gemini_key ではなく GEMINI_API_KEY を使用します。

【標準OCRについて】
標準OCRはGeminiを使用しません。
初回だけOCRライブラリと日本語辞書の準備に時間がかかる場合があります。
同じ端末では2回目以降は軽くなることがあります。
OCRライブラリ等をCDNから取得するため、通常利用時はインターネット接続が必要です。

【既存機能は維持】
・商品マスター / 在庫管理
・入庫（仕入れ）
・販売
・使用（店内使用・製造・試飲等）
・入出庫履歴
・1次元バーコード読取
・BAKERU専用QRコード読取
・QRコード作成・印刷
  - 商品貼付用
  - 棚貼付用
  - ラミネート入出庫カード用
・QR印刷サイズ選択 / 枚数選択
・QR下に商品管理番号を表示
・商品IDと商品詳細の連結
・注意在庫 / レッド在庫の2段階アラート
・在庫回復時のアラート自動解決
・商品別Gmail通知
・棚卸
・テストモード
・Gemini AI補正 / AI分析（設定時のみ）
・NEKKO連携用JSON出力
・Google Sheetsを正本データとして同期

【QRの重要仕様】
QRの中には商品名、グラム数、仕入金額などを固定保存しません。
QRにはBAKERUの商品IDだけを紐付けます。

例:
BAKERU:P:<商品ID>

読み取ると、その商品IDを使ってGoogle Sheetsの商品マスターから
最新の商品名、内容量、仕入金額、販売金額、在庫、棚位置等を取得します。
そのため、同じ商品IDのまま商品詳細を変更してもQRはそのまま使えます。

【重要】
・このZIPにGoogle Sheets内の実データは含まれません。
・APIキー、同期トークン等の秘密情報は含まれません。
・Script PropertiesはGAS側で別管理です。
・古いZIPがなくても、このZIPの index.html と Code.gs から継続開発できます。

Version: 2.7.3
Frontend Build: 2.7.3-standard-ocr-ai-assist
Backend Version: 2.7.3-ocr-hybrid
