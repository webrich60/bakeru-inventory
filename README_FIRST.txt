BAKERU v2.7.2 FULL SOURCE / FOLLOW-UP PACKAGE
================================================

このZIPは「差分ファイル」ではありません。
BAKERU 商品・在庫管理 v2.7.2 を今後フォローアップ・復旧するための
最新完全ソース一式です。

【このZIPだけ残せばよいもの】
source/index.html  ：GitHub Pages側の完全な最新画面コード
source/Code.gs     ：Google Apps Script側の完全な最新バックエンドコード

古いv2.6.x / v2.7.0 / v2.7.1のZIPを持っていなくても、
現在のv2.7.2から開発を継続できます。

【現在の主な機能】
・商品マスター
・在庫管理
・入庫（仕入れ）
・販売
・使用（店内使用・製造・試飲等）
・入出庫履歴
・領収書OCR
・1次元バーコード読取
・BAKERU専用QRコード読取
・QRコード作成・印刷
  - 商品貼付用
  - 棚貼付用
  - ラミネート入出庫カード用
・QR印刷サイズ選択
・QR下に管理番号表示
・商品IDと商品詳細の連結
・内容量 / 単位 / 特徴 / 棚・保管場所
・注意在庫 / レッド在庫の2段階アラート
・在庫回復時のアラート自動解決
・商品別Gmail通知
・棚卸
・テストモード（T00～T07）
・Gemini API連携（gemini-3.6-flash）
・NEKKO連携用JSON出力
・Google Sheetsを正本データとして同期

【QRの重要仕様】
QRに商品名・グラム数・価格等は固定保存しません。
QRの中には BAKERUの商品識別用プレフィックス + 商品ID だけが入ります。

例：
BAKERU:P:prod_xxxxx

読み取り時に商品IDからGoogle Sheetsの商品マスターを検索し、
その時点の最新の商品名、内容量、仕入金額、販売金額、在庫、棚位置等を表示します。

したがって同じ商品IDのまま商品詳細を変更しても、
印刷済み・ラミネート済みのQRは作り直す必要がありません。

【更新・復旧時】
1. GitHub Pages
   source/index.html をリポジトリの index.html に上書き

2. Google Apps Script
   source/Code.gs の全内容で既存 Code.gs を全置換して保存

3. 必要に応じて setupBAKERU() を実行

4. GAS Webアプリは既存デプロイを「新しいバージョン」で更新
   URLは変えない運用

【重要】
・このZIPにはGoogle Sheets内の実データ自体は含まれません。
・Gemini APIキー、同期トークン等の秘密情報は含めていません。
  それらはGAS Script Properties側に保存されます。
・index.htmlはReact/Tailwind/Font Awesome/QR生成ライブラリをCDNから読み込みます。
  通常利用時はインターネット接続が必要です。

Version: 2.7.2
Build: 2.7.2-qr-center-label-print
