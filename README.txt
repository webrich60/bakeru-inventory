BAKERU 商品・在庫管理 v2.5.6 LocalStorage依存削減版

GitHubにこのフォルダー内のファイルを上書きアップロードしてください。
今回の修正:
- 商品マスタ・在庫・入出庫履歴・アラートをLocalStorageへ常駐保存しない
- Googleスプレッドシートを正式データ（正本）として利用
- 旧BAKERUの大容量キャッシュだけを起動時に自動削除
- 他のWEBRICHツールのLocalStorageは削除しない
- LocalStorageに残すのはGAS URL、同期トークン、未送信キュー、軽量な重複スキャン履歴のみ
- LocalStorage容量不足でも画面を停止させない
- Gemini既定モデルは gemini-3.6-flash
- noindex設定維持

Code.gs はGAS側のバックアップです。GitHub Pages上では実行されません。
今回GAS側の上書きは不要です。GitHub側は index.html の上書きだけでも修正できます。
