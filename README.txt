BAKERU v2.5 検索インデックス対策版

GitHub Pagesへ配置する際：
1. BAKERU_v2.5_NoIndex.html を現在使用しているHTMLファイル名に合わせて配置してください。
   ルートで開く運用なら index.html に変更します。
2. robots.txt はリポジトリの公開ルート直下へ配置してください。
3. .nojekyll もルート直下へ配置してOKです。

対策内容：
- robots / googlebot / bingbot に noindex,nofollow,noarchive,nosnippet,noimageindex
- robots.txt は noindex を検索エンジンに認識させるためクロールを許可

重要：
- Google公式仕様では、robots.txt でページをクロール禁止にすると noindex を読めず、URLだけ検索結果に残る可能性があります。
- そのため「検索結果に載せない」ことを優先しています。
- URLそのものへのアクセスを禁止したい場合は、noindexではなくログイン・アクセス制御が必要です。


=== v2.5.1 GAS修正 ===
- setupBAKERU() の最後の待機型alertを廃止し、Apps Scriptエディタから実行しても終了するよう修正。
- Gemini既定モデルを gemini-3.6-flash に統一。
- 既存のGEMINI_MODELが gemini-3.8-flash の場合、setupBAKERU() 実行時に3.6へ自動修正。
- BAKERU管理メニューに「Gemini 3.6 Flashに設定」を追加。
