BAKERU 商品・在庫管理 v2.5.2 GitHub Pages修正版

【今回の修正】
- index.html がReactソースのまま表示される問題を修正
- GitHub Pagesで開ける完全なHTMLに変更
- React / Tailwind / Font Awesome はCDNから読み込み
- Code.gs はGAS側で使用（GitHub上ではバックアップ）
- noindex設定を維持

【GitHubへアップするファイル】
index.html
Code.gs
robots.txt
.nojekyll
README.txt
SECURITY_NOTE.txt

【重要】
GitHub PagesのSourceは Deploy from a branch / main / /(root) を指定してください。
公開後、https://webrich60.github.io/bakeru-inventory/ を開いて確認します。
