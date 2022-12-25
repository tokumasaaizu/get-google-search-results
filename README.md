# get-google-search-results
案件に取り組んだ時のソースコードです。
エクセルから抽出したキーワードをGoogle検索結果にかけます。
上位10サイト(検索結果1ページ目)のタイトル、URL、メタディスクリプショを抽出しました。
抽出した内容をxlsxに保存するソースコードです。

google-search-beautifulsoup.py
は、pythonのBeautifulsoupを使用して抽出しています(非推奨のやり方)

CustomSearchAPI.py
はGoogleのCustom Search APIを使用して抽出しています(推奨のやり方)

