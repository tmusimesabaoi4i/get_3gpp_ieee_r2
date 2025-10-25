# get_3gpp_ieee_r2


# How To PUSH

```
git add ./
```

and

```
git commit --allow-empty -m "　"
```

and

```
git push
```

# main files

1. main_build_doclist.py

RAWリスト → 正規化した「ドキュメントリスト」を生成

1. main_build_matched_list.py

ドキュメントリスト → 条件一致リストを生成

1. main_fetch_and_convert.py

条件一致リスト → ダウンロード → ZIP解凍 → HTML化 → PDF化

1. main_make_citation.py

文献番号を指定 → 引用形式を出力