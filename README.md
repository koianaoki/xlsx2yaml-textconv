# xlsx2yaml-textconv

テーブルチックなxlsxファイルをシートごとにyaml形式で出力します

git用のtextconvツールとしても使用できます

## install
```
npm install -g xlsx2yaml-textconv
```

## USAGE


```bash
Usage: xlsx2yaml-textconv [options] <file>

Options:
  --field_row <n>                       A field_row
  --select_sheet_name <string(regexp)>  A select_sheet_name
```

以下のようなシートのエクセル(Book.xlsx)を例にしていきます

#### Sheet1

|name|age|gender|
|-|-|-|
|naoki|15|male|
|misae|17|female|

### example

```
xlsx2yaml-textconv Book.xlsx --field_row 0

Sheet1 name: naoki
       age: 15
       gender: male

Sheet1 name: misae
       age: 17
       gender: female
```


### options

#### field_row options

フィールド(ヘッダー)指定をします

```
xlsx2yaml-textconv Book.xlsx --field_row 1

Sheet1 '15': 17
       naoki: misae
       male: female
```

- 指定がない場合はxlsxに従いA ~ 表記から続いていきます
- また、空の列は基本的には詰めていきます

```
xlsx2yaml-textconv Book.xlsx

Sheet1 A: name
       B: age
       C: gender

Sheet1 A: naoki
       B: '15'
       C: male

Sheet1 A: misae
       B: '17'
       C: female
```

#### select_sheet_name

正規表現形式でシート名を指定します

上記例のxlsxに以下のようなシートを追加

#### Sheet12

|name|age|gender|
|-|-|-|
|naoya|18|male|
|misako|20|female|

```
xlsx2yaml-textconv Book.xlsx --field_row 0 --select_sheet_name 'Sheet*'

Sheet1 name: naoki
       age: 15
       gender: male

Sheet1 name: misae
       age: 17
       gender: female

Sheet12 name: naoya
        age: 18
        gender: male

Sheet12 name: misako
        age: 20
        gender: female

```

## for git

- textconv設定をかけることによってgit差分を上記表示に従って見れるようになります
- diffmergeを使用して差分を確認することも可能


.gitattributes
```
*.xlsx diff=xlsx
*.XLSX diff=xlsx
```

git config
```
git config --global diff.xlsx.binary true
git config --global diff.xlsx.textconv "Book.xlsx --field_row 0"
git config --global diff.xlsx.cachetextconv true
```

外部ツールを使用する場合
```
git config --global difftool.diffmerge.cmd "difftool diffmerge \$LOCAL \$REMOTE"
git config --global difftool.diffmerge.cmd "xlsx-difftool diffmerge \$LOCAL \$REMOTE"
git config --global difftool.sourcetree.cmd "xlsx-difftool diffmerge \$LOCAL \$REMOTE"
git config --global diff.tool diffmerge
```

## ps

[こちら](https://github.com/pismute/node-textconv)参考にさせて頂きました。感謝いたします


## License

Copyright 2020+ koian Licensed under the MIT license.
