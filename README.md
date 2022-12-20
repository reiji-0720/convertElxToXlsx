# convertElxToXlsx
xlsファイルをxlsxファイルに変換
---

convert.pyを実行することで.xlsを.xlsxに変換する  
convert.pyで行いたいモジュールの関数を呼び出す。  
#### 可能な操作  
- convert.pyのあるディレクトリ内のxlsを変換する.(カレントディレクトリのみ下の階層までは見ない)  
- convert.pyのあるディレクトリ下全てのxlsを変換する。  

---
### 動作環境
- python 3.9.0
- mac book pro 13-inch,2017,Two Thunderbolt 3 ports
- macOS Monterey (version:12.3.1)

### ディレクトリ構造
```
.  
├── README.md  
├── convert.py  
├── libs  
│   ├── __pycache__  
│   │   └── convert_xls_to_xlsx.cpython-39.pyc  
│   └── convert_xls_to_xlsx.py  
└── sample.xls  
```
---

### モジュール  
```./libs/```に格納


#### convert_xls_to_xlsx.py  
基本的にこのファイルにある関数をconvert.pyで呼び出すだけ  

**関数**  

- convert_type_current()  
カレントディレクトリ内に存在するxlsファイルをxlsxファイルに変換する。  
拡張子の変更のみ可能。返還後のxlsxファイルはデータは保たれるがレイアウトが崩れる場合がある。  
xlsファイルを残した状態で新規にxlsxファイルを作成する。  


- simple_check()
カレントディレクトリ内に存在するxlsファイルをxlsxファイルに変換する。  
拡張子の変更のみ可能。返還後のxlsxファイルはデータは保たれるがレイアウトが崩れる場合がある。  
convert_type_current()よりもレイアウトは崩れない  
xlsファイルを残した状態で新規にxlsxファイルを作成する。  

- convertion()
一つのファイルに対して変換処理を行う  
simple_check()では再起的にconvertion()を呼び出している  

### モジュールの追加  
python実行前にCLIで以下のモジュールのインストールを行なってください。  
#### CLI
```
$ pip install pyexcel  
$ pip install pyexcel-xls  
$ pip install pyexcel-xlsx  
$ pip install glob  
$ pip install xlwings  

#########################
↓ あまり意味がなかったモジュール  
#########################
$ pip install openpyxl  
$ pip install xlrd  
$ pip install xlwt  
```
#### 備考
xlsxファイルを操作する時に使用するopenpyxlライブラリはxls形式をサポートしていないので意外と簡単に行かない  
EXCEL上でxlsファイルを開いて「上書き保存」でxlsxにすると基本的にレイアウトが崩れないので行おうとしたが不可能だった  
xls対応モジュールのxlwtはxlsx形式には対応していないので，一度開いて別名で保存する際にxlsxで保存ができない  
pyexcelを用いた変換では拡張子の変換をかけるだけでレイアウト崩れが生じる  

xlwingsライブラリはexcelのアプリケーションを実際にバックグラウンドで起動して操作を行う  
xlwingsで基本的に行おうと考えている  

xlwingsで可能だった  
※ アクセス権の付与必要

