# convertElxToXlsx
xlsファイルをxlsxファイルに変換
---

convert.pyを実行することで.xlsを.xlsxに変換する  
convert.pyで行いたいモジュールの関数を呼び出す。  
#### 可能な操作  
- convert.pyのあるディレクトリ内のxlsを変換する.(カレントディレクトリのみ下の階層までは見ない)  
- convert.pyのあるディレクトリ下全てのxlsを変換する。  

---

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

### モジュールの追加  
python実行前にCLIで以下のモジュールのインストールを行なってください。  
#### CLI
```
$ pip install pyexcel  
$ pip install pyexcel-xls  
$ pip install pyexcel-xlsx  
$ pip install glob  
```
