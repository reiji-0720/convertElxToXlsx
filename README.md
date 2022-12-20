# convertElxToXlsx
xlsファイルをxlsxファイルに変換
---

convert.pyを実行することでconvert.py以下のディレクトリの.xlsを.xlsxに変換する  
convert.pyで行いたいモジュールの関数を呼び出す。  
#### 可能な操作  
- convert.pyのあるディレクトリ内のxlsを変換する.(カレントディレクトリのみ下の階層までは見ない)  
- convert.pyのあるディレクトリ下全てのxlsを変換する。  

### モジュール  
'''./libs/'''に格納


#### convert_xls_to_xlsx.py  
++関数++  

- convert_type_current()  
カレントディレクトリ内に存在するxlsファイルをxlsxファイルに変換する。  
拡張子の変更のみ可能。返還後のxlsxファイルはデータは保たれるがレイアウトが崩れる場合がある。  
xlsファイルを残した状態で新規にxlsxファイルを作成する。  

### モジュールの追加  
> pip install pyexcel  
> pip install pyexcel-xls  
> pip install pyexcel-xlsx  
> pip install glob  
