import pyexcel as p
import glob
import os
import xlwings


## カレントディレクトリ内での拡張子変更 文字データのみ
def convert_type_current():
    it = glob.glob("*.xls")
    for xls in it:
        xlsx = "{}".format(xls) + "x"
        print(xlsx)
        p.save_book_as(file_name='{}'.format(xls), dest_file_name='{}'.format(xlsx))

## カレントディレクトリ内での拡張子変更 罫線データ含む (ディレクトリ下部探索)
def simple_check():
    i = 0
    for pathname,dirname, filenames in os.walk("./"):
        for filename in filenames:
            print('pathname:'+ pathname)
            pathname_stack = pathname # filename文字列の保有（そのままpathnameで渡すとpathがうまく行かない）
            if filename.endswith('.xls'): # フィルタ処理 csvのみ処理
                filenameTo = os.path.join(pathname_stack,filename) # ファイルの相対パス取得
                print('filename' + filename)
                convertion(filenameTo)
                i += 1

## 単一ファイルの拡張子変更 罫線データ含む
def convertion(filename):
    app = xlwings.App() # appの開始
    filename_only = os.path.splitext(filename)[0] # ファイル名拡張子なし
    wb = app.books.open(filename) # ファイルオープン
    print(filename)
    print(filename_only)
    filename_only = filename_only + '.xlsx'
    print('filename_only : '+ filename_only)
    wb.save(filename_only)
    wb.close()
    app.quit() # appの終了