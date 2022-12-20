import pyexcel as p
import glob
import os
import xlwings


## カレントディレクトリ内での拡張子変更
def convert_type_current():
    it = glob.glob("*.xls")
    for xls in it:
        xlsx = "{}".format(xls) + "x"
        print(xlsx)
        p.save_book_as(file_name='{}'.format(xls), dest_file_name='{}'.format(xlsx))


def simple_check():
    i = 0
    for pathname,dirname, filenames in os.walk("./"):
        for filename in filenames:
            if filename.endswith('.xls'): # フィルタ処理 csvのみ処理
                filename = os.path.join(*pathname,filename) # ファイルの相対パス取得

                convertion(filename)

                # app = xlwings.App() # appの開始
                # # wb = app.books.open(filename) # ファイルオープン
                # wb = app.books.open('.'+filename) # ファイルオープン

                # print(filename)
                # print(filename_only)
                # filename_only = '.'+ filename_only + '.xlsx'
                # print('filename_only : '+ filename_only)
                # wb.save(filename_only)
                # wb.close()
                # app.quit() # appの終了
                i += 1


def convertion(filename):
    app = xlwings.App() # appの開始
    filename_only = os.path.splitext(filename)[0] # ファイル名拡張子なし
    wb = app.books.open('.'+filename) # ファイルオープン
    print(filename)
    print(filename_only)
    filename_only = '.'+ filename_only + '.xlsx'
    print('filename_only : '+ filename_only)
    wb.save(filename_only)
    wb.close()
    app.quit() # appの終了