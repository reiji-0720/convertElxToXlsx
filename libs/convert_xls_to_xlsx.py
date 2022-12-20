import pyexcel as p
import glob


## カレントディレクトリ内での拡張子変更
def convert_type_current():
    it = glob.glob("*.xls")
    for xls in it:
        xlsx = "{}".format(xls) + "x"
        print(xlsx)
        p.save_book_as(file_name='{}'.format(xls), dest_file_name='{}'.format(xlsx))
