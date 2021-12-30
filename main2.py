import os
import kiritori as ki
import kousa
import sys
import shutil

#実行の仕方　python main1.py 【PDFの名前（**.pdf）】【csvの名前（**.csv）】【base画像の名前（**.jpg）】

# メイン処理 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
if __name__ == "__main__":
    args = sys.argv

    # PDFの処理
    PDF_seito = args[1] #"seito.pdf"
    CSV_meibo = args[2] #"seito_meibo.csv"
     #もし、すでに存在していればpassまたは移動させる。
    if os.path.exists("./seitoPDF_CSV/"+PDF_seito) and os.path.exists("./seitoPDF_CSV/" + CSV_meibo):
    # PDFを画像に変換
      ki.seitoPDFchengeJPG("./seitoPDF_CSV/"+PDF_seito, "./seitoPDF_CSV/" + CSV_meibo)
    elif os.path.exists(PDF_seito) and os.path.exists(CSV_meibo):
    # PDFを画像に変換
      ki.seitoPDFchengeJPG(PDF_seito, CSV_meibo)
      shutil.move(PDF_seito, "./seitoPDF_CSV/")
      shutil.move(CSV_meibo, "./seitoPDF_CSV/")
    elif os.path.exists("./setting/output/name"):
      pass
    else:
        print("seito.pdfとseito_meibo.csvを用意してください。")
        sys.exit( )

#  切り取り開始
    kousa.allTrim()
# base画像名を入力
# 白紙は０点に移動させる。
    ki.zero_move(basefile=args[3]) # base-0-000.jpg
