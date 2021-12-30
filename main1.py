import os
import sys
import pandas as pd
import shutil
import kiritori as ki
import kousa

#実行の仕方　python main1.py 【PDFの名前（**.pdf）】


# メイン処理 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
if __name__ == "__main__":
    args = sys.argv
    # 初期設定
    if not os.path.exists("setting"):
        os.makedirs("setting")
    if not os.path.exists("./setting/input"):
        os.makedirs("./setting/input")
    if not os.path.exists("./setting/output"):
        os.makedirs("./setting/output")
    if not os.path.exists("./setting"):
        os.makedirs("setting")
    if not os.path.exists("seitoPDF_CSV"):
        os.makedirs("seitoPDF_CSV")
    if not os.path.exists("base_white_mohan"):
        os.makedirs("base_white_mohan")
    if not os.path.exists('./setting/ini.csv') and not os.path.exists('./setting/trimData.csv'):
        ini=pd.DataFrame(columns=['tag', 'start_x', 'start_y', 'end_x', 'end_y'])
        ini.to_csv('./setting/ini.csv',index=False)
    if not os.path.exists("JPG_LISTS"):
        os.makedirs("JPG_LISTS")

    PDF_Wname = args[1]
    #もし、すでに画像が存在していればpassまたは移動させる。
    if os.path.exists(PDF_Wname):
        ki.WhitePDFchengeJPG(PDF_Wname)
        shutil.move(PDF_Wname, "./base_white_mohan/")
    # もし、
    elif os.path.exists("./base_white_mohan/" + PDF_Wname):
        if os.path.exists("./base_white_mohan/white.jpg"):
            pass
        else:
            ki.WhitePDFchengeJPG("./base_white_mohan/" + PDF_Wname)
    else:
        print("white.pdfまたはwhite.jpgを用意してください。")
        sys.exit( )
    # 名前のみを切り取り
    kousa.input_ck()
    print(pd.read_csv("./setting/trimData.csv"))
    # 画像処理開始
    # 名前リスト作成。解答の切り取りに不備があれば画像で確認
    namelist = ki.name_lists("./base_white_mohan/white.jpg",square=4000)
    # 順不同座標をdataframeに保存
    df = ki.jidoukiritori("./base_white_mohan/white.jpg")
    # 二段ヴァージョンに座標リストを並び替え（左上から右下へ。横並びに）
    df_sort = ki.sortdata_2dan_yoko(df, "./base_white_mohan/white.jpg", namelist)
    # trimDataに結合してセーブ
    ki.trimdata_save(df_sort)