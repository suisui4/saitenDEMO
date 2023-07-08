import cv2
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import os
import glob
import fitz
# from pathlib import Path
# from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageDraw2
import shutil
# from keras.models import load_model
# from keras.models import model_from_json


# poppler/binを環境変数PATHに追加する
# poppler_dir = "C:/poppler/bin"
# os.environ["PATH"] += os.pathsep + str(poppler_dir)


def name_lists(file,floor =10000, square=4000):
    # 大問〇の△問目のリスト
    namelist=[]
    # 画像の読み込み、加工
    img = cv2.imread(file)
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # img_gray = cv2.fastNlMeansDenoising(img_gray, h=20) # ノイズがひどい場合
    img_adth = cv2.adaptiveThreshold(img_gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 43, 3)
    img_adth_er = cv2.bitwise_not(img_adth) # 白黒反転

    # 図形の検出（解答欄の番号や文字も検出します。）
    contours, _ = cv2.findContours(img_adth_er, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    # 解答欄の面積を割り出す。また、大問毎の座標４つをzahyoulistに格納。それぞれの座標も格納する。

    # 最大
    ceil = 10000000
    # まとめ座標
    zahyoulist=[]
    rect_big=[]
    for i, rect in enumerate(contours):
    # 指定範囲内(ceil未満、floorより上)の面積(ピクセル)のみをカウントさせる。
        if cv2.contourArea(rect) > floor and ceil > cv2.contourArea(rect):
    # x座標, y座標, 横の長さw,縦の長さhの情報をcv2.boundingRect(rect)で取得する。
            x, y, w, h = cv2.boundingRect(rect)
            x_end=x+w
            y_end=y+h
            zahyoulist.append([x, x_end, y, y_end])
            rectbig = cv2.minAreaRect(rect)
            rect_big_points = cv2.boxPoints(rectbig).astype(int)
            rect_big.append(rect_big_points)
    # 左上から右下順になるようにソート
    rect_big = sorted(rect_big, key=lambda x: (x[1][1], x[0][0]))
    # 描画する。
    img = cv2.imread(file)
    for i, rect in enumerate(rect_big):
        color = np.random.randint(0, 255, 3).tolist()
        cv2.drawContours(img, rect_big, i, color, 2)
        cv2.putText(img, str(i+1), tuple(rect[0]), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 3)
    fig = plt.figure(figsize=(7,10), dpi=100)
    # plt.imshow(img)
#     print("大問の切り取り状況を保存しました。(daimon.jpg)")
#     plt.savefig("./daimon.jpg")
    # x座標に着目し、画像の半分以下のものと、半分より大きいものに分けて別のリストに格納する
    # 分けたものをくっつける。
    left_big=[]
    right_big=[]
    for i in range(len(zahyoulist)):
        if zahyoulist[i][0] < img.shape[1]//2:
            left_big.append(zahyoulist[i])
            left_big.sort(key = lambda y: y[2], reverse=False)

    for i in range(len(zahyoulist)):
        if zahyoulist[i][0] > img.shape[1]//2:
            right_big.append(zahyoulist[i])
            right_big.sort(key = lambda Z: Z[2], reverse=False)
    #左半分に右半分を付ける
    for i in range(len(right_big)):
        left_big.append(right_big[i])
    # 大問を切り分けて保存
#     n = str(input("大問画像と小問画像を保存しますか？保存するなら【yes】と入力"))
    img = cv2.imread(file)
    for i in range(len(left_big)):
        x, x_end, t, t_end = left_big[i]
    #このimg_tissueが切った大問画像
        img_tissue = img[t:t_end, x:x_end]
    # 画像保存処理。
#         if n =="yes":
#             out_dir_suf = f"CUT_NO{i+1}"
#             out_dir = str(out_dir_suf)
#             if not os.path.exists(out_dir):
#                 os.makedirs(out_dir)
#             cv2.imwrite(out_dir + f"\\base_NO{i+1}.jpg",img_tissue)
    # ここからは、切り取った大問を小問にして処理していく。
    # 画像処理
        gray = cv2.cvtColor(img_tissue, cv2.COLOR_BGR2GRAY)
        ret, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    # 膨張処理
        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
        binary = cv2.dilate(binary, kernel)
    # 図形の検出
        contours, hierarchy = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
        rects = []
    # 面積で場合分け
        for cnt, hrchy in zip(contours, hierarchy[0]):
            if cv2.contourArea(cnt) < square:
                continue  # 面積が小さいものは除く
            if hrchy[3] == -1:
                continue  # ルートノードは除く
    # 輪郭を囲む長方形を計算する。
            rect = cv2.minAreaRect(cnt)
            rect_points = cv2.boxPoints(rect).astype(int)
            rects.append(rect_points)
    # 左上から右下順になるようにソート
        rects = sorted(rects, key=lambda x: (x[1][1], x[0][0]))
        print(f"大問{i+1}の小問の数{len(rects)}")
    # 切り取り処理
        for num in range(len(rects)):
            namelist.append(f"NO{str(i+1).zfill(2)}-{str(num+1).zfill(2)}")
            x1=min(rects[num].T[0])
            x2=max(rects[num].T[0])
            y1=min(rects[num].T[1])
            y2=max(rects[num].T[1])
    # このimg_tissue2が切った小問画像
            img_tissue_2 = img_tissue[y1:y2, x1:x2]
    # 保存処理
#             if n=="yes":
#                 out_dir_suf2 = f"\\small_question"
#                 out_dir2 = str(out_dir + out_dir_suf2)
#                 if not os.path.exists(out_dir2):
#                     os.makedirs(out_dir2)
#                 cv2.imwrite(out_dir2 + f"\\base_no{ str(num+1).zfill(2) }.jpg",img_tissue_2)
    print(f"大問数は{len(left_big)}、小問の数は{len(namelist)}です。")
    return namelist


def jidoukiritori(file, square=4000):
    # 小問画像の座標を入れるリストを作成
    x1_list=[]
    x2_list=[]
    y1_list=[]
    y2_list=[]
    # 画像の編集
    img = cv2.imread(file)
    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    ret, binary = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    # 膨張処理
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (5, 5))
    binary = cv2.dilate(binary, kernel)
    # 図形の検出
    contours, hierarchy = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    rect_small = []
    # 面積で場合分け
    for cnt, hrchy in zip(contours, hierarchy[0]):
        if cv2.contourArea(cnt) < square:
            continue  # 面積が小さいものは除く
        if hrchy[3] == -1:
            continue  # ルートノードは除く
        rectsmall = cv2.minAreaRect(cnt)
        rect_smallpoints = cv2.boxPoints(rectsmall).astype(int)
        rect_small.append(rect_smallpoints)
    for num in range(len(rect_small)):
        x1=min(rect_small[num].T[0])
        x2=max(rect_small[num].T[0])
        y1=min(rect_small[num].T[1])
        y2=max(rect_small[num].T[1])
        x1_list.append(x1)
        x2_list.append(x2)
        y1_list.append(y1)
        y2_list.append(y2)
    lists = [x1_list, y1_list, x2_list, y2_list]
    df = pd.DataFrame(lists, index=["start_x","start_y", "end_x","end_y"]).T

    return df

def sortdata_1dan_yoko(df, file, namelist, gosa_tate=10):
    df1_sort = df.sort_values(["end_y","start_x"])
    df1_sort["立幅フラグ"]=0
    label = 1
    for i in range(df1_sort.shape[0]):
        gosa = df1_sort[["end_y"]].iloc[i]-df1_sort[["end_y"]].iloc[i-1]
        if gosa[0] < gosa_tate:
            df1_sort["立幅フラグ"].iloc[i] = label
        else:
            label+=1
            df1_sort["立幅フラグ"].iloc[i] = label
    df3 = df1_sort.sort_values(["立幅フラグ","start_x"]).reset_index(drop=True)
    # 描画する。
    img = cv2.imread(file)
    new_rect=[]
    for i in range(len(df3.index)):
        new_rect.append(np.array([[df3.iloc[i][0],df3.iloc[i][3]],
                                    [df3.iloc[i][0],df3.iloc[i][1]],
                                    [df3.iloc[i][2],df3.iloc[i][1]],
                                    [df3.iloc[i][2],df3.iloc[i][3]]]))
    for i, rect in enumerate(new_rect):
        color = np.random.randint(0, 255, 3).tolist()
        cv2.drawContours(img, new_rect, i, color, 2)
        cv2.putText(img, str(i+1), tuple(rect[0]), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 3)
    fig = plt.figure(figsize=(7,10), dpi=100)
    plt.imshow(img)
    plt.show()
    plt.savefig("./ANS_KENSYUTU.jpg")
    print("画像のように切りとります。(ANS_KENSYUTU.jpg)よければ次に進んでください。")
    return df3


def sortdata_2dan_yoko(df, file, namelist, gosa_tate=10):
    img = cv2.imread(file)
    # 二段に分ける
    df.loc[df['end_x'] < img.shape[1]/2, 'label'] = 'left'
    df.loc[df['end_x'] > img.shape[1]/2, 'label'] = 'right'

    # 左半分
    df1=df[df["label"]=="left"]
    # yの値で高さを小さい順にした後、xの値で小さい順に
    df1_sort=df1.sort_values(["end_y","start_x"])
    # ただし、yの値でsortしても誤差があるので、その部分は許容するようにラベル付けをしてgosa_tateで調整する。
    df1_sort["立幅フラグ"]=0
    label = 1
    for i in range(df1_sort.shape[0]):
        gosa = df1_sort[["end_y"]].iloc[i]-df1_sort[["end_y"]].iloc[i-1]
        if gosa[0] < gosa_tate:
            df1_sort.iloc[i,5] = label
        else:
            label+=1
            df1_sort.iloc[i,5] = label
    df1_sort=df1_sort.sort_values(["立幅フラグ","start_x"])

    # 右半分もどうように行う。
    df2=df[df["label"]=="right"]
    df2_sort=df2.sort_values(["end_y","start_x"])

    # sortしても、誤差が１～５あるので調整
    df2_sort["立幅フラグ"]=0
    label = 1
    for i in range(df2_sort.shape[0]):
        gosa = df2_sort[["end_y"]].iloc[i]-df2_sort[["end_y"]].iloc[i-1]
        if gosa[0] < gosa_tate:
            df2_sort.iloc[i,5] = label
        else:
            label+=1
            df2_sort.iloc[i,5] = label
    df2_sort=df2_sort.sort_values(["立幅フラグ","start_x"])
    df3=pd.concat((df1_sort,df2_sort),axis=0).drop(columns="label").reset_index(drop=True)
    NO = df3.index.tolist()
    NEW=[]
    for i, j in zip(namelist,NO) :
        NEW.append(f"{i}_{str(j+1).zfill(3)}")
    df3.index=NEW
    df3 = df3.drop(["立幅フラグ"],axis=1)
    # 描画する。
    img = cv2.imread(file)
    new_rect=[]
    for i in range(len(df3.index)):
        new_rect.append(np.array([[df3.iloc[i][0],df3.iloc[i][3]],
                                [df3.iloc[i][0],df3.iloc[i][1]],
                                [df3.iloc[i][2],df3.iloc[i][1]],
                                [df3.iloc[i][2],df3.iloc[i][3]]]))
    for i, rect in enumerate(new_rect):
        color = np.random.randint(0, 255, 3).tolist()
        cv2.drawContours(img, new_rect, i, color, 2)
        cv2.putText(img, str(i+1), tuple(rect[0]), cv2.FONT_HERSHEY_SIMPLEX, 0.8, color, 3)
    fig = plt.figure(figsize=(7,10), dpi=100)
    plt.imshow(img)
#     plt.savefig("./ANS_KENSYUTU.jpg")
    plt.show()
#     print("画像のように切りとります。(ANS_KENSYUTU.jpg)よければ次に進んでください。")
    return df3

def trimdata_save(df):
    file_name="./setting/trimData.csv"
    df_trim= pd.read_csv(file_name, index_col=0)
    df_newtrim = pd.concat((df_trim[:1], df), axis=0)
    return df_newtrim.to_csv("./setting/trimData.csv", encoding="utf-8", index=True)



def WMPDFchengeJPG(PDF_WMname,dpi=150):
    # PDF -> Image に変換 dpi
    with fitz.open(PDF_WMname) as doc:
        image = doc[0].get_pixmap(dpi=dpi)
        image.save(f"./base_white_mohan/white.png") 
        image.save(f"./setting/input/white.png")
        print("白紙用紙を保存")
        print("base_white_mohanフォルダに保存しました。")
        print("./setting/inputフォルダに保存しました。")
        image = doc[1].get_pixmap(dpi=dpi)
        image.save(f"./base_white_mohan/mohan.png")
        image.save(f"./setting/input/mohan.png")
        print("模範解答用紙を保存")
        print("base_white_mohanフォルダに保存しました。")
        print("./setting/inputフォルダに保存しました。")    


def PDFchangeJPG(PDF_name, dpi=150):
    # PDF -> Image に変換 dpi
    with fitz.open(PDF_name) as doc:
        for i , page in enumerate(doc):
            image = page.get_pixmap(dpi=dpi)
            file = f"change_{str(i).zfill(4)}" + ".png"
            image_path = "./setting/input/" + file
            image.save(str(image_path), "PNG")
            
            # 画像ファイルを１ページずつ保存
        print("./setting/inputフォルダに保存しました。")

def seitoPDFchengeJPG(PDF_seito, CSV_meibo, dpi=150):
    df_seito_meibo=pd.read_csv(CSV_meibo)
    df_seito_meibo=df_seito_meibo.astype(str)
    df_seito_meibo["番号"]=df_seito_meibo["番号"].apply(lambda x: x.zfill(3))
    df_seito_meibo["連結"] = df_seito_meibo["年"]+"-"+df_seito_meibo["組"]+"-"+(df_seito_meibo["番号"])
    meibo=df_seito_meibo["連結"].to_list()
    # PDF -> Image に変換 dpi
    with fitz.open(PDF_name) as doc:
        if len(meibo)==len(doc):
            print(f"名簿と画像の枚数({len(meibo)}枚)が一致しています。inputフォルダにすべて保存します。")
            # 画像ファイルを１ページずつ保存
            for i, page in enumerate(doc):
                file_seito = f"{meibo[i]}" + ".png"
                PATH_input = "./setting/input"
                image_path = PATH_input +"\\"+ file_seito
            # JPEGで保存
                page.save(str(image_path), "PNG")
                if (i+1) % 5==0:
                    print(f"{i+1}枚目保存しました。")
            print("保存が完了しました。")
        else:
            print("要確認:名簿とPDFの枚数が一致していません。")
            print(f"名簿の人数{len(meibo)}人")
            print(f"PDFの枚数{len(pages)}")

# 画像周りを白で塗る
def line_square(img, bold=25):
    img_line2 = cv2.line(img, (0,0), (0,img.shape[0]), (255,255,255), bold)# 左縦
    img_line2 = cv2.line(img_line2, (0,0), (img.shape[1],0), (255,255,255), bold)# 上横
    img_line2 = cv2.line(img_line2, (img.shape[1],0), (img.shape[1],img.shape[0]), (255,255,255), bold) # 右縦
    img_line2 = cv2.line(img_line2, (0,img.shape[0]), (img.shape[1],img.shape[0]), (255,255,255), bold) # 下横
    return img_line2

def Kensyutu(file, floor=100):
    rects = []
    x1=[]
    x2=[]
    y1=[]
    y2=[]
    img = cv2.imread(file)
    img = line_square(img, bold=10)
    img_g = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.fastNlMeansDenoising(img_g, h=5)
    # 膨張処理
    ret3,th3 = cv2.threshold(gray,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    img_adth_er = cv2.erode(th3, None, iterations = 7)
    img_adth_er = cv2.dilate(img_adth_er, None, iterations = 5)
    img_adth_er_re = cv2.bitwise_not(img_adth_er)
    contours, _ = cv2.findContours(img_adth_er_re, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for i,contour in enumerate(contours):
        if cv2.contourArea(contour) >floor:
            print(cv2.contourArea(contour))
            rect = cv2.minAreaRect(contour)
            rect_points = cv2.boxPoints(rect).astype(int)
            rects.append(rect_points)
    # 左上から右下順になるようにソート
    rects = sorted(rects, key=lambda x: (x[1][1], x[0][0]))
    for i in rects:
        x1.append(min(i.T[0]))
        x2.append(max(i.T[0]))
        y1.append(min(i.T[1]))
        y2.append(max(i.T[1]))
            # このimg_tissue2が切った小問画像
    return x1,x2,y1,y2,gray

        
def expand(img):
# 膨張処理(直線を見つけやすくする)
    img_expand = cv2.erode(img, None, iterations = 1)
    img_expand = cv2.dilate(img_expand, None, iterations = 0)      
    return img_expand
### 全体画像から直線を検出し、塗りつぶす
def Draw_Line_img(img,min_point=85):
    # ノイズ処理
    img = img
    img_blur = cv2.medianBlur(img, 3)
    img_expand = expand(img_blur)
    # ２値化して白と黒の反転処理
    img_gray = cv2.cvtColor(img_expand, cv2.COLOR_BGR2GRAY)
    ret3,th3 = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    img_bit = cv2.bitwise_not(th3) 
    # 直線を検出して塗る
    lines = cv2.HoughLinesP(img_bit, rho=1, theta=np.pi/360, threshold=10, minLineLength=min_point, maxLineGap=5.5)    
    if (np.all(lines == None)):
        img_line = img
    else:
        for line in lines:
            x1, y1, x2, y2 = line[0]
            img_line = cv2.line(img, (x1,y1), (x2,y2), (255,255,255), 5)
    return img_line

def Count_square(img_adth_er_re, floor):
    img_shape_y = img_adth_er_re.shape[0]
    img_shape_x = img_adth_er_re.shape[1]
    x1=[]
    x2=[]
    y1=[]
    y2=[]
    # 面積を計算
    contours, _ = cv2.findContours(img_adth_er_re, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for rect in contours:
#         print(cv2.contourArea(rect))
        if cv2.contourArea(rect) > floor:
            x, y, w, h = cv2.boundingRect(rect)
            if x!=0:
                x1.append(x)
                x2.append(x+w)
                y1.append(y)
                y2.append(y+h)
            else:
                continue
    if x1==[]:
        X1,X2,Y1,Y2 = "", "", "", ""        
    else:    
        X1=min(x1)
        X2=max(x2)
        Y1=min(y1)
        Y2=max(y2)
        if X2-X1 >img_shape_x/2 and Y2-Y1 < img_shape_y/3:
            X1,X2,Y1,Y2 = "", "", "", ""   
        else:
            X1 = 0 if X1-5 < 0 else X1-5
            Y1 = 0 if Y1 < 5 else Y1-5
            Y2 = img_shape_y if (img_shape_y - Y2) < 5 else Y2 + 5
    return X1, X2, Y1, Y2
def Noise_draw(img, floor_max=100000):
    img = img
    img_blur = cv2.medianBlur(img, 3)
    img_gray = cv2.cvtColor(expand(img_blur), cv2.COLOR_BGR2GRAY)
    # 適応閾値を決める。15は今後も調整していく(あげると検出されやすくなる)
    img_ada = cv2.adaptiveThreshold(img_gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 20)
    img_adth_er_re = cv2.bitwise_not(img_ada)
    x1=[]
    x2=[]
    y1=[]
    y2=[]
    # 面積を計算
    contours, _ = cv2.findContours(img_adth_er_re, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    for rect in contours:
#         print(cv2.contourArea(rect))
        if cv2.contourArea(rect) < floor_max:
            x, y, w, h = cv2.boundingRect(rect)
            if x!=0:
                x1.append(x)
                x2.append(x+w)
                y1.append(y)
                y2.append(y+h)
            else:
                continue
            for X1,X2,Y1,Y2 in zip(x1,x2,y1,y2):
                cv2.rectangle(img, (X1, Y1), (X2, Y2), (255, 255, 255), -1)
    return img

def BW_trim_image(img):
# 画像データを二つ用意
    img_tissue_on = np.ones((100,100),np.uint8)*0
    img_tissue_off = np.ones((100,100),np.uint8)*255
    # 画像読み込み
#     img = cv2.imread(file)
    # 画像の大きさによって検出面積、膨張度合いを変える
    if img.shape[0] < 50 and img.shape[1] < 160:
        floor = 50
        Max_blur = 5
    else:
        floor = 600
        Max_blur = 10
### 全体画像から直線を検出し、塗りつぶす
    img_line = Draw_Line_img(img)
### ここから数字や文字の検出処理開始
    #細かいノイズを白塗る
    img_noise = Noise_draw(img_line, floor_max=15)
    img_blur = cv2.medianBlur(img_noise, 3)
    img_gray = cv2.cvtColor(expand(img_blur), cv2.COLOR_BGR2GRAY)
    # 適応閾値を決める。15は今後も調整していく(あげると検出されやすくなる)
    img_ada = cv2.adaptiveThreshold(img_gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 15, 20)
    # 周りを四角で囲む
    img_line2 = line_square(img_ada)
    # 膨張処理
    img_adth_er = cv2.erode(img_line2, None, iterations = Max_blur)
    img_adth_er = cv2.dilate(img_adth_er, None, iterations = 1)
    img_adth_er_re = cv2.bitwise_not(img_adth_er) 
    # 面積から検出
    X1, X2, Y1, Y2 = Count_square(img_adth_er_re, floor)
    # 画像を切り取る
    if X1 != "":
        img_tissue = img[Y1:Y2, X1:X2]
        # 直線を塗る
#         img_tissue_line = Noise_draw(img_tissue, floor_max=5)
#         img_tissue_line = Draw_Line_img(img_tissue,min_point=85)
        # グレースケール
        img_tissue_on = cv2.cvtColor(img_tissue, cv2.COLOR_BGR2GRAY)
        # 白黒反転
        _ ,img_tissue_off = cv2.threshold(img_tissue_on, 0, 255, cv2.THRESH_BINARY_INV+cv2.THRESH_OTSU)
#         img_tissue_on = cv2.bitwise_not(img_tissue_off)
#         cv2.rectangle(img_tissue_on, (X1, Y1), (X2 , Y2), (255, 255, 255), 5)
#         cv2.rectangle(img, (X1, Y1), (X2 + 5, Y2 + 5), (0, 0, 0), 1)
    return img_tissue_on, img_tissue_off
        

def summary_table(df_shomon, ginou, shikou):
    ginou_list = df_shomon.iloc[:,ginou].sum(axis="columns").tolist()
    shikou_list = df_shomon.iloc[:,shikou].sum(axis="columns").tolist()
    kanten = pd.DataFrame(data=[ginou_list,shikou_list], index=["ginou","shikou"]).T
    nyuryoku=pd.read_csv("./seitoPDF_CSV/seito_meibo.csv")
    nyuryoku["合計"] = df_shomon.iloc[:,1:].sum(axis="columns")
    nyuryoku["技能"] = kanten["ginou"]
    nyuryoku["思考"] = kanten["shikou"]
    kanten.to_csv("./setting/kanten.csv",index=False)
    nyuryoku.to_excel("./kousa_nyuryoku.xlsx", index=False, encoding="SHIFT-JIS")
    print("観点別ファイル(kanten.csv)を出力しました。採点の集計を行ってください。入力用ファイル（kousa_nyuryoku.xlsx）も出力しました。素点をコピペしてください。")
    df_summary = pd.concat([nyuryoku, df_shomon.iloc[:,1:]],axis=1)
    df_summary.to_csv("./seitoPDF_CSV/kousa_summary.csv", index=False, encoding="SHIFT-JIS")
    return df_summary

def Saiten_mark():
    if os.path.exists("./setting/kaitoYousi/saiten"):
        shutil.rmtree("./setting/kaitoYousi/saiten")
    df_zahyo = pd.read_csv("./setting/trimData.csv", index_col=0)
    Jpg_list = os.listdir("./setting/kaitoYousi")
    daimon_list = df_zahyo.index[1:-2]
    df_zahyo = df_zahyo.T

    for jpg in Jpg_list:
    # 画像を読み込む
        img = cv2.imread("./setting/kaitoYousi/" + jpg)
    # 問題番号リストで回す
        for daimon in daimon_list:
    # 問題番号の座標を取得
            x_s,y_s,x_g,y_g=df_zahyo[daimon]
            x= round(x_s+(x_g-x_s)/2)
            y=round(y_s+(y_g-y_s)/2)
    # 大きさによって〇のサイズを変える
            if x_g-x_s < y_g-y_s:
                size = (x_g-x_s)/3
            elif y_g-y_s < x_g-x_s:
                size = (y_g-y_s)/3
    # 大問フォルダの中の配点フォルダ名を取得
            haiten_list=os.listdir("./setting/output/"+daimon)
    # 0点フォルダは最初
            haiten_0 = haiten_list[0]
    # 0点フォルダのpass
            img_path_0 = daimon +"/"+ haiten_0 + "/" + jpg
    # バツを付ける
            if os.path.exists("./setting/output/" + img_path_0):
                img = cv2.drawMarker(img, (x, y), (255, 0, 0), thickness=8, markerType=cv2.MARKER_TILTED_CROSS, markerSize=int(size))
            else:
                pass
    # 正解フォルダは最後
            haiten_cor = haiten_list[-1]
    # 正解フォルダのpass
            img_path_cor = daimon +"/"+ haiten_cor + "/" + jpg
    # 丸を付ける
            if os.path.exists("./setting/output/" + img_path_cor):
                img = cv2.circle(img, (x, y), int(size), (0, 0, 255), thickness=3, lineType=cv2.LINE_AA)
            else:
                pass
    # もし配点フォルダが２つなら、○×のみなのでpassする。
            if len(haiten_list) == 2:
                pass
            else:
                haiten_bubun = haiten_list[1:-1]
                for bubun in haiten_bubun:
                    img_path_bubun = daimon +"/"+ bubun + "/" + jpg
    # 三角を付ける
                    if os.path.exists("./setting/output/" + img_path_bubun):
                        img = cv2.drawMarker(img, (x, y), (0, 255, 0), thickness=3, markerType=cv2.MARKER_TRIANGLE_UP, markerSize=int(size))
                    else:
                        pass
    # セーブする
        if not os.path.exists("./setting/kaitoYousi/saiten"):
            os.makedirs("./setting/kaitoYousi/saiten")
        cv2.imwrite("./setting/kaitoYousi/saiten"+"/"+ jpg, img)
        
def merge(files, shape, num_w, canvas, w, margin_w, h, margin_h): 
    for i in range(len(files)):
        img = Image.open(files[i])
        img = img.resize(shape)
        img = img.convert('RGB')
        draw = ImageDraw.Draw(img)
        draw.line((img.size[0], 0, img.size[0], img.size[1]), fill=(0, 0, 0), width=3)
        draw.line((0, img.size[1], img.size[0], img.size[1]), fill=(0, 0, 0), width=3)
        # 切り上げ計算
        #m = math.ceil((i+1) / num_w)   # math関数を使って切り上げを行う場合
        m = -(-(i+1) // num_w)          # 処理速度優先の場合(おすすめ)
        if i < num_w:
            canvas.paste(img, ((w + margin_w) * i, (h + margin_h) * (m-1)))
        else:
            canvas.paste(img, ((w + margin_w) * (i % num_w), (h + margin_h) * (m-1)))
    # 画像を表示
    # canvas.show()
    # 画像の縮小率(1は縮小無し。リサイズしたいときは#を除去し縮小率を設定する)
    # img_size = 10
    # 画像リサイズ(リサイズしたいときは#を除去する)
    # canvas = canvas.resize((w//int(img_size), h//int(img_size)), Image.LANCZOS)
    return canvas

def merge_summary(folda, num_w, num_h, shape):
    path_white = "./deep/" + folda + "/white_image_test/"
    # 白切り抜き画像を取得し、マージしていく
    path_deep_white_file_list = glob.glob(path_white + "*.png")
    file_mohan = glob.glob(path_white + "mohan.png")
    # スタート、フィニッシュの設定
    start = 0
    finish = num_w * num_h
    files = path_deep_white_file_list[start : finish]
    # 保存するフォルダの作成
    if not os.path.exists("merge"):
        os.makedirs("merge")
    #保存するフォルダ名（下の出力を見てコピペ）
    path_merge = "./merge/"
    # 画像と画像の間隔（ゼロは隙間なし）
    margin_w = 1
    margin_h = 1
    # 画像を貼り付けるキャンバスを作成
    if file_mohan != []:
        img_1 = Image.open(file_mohan[0])
        img_1 = img_1.resize(shape)
        w, h = img_1.size
        canvas = Image.new("RGBA", ((w + margin_w) * num_w, (h + margin_h) * num_h))
        for i in range(10):
            j = i+1
            print(f"{j}回目：{start}枚目から")
            print(f"{finish}枚目まで")
            file_list = glob.glob(path_white + "*.png")
            files = file_list[start : finish]
            canvas = Image.new("RGBA", ((w + margin_w) * num_w, (h + margin_h) * num_h))
            canvas = merge(files, shape, num_w, canvas, w, margin_w, h ,margin_h)
            canvas.save(path_merge + f'/{folda}_new{i}.png')
            start = finish
            if finish <= len(file_list):
                finish += num_w * num_h
            else:
                break
    else:
        canvas=0
    return canvas       


def Saiten_jidou(image_width=28, image_height=28, color_setting=1, W=6, H=6):
    # out_dirまでのパス
    PATH_out = "./setting/output/"
    # マルバツモデルを使う準備
    ci_cro = ["〇", "×"]
    model_ci = model_from_json(open('saiten/cnn_model.json', 'r').read())
    folda_ci = ["maru","batu"]
    # ひらがモデルを使う準備
    hiragana = ["あ","い","う","え","お","か","き","く","け","こ","さ","し","す","その他"]
    model_hira = load_model('saiten/hiragana.h5')
    folda_hira =["a","i","u","e","o","ka","ki","ku","ke","ko","sa","si","su","other"]
    #上のまとめ
    a=[0, model_ci , model_hira]
    b=["0", folda_ci, folda_hira]
    c=[0, ci_cro , hiragana]
    # 画像サイズを設定（ディープラーニングに合わせる）
#     image_width, image_height, color_setting = 28, 28, 1
    # 採点するフォルダ(bmpファイル画像)を指定
    folda_list = os.listdir("./deep")[1:]
    for i, folda in enumerate(folda_list):
        PATH_deep = "./deep/" + folda + "/off_image_test"
        bmp_list=glob.glob(PATH_deep +"/*.bmp")
        # 0点フォルダと正解フォルダを作成
        move_folda_0 = PATH_out + folda + "/0"
        if not os.path.exists(move_folda_0):
            os.mkdir(move_folda_0)
        if bmp_list==[]:
            print(f"{folda}フォルダは中身に画像がありません")
            continue
        plt.imshow(cv2.imread("./deep/" + folda + "/mohan.png"))
        plt.show()
        print(f"{folda}の分類開始します。")
        n = str(input("○×なら「1」、ひらがな判定なら「2」、mergeするなら「0」、スキップは「Enter」、途中で止めるなら「99」を入力"))
    # とばす、終了の処理
        if n == "0":
            path_out = "./deep/"
            path_file = path_out + folda +"/*.png"
            file_list = glob.glob(path_file)
            if file_list == []:
                continue
            img = cv2.imread("./deep/" + folda + "/white_image_test/mohan.png")
            ### パラメータ設定 ###
            k=str(input("縦長にするなら「0」横長にするなら「-」"))
            if k=="-":
                num_w = W 
                num_h = H
                shape = (100,50)
            elif k=="0":
                num_w = H
                num_h = W
                shape = (50, 100)
            else:
                print("わかりません")
                break
            canvas = merge_summary(folda, num_w, num_h, shape)
            if canvas==0:
                print(f"{folda}の移動は何もせず、スキップします。")
                continue  
            continue
        elif n == "99":
            print("中断しました。")
            break
        elif n == "1":
            print("正解を入力してください")
            print("【〇：0】【×：1】")
        elif n == "2":
            print("正解を入力してください")
            print("【あ：0】【い：1】【う：2】【え：3】【お：4】【か：5】【き：6】【く：7】【け：8】【こ：9】【さ：10】【し：11】【す：12】")
        else:
            print("スキップします。")
            continue
        l = int(input(""))
        m = str(input("配点を入力してください"))
    # 移動の命令
        move_folda_m = PATH_out + folda + "/" + m
        if not os.path.exists(move_folda_m):
            os.mkdir(move_folda_m)
    # 使うモデルを選択
        model=a[int(n)]
        model.load_weights('saiten/cicle_cross.h5') if n == "1" else model
        for j, bmp in enumerate(bmp_list):
            img = cv2.imread(bmp, 0)
            name = bmp.split("\\")[1]
            move_name = name.split("_")[0] + ".png"
            img = cv2.resize(img, (image_width, image_height))
            img = img.reshape(image_width, image_height, color_setting).astype('float32')/255 
            prediction = model.predict(np.array([img]))
            result = prediction[0]
            num = result.argmax()
            if not os.path.exists(PATH_out + folda+"/" + move_name):
                print("画像がありません")
                continue
            if l == num:
                shutil.move( PATH_out + folda+"/" + move_name, move_folda_m)
                print(f'{name}は、「', os.path.basename(c[int(n)][result.argmax()]),'」で正解。')
            else:
                shutil.move( PATH_out + folda+"/" + move_name, move_folda_0)
                print(f'{name}は、「', os.path.basename(c[int(n)][result.argmax()]),'」で不正解。')
                

def text_edit(text):  
    text  = text.replace("\hline \end{array}", "")
    text  = text.replace("& \end{array}", "")
    text  = text.replace("\end{array}", "")
    text = text.replace("\hline", "")
    text  = text.replace("\begin{array}", "")
    text  = text.replace("l|", "")
    text  = text.replace("c|", "")
    text  = text.replace("{}", "")    
#     text  = text.replace("", None)
    KO = ["\c","\d","\e","\g","\h","\i","\j","\k","\l","\m","\n","\o","\p","\q","\s","\w","\y","\z"]
    for k in KO:
        text  = text.replace(k, "")
    text = text.replace("\\", "& ")
    text = text.split("&")
    text_list=[]
    for i in text:
        t=i.strip()
        text_list.append(t)
#     text_list = list(filter(None, text_list))
    return text_list