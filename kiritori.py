import cv2
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import os
import glob
from pathlib import Path
from pdf2image import convert_from_path
from PIL import Image
import shutil


# poppler/binを環境変数PATHに追加する
poppler_dir = "C:/poppler/bin"
os.environ["PATH"] += os.pathsep + str(poppler_dir)


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
    print("大問の切り取り状況を保存しました。(daimon.jpg)")
    plt.savefig("./daimon.jpg")
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
    n = str(input("大問画像と小問画像を保存しますか？保存するなら【yes】と入力"))
    img = cv2.imread(file)
    for i in range(len(left_big)):
        x, x_end, t, t_end = left_big[i]
    #このimg_tissueが切った大問画像
        img_tissue = img[t:t_end, x:x_end]
    # 画像保存処理。
        if n =="yes":
            out_dir_suf = f"CUT_NO{i+1}"
            out_dir = str(out_dir_suf)
            if not os.path.exists(out_dir):
                os.makedirs(out_dir)
            cv2.imwrite(out_dir + f"\\base_NO{i+1}.jpg",img_tissue)
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
            namelist.append(f"NO{str(i+1).zfill(2)}-{num+1}")
            x1=min(rects[num].T[0])
            x2=max(rects[num].T[0])
            y1=min(rects[num].T[1])
            y2=max(rects[num].T[1])
    # このimg_tissue2が切った小問画像
            img_tissue_2 = img_tissue[y1:y2, x1:x2]
    # 保存処理
            if n=="yes":
                out_dir_suf2 = f"\\small_question"
                out_dir2 = str(out_dir + out_dir_suf2)
                if not os.path.exists(out_dir2):
                    os.makedirs(out_dir2)
                cv2.imwrite(out_dir2 + f"\\base_no{num+1}.jpg",img_tissue_2)
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

def sortdata_2dan_yoko(df, file, namelist):
    img = cv2.imread(file)
    # 二段に分ける
    df.loc[df['end_x'] < img.shape[1]/2, 'label'] = 'left'
    df.loc[df['end_x'] > img.shape[1]/2, 'label'] = 'right'

    # 左半分
    df1=df[df["label"]=="left"]
    # yの値で高さを小さい順にした後、xの値で小さい順に
    df1_sort=df1.sort_values(["end_y","start_x"])
    # ただし、yの値でsortしても誤差が１～５あるので、その部分は許容するようにラベル付けをして調整する。
    df1_sort["立幅フラグ"]=0
    label = 1
    for i in range(df1_sort.shape[0]):
        gosa = df1_sort[["end_y"]].iloc[i]-df1_sort[["end_y"]].iloc[i-1]
        if gosa[0] < 5:
            df1_sort["立幅フラグ"].iloc[i] = label
        else:
            label+=1
            df1_sort["立幅フラグ"].iloc[i] = label
    df1_sort=df1_sort.sort_values(["立幅フラグ","start_x"])

    # 右半分もどうように行う。
    df2=df[df["label"]=="right"]
    df2_sort=df2.sort_values(["end_y","start_x"])

    # sortしても、誤差が１～５あるので調整
    df2_sort["立幅フラグ"]=0
    label = 1
    for i in range(df2_sort.shape[0]):
        gosa = df2_sort[["end_y"]].iloc[i]-df2_sort[["end_y"]].iloc[i-1]
        if gosa[0] < 5:
            df2_sort["立幅フラグ"].iloc[i] = label
        else:
            label+=1
            df2_sort["立幅フラグ"].iloc[i] = label
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
    plt.savefig("./jidoukiritori.jpg")
    print("画像のように切りとります。(jidoukiritori.jpg)よければ次に進んでください。")
    plt.imshow(img)
    return df3

def trimdata_save(df):
    file_name="./setting/trimData.csv"
    df_trim= pd.read_csv(file_name, index_col=0)
    df_newtrim = pd.concat((df_trim[:1], df), axis=0)
    return df_newtrim.to_csv("./setting/trimData.csv", encoding="utf-8", index=True)


def WhitePDFchengeJPG(PDF_Wname, dpi=150):
    # PDFファイルのパス
    pdf_path = Path(PDF_Wname)
    # PDF -> Image に変換 dpi
    pages = convert_from_path(str(pdf_path), dpi)

    image_white_path = "./base_white_mohan/white.jpg"
    image_path = "./setting/input/white.jpg"
    pages[0].save(str(image_white_path), "JPEG")
    pages[0].save(str(image_path), "JPEG")

    print("base_white_mohanフォルダに保存しました。")
    print("./setting/inputフォルダに保存しました。")

def PDFchangeJPG(PDF_name, dpi=150):
    # PDFファイルのパス
    pdf_path = Path(PDF_name)
    # PDF -> Image に変換 dpi
    pages = convert_from_path(str(pdf_path), dpi)
        # 画像ファイルを１ページずつ保存
    for i, page in enumerate(pages):
        file = f"change_{i}" + ".jpg"
        image_path = "./setting/input/" + file
        page.save(str(image_path), "JPEG")
    print("./setting/inputフォルダに保存しました。")

def seitoPDFchengeJPG(PDF_seito, CSV_meibo, dpi=150):
    df_seito_meibo=pd.read_csv(CSV_meibo)
    df_seito_meibo=df_seito_meibo.astype(str)
    df_seito_meibo["番号"]=df_seito_meibo["番号"].apply(lambda x: x.zfill(3))
    df_seito_meibo["連結"] = df_seito_meibo["年"]+"-"+df_seito_meibo["組"]+"-"+(df_seito_meibo["番号"])
    meibo=df_seito_meibo["連結"].to_list()
    # PDFファイルのパス
    pdf_path = Path(PDF_seito)
    # PDF -> Image に変換 dpi
    pages = convert_from_path(str(pdf_path), dpi)
    if len(meibo)==len(pages):
        print(f"名簿と画像の枚数({len(meibo)}毎)が一致しています。inputフォルダにすべて保存します。")
        # 画像ファイルを１ページずつ保存
        for i, page in enumerate(pages):
            file_seito = f"{meibo[i]}" + ".jpg"
            PATH_input = "./setting/input"
            image_path = PATH_input +"\\"+ file_seito
        # JPEGで保存
            page.save(str(image_path), "JPEG")
            if (i+1) % 5==0:
                print(f"{i+1}枚目保存しました。")
        print("保存が完了しました。")
    else:
        print("要確認:名簿とPDFの枚数が一致していません。")
        print(f"名簿の人数{len(meibo)}人")
        print(f"PDFの枚数{len(pages)}")

def point(img, base_shape):
    detector = cv2.ORB_create()
    img = cv2.resize(img, dsize = base_shape)
    # 特徴点と特徴量検出
    kp, des = detector.detectAndCompute(img, None)
    if kp ==():
        return 0
    else:
        return des.shape[0]

# 高さが大きすぎを防ぐ
def scale_to_height(img, height):
    h, w = img.shape[:2]
    width = round(w * (height / h))
    dst = cv2.resize(img, dsize=(width, height))
    return dst

def point(img, base_shape):
    detector = cv2.ORB_create()
    img = cv2.resize(img, dsize = base_shape)
    # 特徴点と特徴量検出
    kp, des = detector.detectAndCompute(img, None)
    if kp ==():
        return 0
    else:
        return des.shape[0]


def zero_move(basefile="base-0-000.jpg"):
    foldaname_list = os.listdir("./setting/output")
    for folda in foldaname_list[1:]:
        file_list = glob.glob("./setting/output/"+ folda+"/*.jpg")
        im_base = cv2.imread("./setting/output/" + folda +"/"+ basefile)
# 画像が小さいと特徴点を検出しないので、１０倍に拡大
        base_shape = (im_base.shape[0]*10, im_base.shape[1]*10)
        base_point = point(im_base, base_shape)
# 横長だったり、解答欄が大きすぎるとずれが出る為、大きさによって処理を変える
        if base_shape[0] > 800 or base_shape[0]*10 < base_shape[1]:
            base_shape = (400, 600)
            base_point = point(im_base, base_shape) + 20
# 座標平面や図形は線の判別がしずらいので移動は行わない。
        if base_point > 300:
            print(f"大問{folda}は座標、図形があるため移動しません。")
        else:
            outdir = "./setting/output/" + folda + "/0"
            if not os.path.exists(outdir):
                os.makedirs(outdir)
            print(f"大問{folda} : base_point:{base_point} : {base_shape}")
            for file in file_list:
                img = cv2.imread(file)
                img_point = point(img, base_shape)
                if img_point <= base_point:
                    shutil.move("./setting/output/" + folda + "/" + file.split("\\")[1], outdir)
                else:
                    name=file.split('\\')[1]
                    print(f"ポイント数{img_point}のため、{name}は移動しませんでした")


def shomon_table(file_name):
    df_saiten=pd.read_excel(file_name)
    df_shomon = df_saiten.iloc[:-1,:-2].drop(['画像',"生徒番号","名前"],axis=1)
    return df_shomon

def summary_table(file_name, ginou, shikou):
    df_shomon = shomon_table(file_name)
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