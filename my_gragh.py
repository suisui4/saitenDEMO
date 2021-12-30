from matplotlib import pyplot as plt
import seaborn as sns
import pandas as pd

#グラフ日本語、メモリ数字の大きさを指定
sns.set(font='Yu Gothic',font_scale = 1.5)

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


def summary_hist(df_summary):
    fig = plt.figure(figsize = (20,10))
    # 全体のヒストグラム
    ax1 = fig.add_subplot(2,1,1)
    ax1 = sns.histplot(df_summary, x="合計", hue="性別", binwidth=10, multiple='stack')
    # ax1 = df_summary["合計"].plot.hist(bins=10)
    ax1.xaxis.set_visible(False)
    ax1 = ax1.set_ylabel("人数", fontsize = 20)
    ax1 = plt.xticks(list(map(lambda r:r*10, list(range(11)))))
    ax1 = plt.yticks(list(map(lambda r:r*5, list(range(6)))))
    # 全体の箱ひげ図
    ax2 = fig.add_subplot(2,1,2)
    ax2 = sns.boxplot(x= "合計",data=df_summary, orient="h")
    ax2 = ax2.set_xlabel("点数", fontsize = 20)
    ax2 = plt.xticks(list(map(lambda r:r*10, list(range(11)))))
    plt.savefig("./JPG_LISTS/SUMMARY_HIST_BOX.jpg", bbox_inches = "tight")

# 平均値と中央値用のデータフレームを作成
def AVE_MED(df_summary, MF="off"):
    zen_lis=[df_summary["合計"].describe()["mean"].round(2),df_summary["合計"].describe()["50%"]]
    ave=df_summary["合計"].groupby(df_summary["組"]).mean().round(2)
    med=df_summary["合計"].groupby(df_summary["組"]).median()
    df_zen=pd.DataFrame(data=[ave, med], index=["平均点","中央値"])
    df_zen["全体"]= zen_lis
    class_=df_summary["組"].unique()
    class_lists = [str(i+1)+"組" for i in range(len(class_))]
    data = class_lists
    data.appned("全体")
    df_zen.columns=data
    dfzen = df_zen.T.reset_index()
    fig, ax = plt.subplots(figsize=(10,2.5))
    ax.set_axis_off()
    table = ax.table(
        cellText = df_zen.values,
        rowLabels = df_zen.index,
        colLabels = df_zen.columns,
        rowColours =["palegreen"] * 2,
        colColours =["palegreen"] * 5,
        cellLoc ='center',
        loc ='upper left')
    ax.set_title('～結果発表～',
                fontweight ="bold")
    # 表をfigure全体に表示させる
    for pos, cell in table.get_celld().items():
        cell.set_height(1/(len(df_zen.values)*2))
    plt.savefig("./JPG_LISTS/AVE_MED_DATA.jpg", bbox_inches = "tight")
    plt.close()
    #グラフ化
    fig, ax = plt.subplots(1,3, figsize = (20, 8), gridspec_kw={'width_ratios': [4,4,1]}, subplot_kw=({"xticks":(), "yticks":()}))
    #横の余白をなし
    plt.subplots_adjust(wspace=0)
    # クラス毎のヒストグラム
    ax1 = fig.add_subplot(1,3,1)
    ax1 = sns.barplot(x="index", y="平均点", data=dfzen)
    ax1.set_xlabel("組ごとの平均", fontsize = 20)
    ax1.set_ylabel("点数", fontsize = 20)
    ax1 = plt.yticks(list(map(lambda r: 30+r*10, list(range(6)))))
    plt.tick_params(labelsize=25)

    # クラス毎の箱ひげ図
    ax2 = fig.add_subplot(1,3,2)
    if MF=="on":
        ax2 = sns.boxplot(x="組",y= "合計",data=df_summary, hue="性別")
    else:
        ax2 = sns.boxplot(x="組",y= "合計",data=df_summary)
    ax2.set_xlabel("組ごとの箱ひげ図", fontsize = 20)
    ax2.set_ylabel("")

    # グリッドを付ける
    ax2.xaxis.set_ticklabels(class_lists)
    ax2.yaxis.set_ticklabels([])
    ax2 = plt.yticks(list(map(lambda r: 30+r*10, list(range(6)))))
    plt.tick_params(labelsize=25)

    # 学年の箱ひげ図
    ax3 = fig.add_subplot(1,3,3)
    if MF=="on":
        ax3 = sns.boxplot(x="年",y= "合計",data=df_summary, hue="性別")
    else:
        ax3 = sns.boxplot(x="年",y= "合計",data=df_summary)
    ax3.xaxis.set_ticklabels(["学年"])
    ax3.yaxis.set_ticklabels([])
    ax3.set_xlabel("")
    ax3.set_ylabel("")
    ax3 = plt.yticks(list(map(lambda r: 30+r*10, list(range(6)))))
    plt.tick_params(labelsize=25)
    plt.suptitle("考査結果")
    plt.savefig("./JPG_LISTS/CLUSS_AVE_MED_DATA.jpg", bbox_inches = "tight")

    return df_zen

def total_kanten(df_summary):
    fig = plt.figure(figsize=(20,8))
    ax1 = fig.add_subplot(1,2,1)
    # 散布図
    ax1 = sns.scatterplot(x=df_summary['技能'], y=df_summary['合計'])
    # 回帰直線
    ax1 = sns.regplot(x=df_summary['技能'], y=df_summary['合計'])
    ax2 = fig.add_subplot(1,2,2)
    ax2 = sns.scatterplot(x=df_summary['思考'], y=df_summary['合計'])
    ax2 = sns.regplot(x=df_summary['思考'], y=df_summary['合計'])
    plt.savefig("./JPG_LISTS/TOTAL_GINOU_SHIKOU.jpg", bbox_inches = "tight")

def wariai_table(df_shomon):
    wariai = pd.DataFrame(columns=["wrong","sankaku","correct"])
    df_saiten4=df_shomon.replace(['未'], 0)
    file_list = df_saiten4.columns[1:]
    for file in file_list:
        # ユニークが２つの場合
        if df_saiten4[file].nunique()==2:
    #         ％テージに変換
            df_2 = df_saiten4[file].value_counts(normalize=True)*100
            df_2 = pd.DataFrame(df_2.sort_index()).T
            df_2.columns = ["wrong", "correct"]
            wariai = pd.concat([wariai,df_2],axis=0)
        # ユニークが３つの場合
        elif df_saiten4[file].nunique()==3:
            df_3 = df_saiten4[file].value_counts(normalize=True)*100
            df_3 = pd.DataFrame(df_3.sort_index()).T
            df_3.columns = ["wrong", "sankaku", "correct"]
            wariai = pd.concat([wariai,df_3],axis=0)
        # ユニークが４つ以上の場合
        elif df_saiten4[file].nunique() > 3:
            df_4=df_saiten4[file].value_counts()
            df_4 = pd.DataFrame(df_4.sort_index()).T
    #         ０点と満点以外をすべて足す。真ん中は削除する。
            z = df_4.columns[1:-1]
            df_4[99]=df_4.iloc[:,1:-1].sum(axis=1)
            df_4 = df_4.drop(z,axis=1).T.sort_index().T
    #         並び替え
            df_4=df_4.iloc[:, [0,2,1]]
            df_4.columns = ["wrong", "sankaku","correct"]
            for i, ans in enumerate(df_4.columns):
                df_4[ans]=(df_4.iloc[0,i]/df_4.sum(axis=1))*100
            wariai = pd.concat([wariai,df_4],axis=0)
        # ユニークが１つの場合
        else:
            df_5 = pd.DataFrame()
            df_5["wrong"]=[100]
            df_5["sankaku"]=[0]
            df_5["correct"]=[0]
            df_5 = df_5.rename({0:file})
            wariai = pd.concat([wariai,df_5],axis=0)
    wariai = wariai.fillna(0)
    return wariai

def shomon_wariai(df_wariai, num=5):
    df_wrong = df_wariai.sort_values('wrong')
    df_correct = df_wariai.sort_values('correct')
    df_sankaku = df_wariai.sort_values('sankaku')

    fig, ax = plt.subplots(3,1, figsize = (20, 15), subplot_kw=({"xticks":(), "yticks":()}))
    plt.subplots_adjust(hspace=0.3)
    ax1 = fig.add_subplot(3,1,1)
    ax1 = df_wrong["wrong"].tail(num).plot.barh()
    ax1 = plt.title(f"WRONG_BEST_{num}")

    ax2 = fig.add_subplot(3,1,2)
    ax2 = df_correct["correct"].tail(num).plot.barh()
    ax2 = plt.title(f"CORRECT_BEST_{num}")

    ax3 = fig.add_subplot(3,1,3)
    ax3 = df_sankaku["sankaku"].tail(num).plot.barh()
    ax3 = plt.title(f"SANKAKU_BEST_{num}")
    plt.savefig("./JPG_LISTS/SHOUMON_WARIAI.jpg", bbox_inches = "tight")

def daimon_detail(df_shomon):
    df_ques_data = df_shomon.drop(["ファイル名"], axis=1).T
    df_ques_data = df_ques_data.reset_index()
    df_ques_data["index"] = df_ques_data["index"].str.split("NO").str.get(1).str.split("_").str.get(0)

    df_ques = pd.DataFrame()
    df_ques["大問番号"] = df_ques_data["index"].str.split("-").str.get(0)
    # df_ques["小問番号"] = df_ques_data["index"].str.split("-").str.get(1)
    df_ques_data = df_ques_data.drop(["index"], axis=1)
    df_ques_data_sumarry = pd.concat([df_ques,df_ques_data], axis=1)

    df_ques_data_sumarry = df_ques_data_sumarry.replace("未",0)
    df_ques_data_sumarry["大問番号"]=df_ques_data_sumarry["大問番号"].apply(lambda x: f"NO{str(x).zfill(2)}")
    df_ques_data_DAI = pd.DataFrame()
    for i in range(12):
        new = df_ques_data_sumarry[i].groupby(df_ques_data_sumarry["大問番号"]).apply(lambda x: x.sum())
        df_ques_data_DAI[i] = new
    sns.pairplot(df_ques_data_DAI.T, diag_kind="kde",plot_kws={"alpha":0.2})
    plt.savefig("./JPG_LISTS/DAIMON_DETAIL.jpg", bbox_inches = "tight")
    return df_ques_data_DAI.T