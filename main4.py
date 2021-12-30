import my_gragh as gra
import seaborn as sns
import pandas as pd

# メイン処理 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
if __name__ == "__main__":
  #グラフ日本語、メモリ数字の大きさを指定
    sns.set(font='Yu Gothic',font_scale = 1.5)
    file_name="./setting/saiten.xlsx"
    df_summary = pd.read_csv("./seitoPDF_CSV/kousa_summary.csv",encoding="shift_jis")
    # 下のほうに余分なのがあるため調整
    df_summary=df_summary.iloc[0:-5,:]
    df_shomon = gra.shomon_table(file_name)
    df_wariai = gra.wariai_table(df_shomon)
    #小数表示が嫌なら
    # df_wariai = pd.DataFrame(data=df_wariai.values.round(2),columns=df_wariai.columns, index=df_wariai.index)
    gra.summary_hist(df_summary)
    gra.AVE_MED(df_summary, MF="off")
    gra.total_kanten(df_summary)
    gra.shomon_wariai(df_wariai, num=7)
    gra.daimon_detail(df_shomon)