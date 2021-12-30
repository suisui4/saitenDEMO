import kiritori as ki

# メイン処理 - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
if __name__ == "__main__":
    file_name="./setting/saiten.xlsx"
    # 採点の振り分けを
    ginou = [1,2,3,4,6,7,8,11,13,14,15,16,17,18,19,20,21,22,23,24,25,26]
    shikou = [5,9,10,12,27,28,29,30,31,32,33,34]
    ki.summary_table(file_name, ginou, shikou)