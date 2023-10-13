#工数CSVファイルを一部列抽出して、Excelファイルへ出力する（当月以外のデータを除外）
#2023/04/24

import pandas as pd 
import sys
import datetime

#2023/06/23 autofilter
import xlwings as xw


def my_autofilter(filename, col):
  xw.App(visible=False)	#   エクセルの非表示を設定
  wb = xw.Book(filename)	#   既存のBOOKを開く
  if col == 5:
    rng = wb.sheets[0].range('a1:e10000')
  if col == 4:
    rng = wb.sheets[0].range('a1:d10000')

  #   No.が1
  rng.api.AutoFilter(Field=1, Criteria1="1")	# Field=1 は左から数えて1番目の列
  wb.sheets[0].api.ShowAllData()
  wb.save()    # 保存
  wb.close()    #   プロセスを削除
#End of my_autofiler

#入力：ファイル名.csv 出力: 1)ファイル名.xlsx  2)ファイル名_pivot.xlsx
args = sys.argv
filename=args[1]
filename1=filename[:-4]+'.xlsx'

print ('filename=' + filename)
print ('filename1=' + filename1)

#日付のYYYYMMを取得
dt_now = datetime.datetime.now()
#cur_year_month=int(dt_now.strftime('%Y%m')+'01')
cur_year_month=int('09'+'01')



#①社員名②勤務年月日③工数④製造オーダコード⑤製造オーダ名称 をCSVから取得

df = pd.read_csv(filename, 
     encoding='shift_jis',usecols=[1, 3, 18, 21, 22 ])
#print(df)

#②勤務年月日から先頭4桁YYYYMMを一致するものをfilter1と定例する
filter1=df['勤務年月日'] >= cur_year_month 

#filter1と一致するデータをExcelファイルへ出す
df[filter1].to_excel(filename1, sheet_name='工数実績',index=False)

print('Start: Call my_autofiler '+filename1)
my_autofilter(filename1,5)
print('End: Call my_autofiler')

filename2=filename1[:-5]+'_pivot.xlsx'
print('filename2='+filename2)


df=pd.read_excel(filename1, sheet_name='工数実績')
df_1=pd.pivot_table(df, index=['製造オーダコード','製造オーダ名称','社員名'], values='工数',aggfunc='sum')

df_1.to_excel(filename2, sheet_name='pivot')
print('Start: Call my_autofiler ' + filename2)
my_autofilter(filename2,5)
print('End: Call my_autofiler')
#End of Main

