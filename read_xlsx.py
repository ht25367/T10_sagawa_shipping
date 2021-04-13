import os
import tkinter as tk
import tkinter.filedialog as fd
import eel
import pandas as pd
import openpyxl


### Tkinter にて ファイルダイアログで xlsxファイル選択
def file_select_py(select_flg):
	file_name =""

	# Tkinter の準備(謎窓の非表示、前面表示）
	root = tk.Tk()
	root.withdraw()
	root.attributes('-topmost', True)
	root.focus_force()

	# ファイルダイアログを1つ開く(複数開かない為の select_flg)
	if select_flg == False:
		eel.file_select_flg(True)

		f_type = [("エクセルファイル","*.xlsx"),("all file","*")]
		i_dir = os.path.abspath(os.path.dirname(__file__))
		file_name = fd.askopenfilename( filetypes=f_type, initialdir=i_dir,title='エクセルファイルを選択' )

		eel.file_select_flg(False)


	if file_name.find(".xlsx") > 0:
		# エクセルファイルが選ばれたら、<table>の初期化 > 書込み
		file_name = file_name[file_name.rfind("/") +1:]
		eel.clear_table(file_name)
		df_write_table(file_name)




def df_write_table(file_name):
	# xlsx ファイルを開く
	delivery_pd = pd.read_excel(file_name, header=0,sheet_name=0)
	
	# １行づつ読み ＞ 書き
	j = 0
	for i, row in enumerate(delivery_pd.itertuples()):
		# print(f"i:{i}, {row[1]}, {row[2]}, {row[3]}")
		eel.write_rowdata(i,row[1], str(row[2]),str(row.結果))
		j += 1

	# なぜ？最終行だけループしないので手動で
	if j > 9:
		tail_row = delivery_pd[-1:]
		# print(f"j:{j},{tail_row.values[0,0]}, {tail_row.values[0,1]},{tail_row.values[0,2]}" )
		eel.write_rowdata(j,tail_row.values[0,0], tail_row.values[0,1], str(tail_row.values[0,2]) )
		