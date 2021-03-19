from selenium.webdriver import Chrome, ChromeOptions
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import openpyxl


def main():
	# Elcel読込。引数 index=0 で読むと1列目が indexになり取得不可に
	delivery_pd = pd.read_excel("佐川フォーマット.xlsx",sheet_name=0)
	
	# for i,row in delivery_pd.iterrows():
	# 	print(f"{row['パターン']}, {row['問い合わせ番号']}, {row['結果']}")

	# ２、佐川急便HPを開く

	# ３、配送状況を取得するためのエレメントを取得

	# ４、伝票番号を入力

	# ５、実行結果を取得

	# ６、取得した結果をExcelに記述

	# ７、eelで GUI(「ファイル選択」「実行ボタン」)を実装



if __name__ == "__main__":
	main()