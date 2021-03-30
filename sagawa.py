from selenium.webdriver import Chrome, ChromeOptions
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
import openpyxl


# Chromeを起動する関数
def set_driver(driver_path, headless_flg):
	# Chromeドライバーの読み込み
	options = ChromeOptions()
	
	# ヘッドレスモード（画面非表示モード）の設定
	if headless_flg == True:
		options.add_argument('--headless')

	# 起動オプションの設定
	options.add_argument(
		'--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36')
	# options.add_argument('log-level=3')
	options.add_argument('--ignore-certificate-errors')
	options.add_argument('--ignore-ssl-errors')
	options.add_argument('--incognito')		  # シークレットモードの設定を付与
	
	# os.chdir( os.path.dirname(os.path.abspath(__file__)) )
	# return Chrome(executable_path=os.getcwd() + "/" + driver_path, options=options)
	return Chrome(ChromeDriverManager().install(), options=options)


def main():
	toi_input_url = "https://k2k.sagawa-exp.co.jp/p/sagawa/web/okurijoinput.jsp"

	# １、Excel読込。引数 index=0 で読むと1列目が indexになりキー取得不可に
	delivery_pd = pd.read_excel("佐川フォーマット.xlsx",sheet_name=0)
	# print(f"pd_tpye:{ type(delivery_pd)}, ")
	# print(f"{delivery_pd.iat[2,1]}")

	# ２、佐川急便HPを開く
	driver = set_driver("chromedriver.exe", False)
	driver.get(toi_input_url)
	time.sleep(3)

	# ３、送り状番号 elm取得＞ No入力
	i_cnt = 0
	delivery_elm = driver.find_elements_by_css_selector(".toiban-dt1")
	for i,row in delivery_pd.iterrows():
		toiStart_elm = driver.find_element_by_id("main:toiStart")
		i_cnt += 1
		input_i = i%10
		time.sleep(0.3)
		delivery_elm[input_i].send_keys(row["問い合わせ番号"])
		
		# １０件ごとに「問い合わせ」結果ページへ
		if i>1 and (i+1)%10==0 :
			toiStart_elm.click()
			time.sleep(2)

			# 問い合わせ結果取得 & pd入力
			toi_state_elm = driver.find_elements_by_class_name("state")
			for j,elm in enumerate(toi_state_elm):
				delivery_pd.iat[(i-9)+j,2] = elm.text
			
			driver.back()
			time.sleep(2)

			# 入力番号のクリア
			delivery_elm = driver.find_elements_by_css_selector(".toiban-dt1")
			for j in range(10):
				delivery_elm[j].clear()

	
	print(f"i_cnt:{i_cnt},  input_i:{input_i}")
	if i_cnt %10 != 0 :
		toiStart_elm.click()
		time.sleep(2)

		# 問い合わせ結果取得 & pd入力
		input_i += 1
		i_cnt -= (input_i)
		toi_state_elm = driver.find_elements_by_class_name("state")
		for j in range(input_i):
			delivery_pd.iat[i_cnt+j,2] = toi_state_elm[j].text


	# ５、取得した結果をExcelに記述
	delivery_pd.to_excel( "佐川フォーマット.xlsx" )


	# ６、eelで GUI(「ファイル選択」「実行ボタン」)を実装



if __name__ == "__main__":
	main()