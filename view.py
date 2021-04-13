import eel
import desktop
import read_xlsx,sagawa

app_name="html"
end_point="index.html"
size=(900,660)

@ eel.expose
def file_select_py(select_flg):
	read_xlsx.file_select_py(select_flg)

@ eel.expose
def delivery_update(file_name):
	sagawa.delivery_update(file_name)

desktop.start(app_name,end_point,size)
#desktop.start(size=size,appName=app_name,endPoint=end_point)
