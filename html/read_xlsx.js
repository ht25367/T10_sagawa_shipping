var select_flg = False;

alert("js が呼ばれました");

file_select.addEventListener("click", file_select_js, False);
function file_select_js() {
	alert("select_flt:" + select_flg );
	eel.file_select_py(select_flg);
}

eel.expose(file_select_flg);
function file_select_flg(flg) {
	select_flg = flg;
}
