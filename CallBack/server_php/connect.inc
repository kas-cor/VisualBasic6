<?php
	@mysql_connect("host","user","password") or die("Не могу создать соединение");
	@mysql_select_db("base") or die("Не могу выбрать базу");  
?>
