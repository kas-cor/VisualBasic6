<?php

session_start();
include("connect.inc");
include("func.inc");

if ($act=="init") {
	ob_start();
	include("form.inc");
	$ret = ob_get_contents();
	ob_end_clean();
	if (!preg_match("//u", $ret)) {$ret=iconv("CP1251","UTF-8",$ret);}
	echo $ret;
	mysql_close();
	exit();
}

if ($act=="send") { // Создание запроса на звонок в базе
	$org = trim(htmlspecialchars(stripslashes($org)));
	$name = trim(htmlspecialchars(stripslashes($name)));
	$telefon = trim(htmlspecialchars(stripslashes($telefon)));
	$email = trim(htmlspecialchars(stripslashes($email)));
	$region = trim(htmlspecialchars(stripslashes($region)));
	$comment = trim(htmlspecialchars(stripslashes($comment)));
	if (preg_match("//u", $org)) {$org=iconv("UTF-8","CP1251",$org);}
	if (preg_match("//u", $name)) {$name=iconv("UTF-8","CP1251",$name);}
	if (preg_match("//u", $telefon)) {$telefon=iconv("UTF-8","CP1251",$telefon);}
	if (preg_match("//u", $email)) {$email=iconv("UTF-8","CP1251",$email);}
	if (preg_match("//u", $region)) {$region=iconv("UTF-8","CP1251",$region);}
	if (preg_match("//u", $comment)) {$comment=iconv("UTF-8","CP1251",$comment);}
	if (isset($_SESSION['captcha_keystring']) && $_SESSION['captcha_keystring'] == $_POST['img_check']) {
		if ($org . $name . $telefon . $region . $comment != "") {
			mysql_query("INSERT INTO callback(org,name,telefon,email,region,comment,time) VALUES('$org','$name','$telefon','$email','$region','$comment',UNIX_TIMESTAMP())");
			echo "ok";
		} else {
			echo "err1";
		}
	} else {
		echo "err2";
	}
	mysql_close();
	exit();
}

/*
 * Не изменяемые данные !!!
 */
if (!empty($calls)) { // Вывод списка обратных звонков
	$result = mysql_query("SELECT * FROM callback WHERE manager='' ORDER BY id ASC");
	if (mysql_numrows($result) != 0) {
		while ($calls_arr = mysql_fetch_array($result)) {
			$comment = wordwrap($calls_arr[comment], 80);
			$txt.="$calls_arr[id]#;#$calls_arr[org]#;#$calls_arr[name]#;#$calls_arr[telefon]#;#$calls_arr[email]#;#" . $region[$calls_arr[region]] . "#;#$comment#;#" . date("d.m.Y (H:i:s)", $calls_arr[time]) . "{LineBreak}";
		}
		echo rc4_encode($txt);
	} else {
		echo rc4_encode("Empty");
	}
	mysql_close();
	exit();
}
if (!empty($mark_call)) { // Пометка звонка менеджером
	if (rc4_decode($key) == $rc4_pwd) {
		mysql_query("UPDATE callback SET manager='$manager' WHERE id='$mark_call'");
		echo rc4_encode("Ok");
	} else {
		echo rc4_encode("Error");
	}
	mysql_close();
	exit();
}
?>