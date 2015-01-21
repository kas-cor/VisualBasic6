<?php
session_start();
$email = !empty($_SESSION['cook_email']) ? $_SESSION['cook_email'] : "";
echo "<div class=\"box-modal\" id=\"callback\">\n";
echo "<div class=\"box-modal_close arcticmodal-close\">закрыть</div>\n";
echo "<div class=\"msg\"></div>\n";
echo "<div align=\"center\">\n";
echo "<table border=\"0\" cellpadding=\"2\" style=\"border-collapse: collapse\" width=\"90%\" cellspacing=\"2\">\n";
echo "<tr>\n";
echo "<td align=\"center\" width=\"100%\" colspan=\"2\">Для заказа обратного звонка заполните анкету<br />и мы перезвоним Вам</td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">&nbsp;</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">&nbsp;</td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">Организация&nbsp;<font color=\"#F00\">*</font>:</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">&nbsp;<input type=\"text\" name=\"org\" size=\"40\" value=\"\"></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">Контактное лицо&nbsp;<font color=\"#F00\">*</font>:</td>\n";
echo "<td align = \"left\" width=\"70%\" valign=\"top\">&nbsp;<input type=\"text\" name=\"name\" size=\"40\" value=\"\"></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">Телефон с кодом города&nbsp;<font color=\"#F00\">*</font>:</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">&nbsp;<input type=\"text\" name=\"telefon\" size=\"40\" value=\"\"></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">e-mail:</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">&nbsp;<input type=\"text\" name=\"email\" size=\"40\" value=\"$email\"></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">Регион&nbsp;<font color=\"#F00\">*</font>:</td>\n";
echo "<td align = \"left\" width=\"70%\" valign=\"top\">&nbsp;<select size=\"1\" name=\"region\">\n";
echo "<option value=\"\">[Выберите]</option>\n";

foreach ($region as $i => $value) {
	$sel = $reg == $i ? " selected" : "";
	echo "<option value=\"$i\"$sel>$value</option>\n";
}

echo "</select></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">Тема обращения:&nbsp;<font color=\"#F00\">*</font></td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">&nbsp;<textarea rows=\"5\" name=\"comment\" cols=\"30\"></textarea></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">Введите текст&nbsp;<font color=\"#F00\">*</font>:</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">\n";
echo "<img id=\"captcha\" src=\"/kcaptcha/\" width=\"120\" height=\"60\" align=\"left\" alt=\"\" />&nbsp;&nbsp;<input type=\"text\" name=\"img_check\" size=\"10\" value=\"\"><br />&nbsp;&nbsp;<a href=\"javascript:void(0);\" onclick=\"$('#captcha').attr('src','/kcaptcha/?rand='+Math.random());\"><span>Не видно !</span></a></td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">&nbsp;</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\"><font color=\"#F00\">*</font>&nbsp;Поля обязательные для заполнения</td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">&nbsp;</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\">&nbsp;</td>\n";
echo "</tr>\n";
echo "<tr>\n";
echo "<td align=\"right\" width=\"30%\" valign=\"top\">&nbsp;</td>\n";
echo "<td align=\"left\" width=\"70%\" valign=\"top\"><input type=\"button\" class=\"button white\" value=\"Отправить\" name=\"send\" onclick=\"callback.send();\"></td>\n";
echo "</tr>\n";
echo "</table>\n";
echo "</div>\n";
echo "</div>\n";
?>