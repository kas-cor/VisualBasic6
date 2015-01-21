var callback = {
	init: function() {
		$.post("/callback/callback.php", "act=init", function(res) {
			document.write(res);
		});
	},
	getinp: function(inp, t) {
		switch (t) {
			case 0:
				return $("#callback input[name='" + inp + "']").val();
			case 1:
				return $("#callback select[name='" + inp + "'] option:selected").val();
			case 2:
				return $("#callback textarea[name='" + inp + "']").val();
		}
	},
	checkform: function(org, name, telefon, region, comment, img_check) {
		var msg = "";
		msg += !org ? "Не введено название организации\n" : "";
		msg += !name ? "Не введено имя контактного лица\n" : "";
		msg += !telefon ? "Не введен контактный телефон\n" : "";
		msg += !region ? "Не выбран регион\n" : "";
		msg += !comment ? "Не введена тема обращения\n" : "";
		msg += !img_check ? "Не введен контрольный текст\n" : "";
		return msg;
	},
	send: function() {
		var org = this.getinp("org", 0);
		var name = this.getinp("name", 0);
		var telefon = this.getinp("telefon", 0);
		var email = this.getinp("email", 0);
		var region = this.getinp("region", 1);
		var comment = this.getinp("comment", 2);
		var img_check = this.getinp("img_check", 0);
		var msg = this.checkform(org, name, telefon, region, comment, img_check);
		if (msg) {
			alert("Ошибки!\n\n" + msg);
			return;
		} else {
			$.post("/callback/callback.php", "act=send&org=" + org + "&name=" + name + "&telefon=" + telefon + "&email=" + email + "&region=" + region + "&comment=" + comment + "&img_check=" + img_check, function(res) {
				if (res === "ok") {
					$("#callback").css({
						"text-align": "center",
						"font-weight": "bold"
					}).html("Запрос на обратный звонок отправлен,<br />мы перезвоним Вам в течении 10 мин.<br /><br />Менеджеры отвечают в рабочее время,<br>по будням с 9:00 до 18:00 по Московскому времени.");
				} else {
					$("#callback .msg").css({
						"color": "#f00",
						"text-align": "center",
						"height": "40px",
						"font-weight": "bold"
					});
					if (res === "err1") {
						$("#callback .msg").html("Не заполненно одно из обязательных полей");
					} else {
						$("#callback .msg").html("Не верно введен контрольный текст");
					}
				}
			});
		}
	}
};