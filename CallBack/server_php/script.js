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
		msg += !org ? "�� ������� �������� �����������\n" : "";
		msg += !name ? "�� ������� ��� ����������� ����\n" : "";
		msg += !telefon ? "�� ������ ���������� �������\n" : "";
		msg += !region ? "�� ������ ������\n" : "";
		msg += !comment ? "�� ������� ���� ���������\n" : "";
		msg += !img_check ? "�� ������ ����������� �����\n" : "";
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
			alert("������!\n\n" + msg);
			return;
		} else {
			$.post("/callback/callback.php", "act=send&org=" + org + "&name=" + name + "&telefon=" + telefon + "&email=" + email + "&region=" + region + "&comment=" + comment + "&img_check=" + img_check, function(res) {
				if (res === "ok") {
					$("#callback").css({
						"text-align": "center",
						"font-weight": "bold"
					}).html("������ �� �������� ������ ���������,<br />�� ���������� ��� � ������� 10 ���.<br /><br />��������� �������� � ������� �����,<br>�� ������ � 9:00 �� 18:00 �� ����������� �������.");
				} else {
					$("#callback .msg").css({
						"color": "#f00",
						"text-align": "center",
						"height": "40px",
						"font-weight": "bold"
					});
					if (res === "err1") {
						$("#callback .msg").html("�� ���������� ���� �� ������������ �����");
					} else {
						$("#callback .msg").html("�� ����� ������ ����������� �����");
					}
				}
			});
		}
	}
};