<?php
	$rc4_pwd="{password}";
	function rc4($data) {
		global $rc4_pwd;
		$pwd=$rc4_pwd;
		$pwd_length=strlen($pwd);
		$data_length=strlen($data);
		for($i=0;$i<256;$i++) {
			$key[$i]=ord($pwd[$i%$pwd_length]);
			$box[$i]=$i;
		}
		$j=0;
		for($i=0;$i<256;$i++) {
			$j=($j+$box[$i]+$key[$i])%256;
			$tmp=$box[$i];
			$box[$i]=$box[$j];
			$box[$j]=$tmp;
		}
		$a=0;
		$j=0;
		for($i=0;$i<$data_length;$i++) {
			$a=($a+1)%256;
			$j=($j+$box[$a])%256;
			$tmp=$box[$a];
			$box[$a]=$box[$j];
			$box[$j]=$tmp;
			$k=$box[(($box[$a]+$box[$j])%256)];
			$cipher.=chr(ord($data[$i])^$k);
		}
		return $cipher;
	}
	function rc4_decode($txt) {
		for ($i=0;$i<strlen($txt);$i=$i+2) {
		$rs1=$txt[$i];
			$rs2=$txt[$i+1];
			if (ord($rs1)-48>-1 && ord($rs1)-48<10) {$r1=ord($rs1)-48;}
			if (ord($rs1)-55>9 && ord($rs1)-55<16) {$r1=ord($rs1)-55;}
			if (ord($rs2)-48>-1 && ord($rs2)-48<10) {$r2=ord($rs2)-48;}
			if (ord($rs2)-55>9 && ord($rs2)-55<16) {$r2=ord($rs2)-55;}
			$s.=chr($r1*16+$r2);
		}
		return $s;
	}
	function rc4_encode($txt) {
		for($i=0;$i<strlen($txt);$i++) {
			$c=strtoupper(dechex(ord($txt[$i])));
			if (strlen($c)==1) {$c="0".$c;}
			$s.=$c;
		}
		return $s;
	}
?>