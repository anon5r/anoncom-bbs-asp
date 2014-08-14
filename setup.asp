<% @Language = "VBScript" %>
<!-- #Include file="bbsdb.asp" -->
<%
'バッファを有効にする
Response.Buffer = True


If Session("login") <> 1 Then
	'ログインされていない状態
	Response.Redirect "admin.asp"
Else

If BBSBlank = True Then
	Response.Redirect "blankbbs.html"
End If


%>
<html lang="ja">
<head>
<title>anoncomBBS Installer</title>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/css; charset=shift_jis">
<style type="text/css">
<!--
a:link { text-decoration:none; color:#0000ff; }
a:visited { text-decoration:none; color:#0000ff; }
a:active { text-decoration:none; color:#ff0000; }
a:hover { text-decoration:underline; color:#ff0000; cursor:default; }

body{
	bakground-color:#ffffff;
	color:#000000;
	font-size:10pt;
	font-family:'MS UI Gothic','ＭＳゴシック';
	overflow-y:auto;
	scrollbar-base-color:#ffffff;
	scrollbar-face-color:#ffffff;
	scrollbar-arrow-color:#888888;
	scrollbar-highlight-color:#ffffff;
	scrollbar-3dlight-color:#888888;
	scrollbar-shadow-color:#888888;
	scrollbar-darkshadow-color:#ffffff;
}

table#bbs{
	border-style:solid;
	border-color:#888888;
	border-width:1px;
}

td{
	color:#000000;
	font-size:12pt;
	font-family:'MS UI Gothic','ＭＳゴシック';
}


input{
	color:#000000;
	background-color:#ffffff;
	font-size:12pt;
	font-family:'MS UI Gothic';
	border-style:solid;
	border-width:1px;
	border-color:#888888;
}

-->
</style>
</head>
<body bgcolor="#ffffff" text="#000000" link="#0000ff" alink="#ff0000" vlink="#0000ff">
<font face="Times New Roman" size="+3"><b><i>
anoncomBBS Installer
</i></b></font><br>
<br>
<hr size="1">
<br>
<br>
<br>
<%

'###########################################
'####					####
'####	     a n o n c o m B B S	####
'####					####
'###########################################
'				by anoncom.net

'ＢＢＳ初期化ファイル

If Request.Form = "" Then
%>
<center>
<table border="0" height="40%" width="70%">
<tbody>
<tr>
<td align="center" bgcolor="#ddffdd">
<table border="0">
<tbody>
<tr>
<td align="center" bgcolor="#ddffdd" colspan="2">
<b>
フォルダ内のanoncomBBSの設定を初期化します。よろしいですか？<br>
<u>掲示板の書き込み内容は初期化されません。</u><br>
※この作業は取り消すことが出来ません。
</td>
</tr><tr>
<td align="center" height="50">&nbsp;</td>
</tr><tr>
<td align="center" valign="bottom">
<form action="setup.asp" method="post">
<input type="hidden" name="setup" value="execute">
<input type="submit" value="はい">
</td>
<td align="center" valign="bottom">
<input type="button" value="いいえ" onClick="javascript:history.back(-1)">
</td>
</td>
</tr>
</tbody>
</table>
</tr>
</tbody>
</table>
</center>
<%
Else

'初期化実行


'データベース接続
Set db=Server.CreateObject("ADODB.Connection")

db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Mode = 3
db.ConnectionString=BBSDBFileName
db.Open


'BBS Setting の読み込み
Set rs_set=Server.CreateObject("ADODB.Recordset")
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'bbs_" & BBSQuery & "'",db,3,2


rs_set("SiteName") = "サイト名"
rs_set("SiteURL") = "http://" & Request.ServerVariables("HTTP_HOST") & "/"
rs_set("BBSName") = "掲示板"
rs_set("BBSComment") = "何か書いていってくださいね～"
rs_set("BaseURL") = "http://" & Request.ServerVariables("HTTP_HOST") & _
		Replace(Request.ServerVariables("SCRIPT_NAME"),"/setup.asp","/")
rs_set("BGColor") = "#ffffff"
rs_set("TextColor") = "#000000"
rs_set("LinkColor") = "#0000ff"
rs_set("aLinkColor") = "#ffff00"
rs_set("BorderColor") = "#888888"
rs_set("TitleColor") = "#ff0000"
rs_set("ViewCount") = 10
rs_set("CountFile") = "count.dat"
rs_set("Tag") = False
rs_set("TagSourceView") = False
rs_set("MailSend") = False
rs_set("MailServer") = "mail.example.com"
rs_set("SendToAddr") = "send@example.com"
rs_set("MailFromAddr") = "bbs@example.com"

rs_set("MailBBSBodyCut") = True
rs_set("NotFoundName") = "名無しさん"
rs_set("DelMailAddr") = "name@deleted"
rs_set("DelName") = "あぼーん"
rs_set("DelTitle") = "削除済み"
rs_set("DelBody") = "管理者により削除"
rs_set("DelDevType") = "system"
rs_set("DelTitleColor") = "#ff99aa"
rs_set("DelBodyColor") = "#ff0000"

rs_set.Update

rs_set.Close

Set rs_set = Nothing

'掲示板データベースの初期化処理終了


Set Fso = Nothing
%>
<b>掲示板設定の初期化処理が完了しました。</b><br>
<br>
<a href="admin.asp">掲示板設定ツールへ</a><br>
<br><%

End If

End If
%>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<font align="center" size="-2">&copy;2004 anoncom.net</font>
</body>
</html>