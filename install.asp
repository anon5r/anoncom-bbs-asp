<% @Language = "VBScript" %>
<!-- #Include file="bbsdb.asp" -->
<%
'バッファを有効にする
Response.Buffer = True

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
'####		ver. 1.8		####
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
anoncomBBSの初期設定を行います。よろしいですか？<br>
</td>
</tr><tr>
<td align="center" height="50">&nbsp;</td>
</tr><tr>
<td align="center" valign="bottom">
<form action="install.asp" method="post">
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
Set db = Server.CreateObject("ADODB.Connection")

db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Mode = 3
db.ConnectionString = BBSDBFileName
db.Open


'BBS Setting の読み込み
Set rs_set=Server.CreateObject("ADODB.Recordset")
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'default_settings'", db, 3, 2

'デフォルト設定
rs_set("act_flag") = 9
rs_set("debug_flag") = 0
rs_set("SiteName") = "サイト名"
rs_set("SiteURL") = "http://" & Request.ServerVariables("HTTP_HOST") & "/"
rs_set("BBSName") = "新規掲示板"
rs_set("BBSComment") = "コメント"
rs_set("BaseURL") = "http://" & Request.ServerVariables("HTTP_HOST") & _
		Replace(Request.ServerVariables("SCRIPT_NAME"),"/install.asp","/")
rs_set("BGColor") = "#ffffff"
rs_set("TextColor") = "#000000"
rs_set("LinkColor") = "#0000ff"
rs_set("aLinkColor") = "#ffff00"
rs_set("BorderColor") = "#888888"
rs_set("TitleColor") = "#ff0000"
rs_set("ViewCount") = 10
rs_set("CountFile") = "bbscnt.dat"
rs_set("Tag") = True
rs_set("TagSourceView") = False
rs_set("MailSend") = False
rs_set("MailServer") = "mail.example.com"
rs_set("SendToAddr") = "send@example.com"
rs_set("MailFromAddr") = "bbs@example.com"
rs_set("groups") = "admin"

rs_set("MailBBSBodyCut") = False
rs_set("NotFoundName") = "名無しさん"
rs_set("DelMailAddr") = "name@deleted"
rs_set("DelName") = "削除"
rs_set("DelTitle") = "削除済み"
rs_set("DelBody") = "管理者により削除"
rs_set("DelDevType") = "system"
rs_set("DelTitleColor") = "#ff99aa"
rs_set("DelBodyColor") = "#ff0000"

rs_set.Update

rs_set.Close


'新規掲示板
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'bbs_default'", db, 3, 2


rs_set("act_flag") = 9
rs_set("debug_flag") = 0
rs_set("SiteName") = "サイト名"
rs_set("SiteURL") = "http://" & Request.ServerVariables("HTTP_HOST") & "/"
rs_set("BBSName") = "新規掲示板"
rs_set("BBSComment") = "コメント"
rs_set("BaseURL") = "http://" & Request.ServerVariables("HTTP_HOST") & _
		Replace(Request.ServerVariables("SCRIPT_NAME"),"/install.asp","/")
rs_set("BGColor") = "#ffffff"
rs_set("TextColor") = "#000000"
rs_set("LinkColor") = "#0000ff"
rs_set("aLinkColor") = "#ffff00"
rs_set("BorderColor") = "#888888"
rs_set("TitleColor") = "#ff0000"
rs_set("ViewCount") = 10
rs_set("CountFile") = "bbscnt_default.dat"
rs_set("Tag") = True
rs_set("TagSourceView") = False
rs_set("MailSend") = False
rs_set("MailServer") = "mail.example.com"
rs_set("SendToAddr") = "send@example.com"
rs_set("MailFromAddr") = "bbs@example.com"
rs_set("groups") = "admin"

rs_set("MailBBSBodyCut") = False
rs_set("NotFoundName") = "名無しさん"
rs_set("DelMailAddr") = "name@deleted"
rs_set("DelName") = "削除"
rs_set("DelTitle") = "削除済み"
rs_set("DelBody") = "管理者により削除"
rs_set("DelDevType") = "system"
rs_set("DelTitleColor") = "#ff99aa"
rs_set("DelBodyColor") = "#ff0000"

rs_set.Update

rs_set.Close

'管理情報も初期化
rs_set.Open "SELECT * FROM admin_settings WHERE SERIAL = 1", db, 3, 2

rs_set("adminName") = "掲示板管理者"
rs_set("adminID") = "admin"
rs_set("adminPass") = "1234"

rs_set.Update

rs_set.Close

Set rs_set = Nothing

db.Close
Set db = Nothing
'掲示板データベースの初期化処理終了




'default.aspの設定

Set Fso = Server.CreateObject("Scripting.FileSystemObject")

'default.dist.aspをdefault.aspとして上書きコピーする。
Fso.CopyFile Server.MapPath("./default.dist.asp"), Server.MapPath("./default.asp"), True


'カウンタファイル設置
If Fso.FolderExists(Server.MapPath("./") & "\count") = False Then
	Fso.CreateFolder Server.MapPath("./count")
End If

If Fso.FileExists(Server.MapPath("./count/bbscnt_default.dat")) = False Then
	Fso.CreateTextFile Server.MapPath("./count/bbscnt_default.dat")
	Set Txt = Fso.OpenTextFile(Server.MapPath("./count/bbscnt_default.dat"), 2)
	Txt.Write "0"
	Txt.Close
	Set Txt = Nothing
End If

'インストール時には管理者モードで強制ログイン
Session("login") = 1
Session("AdminLevel") = 9
Session("AdminID") = "admin"
Session("AdminPass") = "1234"
Session("AdminName") = "管理者"
Session("AdminMail") = "admin@example.com"
Session("AdminBBS") = "allbbs"
Session.TimeOut = 60

'BBS格納フォルダ名の生成
TempName1 = Fso.GetTempName
TempName1 = Replace(TempName1, ".tmp", "")
TempName1 = Replace(TempName1, "rad", "")

TempName2 = Fso.GetTempName
TempName2 = Replace(TempName2, ".tmp", "")
TempName2 = Replace(TempName2, "rad", "")

TempName = TempName1 & TempName2

'フォルダ生成
If Fso.FolderExists(Server.MapPath("./" & TempName)) = False Then
	Fso.CreateFolder Server.MapPath("./" & TempName)
End If

'データベースの移動
Fso.CopyFile Server.MapPath("./bbs.mdb"), _
		Server.MapPath("./" & TempName & "/bbs.mdb")


'bbsdb.aspの書き換え

bbsdb = "[%" & vbCrLf & _
	"'Accessデータベースファイル名" & vbCrLf & vbCrLf & _
	"DBFileName = ""./" & TempName & "/bbs.mdb""" & vbCrLf & _
	vbCrLf & _
	"BBSDBFileName = Server.MapPath(DBFileName)" & vbCrLf & _
	"%]" & vbCrLf

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
bbsdb = Replace(bbsdb, "[", "<")
bbsdb = Replace(bbsdb, "]", ">")
Set Fl = Fso.GetFile(Server.MapPath("./bbsdb.asp"))
Set Ts = Fl.OpenAsTextStream(ForWriting, TristateUseDefault)
Ts.Write bbsdb
Ts.Close
Set Ts = Nothing

'最後にinstaller自身も削除
Fso.DeleteFile Server.MapPath("./install.asp"), True
Fso.DeleteFile Server.MapPath("./bbs.mdb"), True
Set Fso = Nothing
%>
<b>掲示板の初期設定が完了しました。</b><br>
<br>
<a href="admin.asp">掲示板設定ツールへ</a><br>
<br><%
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