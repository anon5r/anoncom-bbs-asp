<% @Language = "VBScript" %>
<!-- #Include file="bbsdb.asp" -->
<%
'�o�b�t�@��L���ɂ���
Response.Buffer = True


If Session("login") <> 1 Then
	'���O�C������Ă��Ȃ����
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
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
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
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
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

'�a�a�r�������t�@�C��

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
�t�H���_����anoncomBBS�̐ݒ�����������܂��B��낵���ł����H<br>
<u>�f���̏������ݓ��e�͏���������܂���B</u><br>
�����̍�Ƃ͎��������Ƃ��o���܂���B
</td>
</tr><tr>
<td align="center" height="50">&nbsp;</td>
</tr><tr>
<td align="center" valign="bottom">
<form action="setup.asp" method="post">
<input type="hidden" name="setup" value="execute">
<input type="submit" value="�͂�">
</td>
<td align="center" valign="bottom">
<input type="button" value="������" onClick="javascript:history.back(-1)">
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

'���������s


'�f�[�^�x�[�X�ڑ�
Set db=Server.CreateObject("ADODB.Connection")

db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Mode = 3
db.ConnectionString=BBSDBFileName
db.Open


'BBS Setting �̓ǂݍ���
Set rs_set=Server.CreateObject("ADODB.Recordset")
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'bbs_" & BBSQuery & "'",db,3,2


rs_set("SiteName") = "�T�C�g��"
rs_set("SiteURL") = "http://" & Request.ServerVariables("HTTP_HOST") & "/"
rs_set("BBSName") = "�f����"
rs_set("BBSComment") = "���������Ă����Ă��������ˁ`"
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
rs_set("NotFoundName") = "����������"
rs_set("DelMailAddr") = "name@deleted"
rs_set("DelName") = "���ځ[��"
rs_set("DelTitle") = "�폜�ς�"
rs_set("DelBody") = "�Ǘ��҂ɂ��폜"
rs_set("DelDevType") = "system"
rs_set("DelTitleColor") = "#ff99aa"
rs_set("DelBodyColor") = "#ff0000"

rs_set.Update

rs_set.Close

Set rs_set = Nothing

'�f���f�[�^�x�[�X�̏����������I��


Set Fso = Nothing
%>
<b>�f���ݒ�̏������������������܂����B</b><br>
<br>
<a href="admin.asp">�f���ݒ�c�[����</a><br>
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