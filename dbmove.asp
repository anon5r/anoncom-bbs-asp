<% @Language = "VBScript" %>
<!-- #Include file="bbsdb.asp" -->
<%
'�o�b�t�@��L���ɂ���
Response.Buffer = True

If Session("login") <> 1 Then
	Response.Redirect "admin.asp"
End If
%>
<html lang="ja">
<head>
<title>anoncomBBS DB Move</title>
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
anoncomBBS DB Move
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
���݂̏ꏊ����f���̃f�[�^�x�[�X���ړ����܂��B��낵���ł����H<br>
<br>
���݂̏ꏊ�F�@<%=BBSDBFileName %>
</td>
</tr><tr>
<td align="center" height="50">&nbsp;</td>
</tr><tr>
<td align="center" valign="bottom">
<form action="dbmove.asp" method="post">
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

Set Fso = Server.CreateObject("Scripting.FileSystemObject")


'BBS�i�[�t�H���_���̐���
TempName1 = Fso.GetTempName
TempName1 = Replace(TempName1, ".tmp", "")
TempName1 = Replace(TempName1, "rad", "")

TempName2 = Fso.GetTempName
TempName2 = Replace(TempName2, ".tmp", "")
TempName2 = Replace(TempName2, "rad", "")

TempName = TempName1 & TempName2

'�t�H���_����
Fso.CreateFolder Server.MapPath("./" & TempName)

'�f�[�^�x�[�X�̃R�s�[
Fso.CopyFile BBSDBFileName, Server.MapPath("./" & TempName & "/bbs.mdb")


'bbsdb.asp�̏�������

bbsdb = "[%" & vbCrLf & _
	"'Access�f�[�^�x�[�X�t�@�C����" & vbCrLf & vbCrLf & _
	"DBFileName = ""./" & TempName & "/bbs.mdb""" & vbCrLf & vbCrLf & _
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

'�������g�̃p�X���擾
Set Fl = Fso.GetFile(Server.MapPath("./dbmove.asp"))
SelfDir = Fl.ParentFolder
Set Fl = Nothing

'bbs.mdb�̃p�X���擾
BBSDBDir = Replace(BBSDBFileName, "\bbs.mdb", "")

'dbmove.asp��bbs.mdb���������[�g�ɂ��邩����
RootBool = CBool(BBSDBDir = SelfDir)

If RootBool = True Then
'�����ꍇ
	'bbs.mdb�������폜
	Fso.DeleteFile BBSDBFileName, True
Else
'�Ⴄ�ꍇ
	Fso.DeleteFile BBSDBFileName, True
	Fso.DeleteFolder BBSDBDir
End If

Set Fso = Nothing

'bbs.mdb�̈ړ���
BBSDBMoveFile = Server.MapPath("./" & TempName & "/bbs.mdb")


'�Ǘ��҃��[�h�ŋ������O�C��
Session("login") = 1
Session.TimeOut = 60

%>
<b>�ړ����������܂����B</b><br>
<br>
<b>�ړ���F<%=BBSDBMoveFile %></b>
<br>
<a href="admin.asp" target="_top">�f���ݒ�c�[����</a><br>
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