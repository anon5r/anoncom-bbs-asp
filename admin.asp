<%@ Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'�g�т̏ꍇ�͌g�їp�̊Ǘ��y�[�W�ֈړ�
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'�o�b�t�@��L���ɂ���
Response.Buffer = True


'���O�C���O�̓J���[���Ȃǂ͕W���ݒ���g��

If Session("login") <> 1 And Request.Form = "" Then	'���O�C���O�̂ݓK��

Set rs_set = Server.CreateObject("ADODB.Recordset")
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'default_settings'",db,3,2

'�W���̐ݒ��ǂݍ���
BBSBlank = True
SiteURL = rs_set("SiteURL")
BBSURL = rs_set("BaseURL")
BGColor = rs_set("BGColor")
TextColor = rs_set("TextColor")
LinkColor = rs_set("LinkColor")
ActiveLinkColor = rs_set("aLinkColor")
HoverLinkcolor = ActiveLinkColor
BorderColor = rs_set("BorderColor")
TitleColor = rs_set("TitleColor")

rs_set.Close
Set rs_set = Nothing

End If

%>
<html lang="ja">
<head>
<title>BBS Administrator</title>
<meta http-equiv="Content-Type" content="text/html; charset=x-sjis">
<meta http-equiv="Content-Style-Type" content="text/css; charset=x-sjis">
<style type="text/css">
<!--
a:link { text-decoration:none; color:<%=LinkColor %>; }
a:visited { text-decoration:none; color:<%=LinkColor %>; }
a:active { text-decoration:none; color:<%=ActiveLinkColor %>; }
a:hover { text-decoration:underline; color:<%=HoverLinkColor %>; cursor:default; }

body{
	background-color:<%=BGColor %>;
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
	overflow-y:auto;
	scrollbar-base-color:<%=BGColor %>;
	scrollbar-face-color:<%=BGColor %>;
	scrollbar-arrow-color:<%=BorderColor %>;
	scrollbar-highlight-color:<%=BGColor %>;
	scrollbar-3dlight-color:<%=BorderColor %>;
	scrollbar-shadow-color:<%=BorderColor %>;
	scrollbar-darkshadow-color:<%=BGColor %>;
}

table{
	border-style:solid;
	border-width:0px;
}

td{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
}

input{
	color:<%=TextColor %>;
	background-color:<%=BGColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic';
	border-style:solid;
	border-width:1px;
	border-color:<%=BorderColor %>;
}

-->
</style>
</head>
<%
If Session("login") <> 1 Then

	If Request.Form = "" Then	'���O�C���|�X�g�O

'���O�C�����

'�W���̐ݒ��ǂݍ���
BBSBlank = True
SiteName = "anoncomBBS"
SiteURL = "http://" & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptName, "")
BBSURL = "http://" & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptName, "")
BGColor = "#ffffff"
TextColor = "#000000"
LinkColor = "#0000ff"
ActiveLinkColor = "#ff0000"
HoverLinkcolor = ActiveLinkColor
BorderColor = "#888888"
TitleColor = "#ff0000"
%>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<font color="<%=TitleColor %>" size="+2">BBS Administrator</font><br>
<br>
<form action="admin.asp" method="post">
�Ǘ���ID�F<input type="text" name="id" size="8"><br>
Password�F<input type="password" name="pw" size="8"><br>
<input type="submit" value="���O�C��">
</form>
<br>
<a href="bbs.asp">�f����</a><br>
<hr size="1">
system by <a href="http://anoncom.net/">anoncom.net</a>
</body><%

	Else
	'���O�C������
		'�Ǘ����ǂݍ���
		Set rs_admin = Server.CreateObject("ADODB.Recordset")
		rs_admin.Open "SELECT * FROM admin_settings " & _
			"WHERE adminID = '" & Request.Form("id") & "' AND " & _
			"adminPass = '" & Request.Form("pw") & "'", db, 3, 2

		'�F�؊���
		If  rs_admin.EOF = False Then

			Session("AdminID") = rs_admin("adminID")
			Session("AdminPass") = rs_admin("adminPass")
			Session("AdminName") = rs_admin("adminName")
			Session("AdminMail") = rs_admin("adminMail")
			Session("AdminLevel") = CInt(rs_admin("adminRank"))
			Session("AdminBBS") = rs_admin("adminBBS")
			rs_admin.Close

			Session("login") = 1
			'Session("bbsquery") = Request.Form("bbs")
			Session.TimeOut = 30	'�^�C���A�E�g��30��
		End If

		Set rs_admin = Nothing

		Response.Redirect "admin.asp"

	End If

Else
'�Z�b�V�������m�F���A���O�C�����Ă�����
	If Request("bbs") <> "" Then
		Session("BBSQuery") = Request("bbs")
		Session("BBSNo") = CInt(0)
		Response.Redirect "admin.asp"
	End If

	BBSQuery = Session("BBSQuery")
	BBSSelectNo = Session("BBSNo")


	If Session("AdminLevel") = 0 Then
		'AdminRank=0�́A���O�C�����̃Z�b�V���������ׂĔj�����A���O�C���r��
		BBSBlank = True
		SiteName = "anoncomBBS"
		ScriptName = "admin.asp"
		SiteURL = "http://" & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptName, "")
		BBSURL = "http://" & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptName, "bbs.asp")
		BGColor = "#ffffff"
		TextColor = "#000000"
		LinkColor = "#0000ff"
		ActiveLinkColor = "#ff0000"
		HoverLinkcolor = ActiveLinkColor
		BorderColor = "#888888"
		TitleColor = "#ff0000"
		Session.Abandon
%>
<head>
<meta http-equiv="refresh" content="15;url=<%=BBSURL %>">
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
����ID�͌��ݒ�~���u���{����Ă��邽�߁A�Ǘ���ʂ�\�����邱�Ƃ͂ł��܂���B<br>
�����͎��̂��̂��l�����܂��F<br>
<br>
�E�Ǘ��҂̗��p�����ɂ��A�J�E���g�̈ꎟ���p��~<br>
�E�����ȃ��[�U�[����̌������D<br>
�E�f���Ǘ��@�\�̃f�o�b�O�̂��߂̈ꎞ�I�Ȓ�~���u<br>
<br>
<br>
�Ȃ��A���̃y�[�W��15�b��Ɍf���̃g�b�v�ֈړ����܂��B<br>
</body>
<%
	Else

		Select Case Request.QueryString

			'���j���[���
			Case "menu"
%>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<b><i><font color="<%=TitleColor %>" size="+2" face="Times New Roman">BBS Admin</font></i></b><br>
BBS ID: <%=BBSQuery %><br>
<br>
<table border="0" height="88%" width="100%">
<tr>
<td align="center" valign="top">

<table border="0">
<tr>
<td align="left"><form action="admin.asp" method="get" target="_top">
<b>�f���F</b><br>
<select name="bbs" onchange="this.form.submit()">
<option value="">(�f����...)</option>
<%

SQL = "SELECT [bbs_table], [BBSName] FROM [settings] WHERE [bbs_table] Like 'bbs_%' ORDER BY [SERIAL] ASC"
Set rs = db.Execute(SQL)

'���[�v�J�E���^�����l
i = 1

If rs.EOF = False Then
	Do While Not rs.EOF
		bbstable = Replace(rs("bbs_table"), "bbs_", "")
		BBSName = rs("BBSName")

		If Session("AdminBBS") = "allbbs" Then

			'�J�E���^�ƌf���V���A��No����v�����select
			If i = BBSSelectNo Then
				BBSSelectValue = " selected"
			Else
				BBSSelectValue = ""
			End If
			%><option value="<%=bbstable %>"<%=BBSSelectValue %>><%=BBSName %></option><%
		Else
			aryAdminBBS = Split(Session("AdminBBS"), ",")
			For x = 0 To UBound(aryAdminBBS)
				If bbstable = aryAdminBBS(x) Then
					%><option value="<%=bbstable %>"<%=BBSSelectValue %>><%=BBSName %></option><%
				End If
			Next
		End If
		rs.MoveNext
		i = i + 1
	Loop
Else
%>
<option value="">�f�����쐬����Ă��܂���</option><%
End If
%>
</select>
<noscript><input type="submit" value="�I��"></noscript></form></td>
</tr><%
	If BBSQuery <> "" Then

%><tr>
<td align="center">
<b><a href="admin.asp?main" target="adminmain">�g�b�v</a></b>
</td>
</tr><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=bbs" target="adminmain">�f����</a></b>
</td>
</tr><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=access" target="adminmain">�������݉��</a></b>
</td>
</tr><%
		If Session("AdminLevel") > 1 Then
%><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=delete" target="adminmain">�폜�Ǘ�</a></b>
</td>
</tr><%
		End If
%><%
		If Session("AdminLevel") > 3 Then
%><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=setting" target="adminmain">�f���ݒ�</a></b>
</td>
</tr><%
		End If
%><%
		If Session("AdminLevel") > 4 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=clear" target="adminmain"><font color="<%=TitleColor %>">�������ݑS����</font></a>
</td>
</tr><tr>
<td align="center">&nbsp;</td>
</tr><tr>
<td align="center">
<a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=settingclear" target="adminmain">
<font color="<%=TitleColor %>">�ݒ�̏�����</font></a>
</td>
</tr><%
		End If



	End If

%><tr>
<td align="center">&nbsp;</td>
</tr><%
	If Session("AdminLevel") > 6 Then
%><tr>
<td align="center"><a href="setbbs.asp?mode=create" target="adminmain">�V�K�f���쐬</a></td>
</tr><%
	End If
%><%
	If Session("AdminLevel") > 7 Then
%><tr>
<td align="center"><a href="setbbs.asp?mode=delete" target="adminmain">�f���폜</a></td>
</tr><%
	End If
%><%
	If Session("AdminLevel") > 2 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=adminsetting" target="adminmain">�Ǘ��ҏ��ύX</a>
</td>
</tr><%
		End If
%><%
	If Session("AdminLevel") >= 9 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=adminuserreg" target="adminmain">�Ǘ��҃��[�U�Ǘ�</a>
</td>
</tr><%
	End If
%><tr>
<td align="center">&nbsp;</td>
</tr><%
	If Session("AdminLevel") >= 9 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=dbmove" target="adminmain">
<font color="<%=TitleColor %>">�f�[�^�x�[�X�ړ�</font></a>
</td>
</tr><%
	End If

%><tr>
<td align="center">
<a href="admin.asp?logout" target="_top">���O�A�E�g</a>
</td>
</tr>
</table>

</td>
</tr><tr>
<td align="center" valign="bottom">
&copy;2004 <a href="http://anoncom.net/" target="_top">anoncom.net</a>
</tr>
</table>
</body>
<%
'�g�b�v���C�����
Case "main"
%>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<font color="<%=TitleColor %>" face="Times New Roman"><b><i>anoncomBBS</i></b></font><br>
<font color="<%=TitleColor %>" size="+3" face="Times New Roman"><b><i>BBS Administrator for <%=BBSName %></i></b></font><br>
<br>
<br>
<br>
���̃��j���[����A�ݒ肷�鍀�ڂ��N���b�N���Ă��������B
</body>
<%
'���O�A�E�g����
Case "logout"

	Session("login") = 0
	Session("AdminID") = ""
	Session("AdminPass") = ""
	Session("AdminName") = ""
	Session("AdminMail") = ""
	Session("AdminLevel") = ""
	Session("BBSQuery") =""
	Session.TimeOut = 1
	Response.Redirect "admin.asp"

'URL��page�N�G�����Ȃ��ꍇ
Case Else

If BBSBlank = True Then
	Response.Redirect "nobbs.html"
End If

%>
<frameset border="1" framespacing="3" cols="180,*" frameborder="1">
<frame scrolling="yes" src="admin.asp?menu" name="adminmenu">
<frame scrolling="yes" src="admin.asp?main" name="adminmain">
</frameset>
<%
End Select

End If

End If
%></html>
