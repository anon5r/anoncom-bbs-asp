<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'�g�т̏ꍇ�͌g�їp�̊Ǘ��y�[�W�ֈړ�
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'�o�b�t�@��L���ɂ���
Response.Buffer = True

If BBSBlank = True Then
	Response.Redirect "blankbbs.html"
End If

'���O�C���Z�b�V�������m�����Ă��邩�m�F
If Session("login") = 1 Then

%>
<html lang="ja">
<head>
<meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/html; charset=shift_jis">
<title>�ݒ�ύX�c�[��</title>
<style>
<!--
table,td{
	color:#000000;
	font-size:10pt;
	font-family:"MS UI Gothic";
	border-style:solid;
	border-width:1px;
	border-color:#cccccc;
}
-->
</style>
</head>
<body>
<font color="#ff0000" face="Times New Roman" size="+3"><b><i>BBS Setting Tool for <%=BBSName %></i></b></font><br>
<%


If Request.Form("edit") <> "on" Then


	Select Case Request.QueryString("set")
		Case "bbs"
%>
<form action="setting.asp?bbs=<%=BBSQuery %>&set=bbs" method="post" enctype="application/x-www-form-urlencoded">
<!--<input type="hidden" name="bbs" value="<%=BBSQuery %>">-->
<input type="hidden" name="edit" value="on">
<table border="1" cellspacing="0">
<tbody>
<tr>
<td align="center" bgcolor="#88ff88">�ݒ荀��</td>
<td align="center" bgcolor="#ffaaaa">�ݒ�l</td>
<td align="center" bgcolor="#aaaaff">�����E���l</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�T�C�g���F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="sitename" value="<%=SiteName %>" maxlength="100" size="20"></td>
<td align="left" bgcolor="#ccccff">��100�����܂�</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�T�C�gURL�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="siteurl" value="<%=SiteURL %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�����p�p����</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f�����F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="bbsname" value="<%=BBSName %>" maxlength="100" size="20"></td>
<td align="left" bgcolor="#ccccff">��100�����܂�</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f���R�����g�F</td>
<td align="left" bgcolor="#ffcccc"><textarea name="bbscomment" rows="2" cols="20"><%=BBSComment %></textarea></td>
<td align="left" bgcolor="#ccccff">��100�����܂�</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f����URL�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="baseurl" value="<%=BBSURL %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�����p�p����</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f���̏�ԁF</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case BBSStatus
	Case 9: status9 = " selected"
	Case 3: status3 = " selected"
	Case 0: status0 = " selected"
End Select %><select name="bbsstatus">
<option value="9"<%=status9 %>>�{���E�������݉�</option>
<option value="3"<%=status3 %>>�������ݕs��</option>
<option value="0"<%=status0 %>>�{���E�������ݕs��</option>
</select></td>
<td align="left" bgcolor="#ccccff">�f���̉^�c��Ԃ�ύX���܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f�o�b�O���[�h�F</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case debug_flag
	Case 9: debug9 = " selected"
	Case 8: debug9 = " selected"
	Case 7: debug5 = " selected"
	Case 6: debug5 = " selected"
	Case 5: debug5 = " selected"
	Case 4: debug5 = " selected"
	Case 3: debug3 = " selected"
	Case 2: debug3 = " selected"
	Case 1: debug3 = " selected"
	Case 0: debug0 = " selected"
End Select %><select name="DebugMode">
<option value="0"<%=debug0 %>>�f�o�b�O�Ȃ�</option>
<option value="3"<%=debug3 %>>�ȈՃf�o�b�O</option>
<option value="5"<%=debug5 %>>�f�o�b�O</option>
<option value="9"<%=debug9 %>>���S�f�o�b�O</option>
</select></td>
<td align="left" bgcolor="#ccccff">�f�o�b�O���[�h�ł̎��s�ۂ�ύX���܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�w�i�F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="bgcolor" value="<%=BGColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">��F<i>#ffffff</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�����̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="textcolor" value="<%=TextColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">��F<i>#000000</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�����N�̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="linkcolor" value="<%=LinkColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">��F<i>#0000ff</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�A�N�e�B�u�����N�̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="alinkcolor" value="<%=ActiveLinkColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">��F<i>#ffff00</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="bordercolor" value="<%=BorderColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">��F<i>#888888</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f�����̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="titlecolor" value="<%=TitleColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">��F<i>#ff0000</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">1�y�[�W�̕\�������F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="cntnum" value="<%=CntNum %>" maxlength="3" size="3"></td>
<td align="left" bgcolor="#ccccff">�����p����</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�J�E���^�t�@�C�����F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="countfilename" value="<%=CountFileName %>"></td>
<td align="left" bgcolor="#ccccff">�J�E���^�t�@�C�������w�肵�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�f���ł̃^�O�̎g�p�F</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case TagUse
	Case 1: TUon = " selected"
	Case 0: TUoff = " selected"
End Select
%><select name="tag">
<option value="1"<%=TUon %>>�L��</option>
<option value="0"<%=TUoff %>>����</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">�����ɂ���Ə������݂��ꂽ�^�O�����ׂĖ����ɂȂ�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�^�O�̕\���F</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case TagSourceView
	Case 1: TSVon = " selected"
	Case 0: TSVoff = " selected"
End Select
%><select name="tagsourceview">
<option value="1"<%=TSVon %>>�\��</option>
<option value="0"<%=TSVoff %>>��\��</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">�\���ɂ���ƃ^�O�̃\�[�X���\������܂��B�i�^�O�g�p���u�����v�̏ꍇ�̂ݗL���j</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�������ݒʒm�z�M�F</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case BBSMailSend
	Case 1: BMSon = " selected"
	Case 0: BMSoff = " selected"
End Select
%><select name="bbsmailsend">
<option value="1"<%=BMSon %>>�L��</option>
<option value="0"<%=BMSoff %>>����</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">�L���ɂ���Ə������݂��w��̃A�h���X�֒ʒm�z�M����܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���[���T�[�o�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="mailserver" value="<%=MailServer %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�ʒm�z�M�𗘗p����ꍇ�͕K���w�肵�Ă��������I</td>
</tr><!--<tr>
<td align="right" bgcolor="#ccffcc">���M��A�h���X�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="sendtoaddr" value="<%=SendToAddr %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">���ʒm�z�M�̑��M��A�h���X</td>
</tr>--><tr>
<td align="right" bgcolor="#ccffcc">���M��O���[�v�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="SendGroup" value="<%=UserGroup %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">���ʒm�z�M�̑��M��O���[�v</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���M���A�h���X�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="mailfromaddr" value="<%=MailFromAddr %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�ʒm�z�M���p���ɁA�����Ŏw�肵���A�h���X���烁�[���������Ă��܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�ʒm���[���J�b�g�F</td>
<td align="left" bgcolor="#ffcccc"><%
	Select Case MailBodyCut
		Case 1: MBCon = " selected"
		Case 0: MBCoff = " selected"
	End Select
%><select name="mailbbsbodycut">
<option value="1"<%=MBCon %>>�L��</option>
<option value="0"<%=MBCoff %>>����</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">�L���ɂ���Ə������݂̖{���������ꍇ�A�{���ȗ����Ĕz�M����܂��Bi-mode�ŕ����ݒ肵�Ă��Ȃ��ꍇ�ɓ��ɕ֗��ł��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�������̖��O�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="notfoundname" value="<%=NotFoundName %>" maxlength="40" size="20"></td>
<td align="left" bgcolor="#ccccff">���O���Ȃ��ꍇ�A���O�����ɕ\�����镶������w�肵�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���O�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="delname" value="<%=DeleteName %>" maxlength="50" size="20"></td>
<td align="left" bgcolor="#ccccff">�폜���X�̖��O�����ɕ\�����镶������w�肵�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���[���A�h���X�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="delmailaddr" value="<%=DeleteMailAddr %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�폜���X�̃��[���A�h���X�ɕ\�����镶������w�肵�܂��B�w�肵�Ȃ��Ă����܂��܂���B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�^�C�g���F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="deltitle" value="<%=DeleteTilte %>" maxlength="100" size="40"></td>
<td align="left" bgcolor="#ccccff">�폜���X�̑���̃^�C�g�������ɕ\�����镶������w�肵�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�{���F</td>
<td align="left" bgcolor="#ffcccc"><textarea name="delbody" rows="2" cols="20"><%=DeleteBody %></textarea></td>
<td align="left" bgcolor="#ccccff">�폜���X�̖{�������ɕ\�����镶������w�肵�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�[���F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="deldevtype" value="<%=DeleteDeviceType %>"></td>
<td align="left" bgcolor="#ccccff">�폜���X�̒[�������ɕ\�����镶������w�肵�܂��B</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�^�C�g���̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="deltitlecolor" value="<%=DelTitleColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">�폜���X�̃^�C�g���̐F�B��F<i>#FF99AA</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�{���̐F�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="delbodycolor" value="<%=DelBodyColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">�폜���X�̖{���̐F�B��F<i>#ff0000</i></td>
</tr>
</tbody>
</table>
<br>
<input type="submit" value="�@�ρ@�@�X�@">
</form>
<%

		Case "admin"
%>
<form action="setting.asp?set=admin" method="post" enctype="application/x-www-form-urlencoded">
<input type="hidden" name="edit" value="on">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<table border="1" cellspacing="0">
<tbody>
<tr>
<td align="center" bgcolor="#88ff88">�ݒ荀��</td>
<td align="center" bgcolor="#ffaaaa">�ݒ�l</td>
<td align="center" bgcolor="#aaaaff">�����E���l</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�Ǘ���ID�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminid" value="<%=Session("AdminID") %>" maxlength="20" size="20"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��҃��O�C���pID</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�p�X���[�h�F</td>
<td align="left" bgcolor="#ffcccc"><input type="password" name="adminpass" value="<%=Session("AdminPass") %>" maxlength="10" size="16"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��p�p�X���[�h</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�Ǘ��Җ��F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminname" value="<%=Session("AdminName") %>" maxlength="20" size="30"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��Җ�</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���[���A�h���X�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminmail" value="<%=Session("AdminMail") %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��҃��[���A�h���X</td>
</tr>
</tbody>
</table>
<br>
<input type="submit" value="�@�ρ@�@�X�@">
</form>
<%

	End Select


'�|�X�g�� ****************************************************************
Else

   Select Case Request.QueryString("set")

	Case "bbs"
		'�f���ݒ�

		Set rs_set = Server.CreateObject("ADODB.Recordset")
		upSQL = "SELECT * FROM settings WHERE bbs_table = 'bbs_" & BBSQuery & "'"
		rs_set.Open upSQL,db,3,3

		'�ݒ�̔��f
		rs_set("SiteName") = Request.Form("SiteName")
		rs_set("SiteURL") = Request.Form("SiteURL")
		rs_set("BBSName") = Request.Form("BBSName")
		rs_set("BBSComment") = Request.Form("BBSComment")
		rs_set("BaseURL") = Request.Form("BaseURL")
		rs_set("act_flag") = CInt(Request.Form("BBSStatus"))
		rs_set("debug_flag") = CInt(Request.Form("DebugMode"))
		rs_set("BGColor") = Request.Form("BGColor")
		rs_set("TextColor") = Request.Form("TextColor")
		rs_set("LinkColor") = Request.Form("LinkColor")
		rs_set("aLinkColor") = Request.Form("aLinkColor")
		rs_set("BorderColor") = Request.Form("BorderColor")
		rs_set("TitleColor") = Request.Form("TitleColor")
		rs_set("ViewCount") = Request.Form("CntNum")
		rs_set("CountFile") = Request.Form("CountFileName")
		Select Case Request.Form("Tag")
			Case "1": rs_set("Tag") = True
			Case "0": rs_set("Tag") = False
		End Select
		Select Case Request.Form("TagSourceView")
			Case "1": rs_set("TagSourceView") = True
			Case "0": rs_set("TagSourceView") = False
		End Select
		Select Case Request.Form("BBSMailSend")
			Case "1": rs_set("MailSend") = True
			Case "0": rs_set("MAilSend") = False
		End Select
		rs_set("MailServer") = Request.Form("MailServer")
		'rs_set("SendToAddr") = Request.Form("SendToAddr")
		rs_set("groups") = Request.Form("SendGroup")
		rs_set("MailFromAddr") = Request.Form("MailFromAddr")
		Select Case Request.Form("MailBBSBodyCut")
			Case "1": rs_set("MailBBSBodyCut") = True
			Case "0": rs_set("MailBBSBodyCut") = False
		End Select
		rs_set("NotFoundName") = Request.Form("NotFoundName")
		rs_set("DelMailAddr") = Request.Form("DelMailAddr")
		rs_set("DelName") = Request.Form("DelName")
		rs_set("DelTitle") = Request.Form("DelTitle")
		rs_set("DelBody") = Request.Form("DelBody")
		rs_set("DelDevType") = Request.Form("DelDevType")
		rs_set("DelTitleColor") = Request.Form("DelTitleColor")
		rs_set("DelBodyColor") = Request.Form("DelBodyColor")

		rs_set.Update

		rs_set.Close
		Set rs_set = Nothing



	Case "admin"
		'�Ǘ��Ґݒ�

		Set rs_set = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM admin_settings WHERE " & _
			"AdminID = '" & Session("adminID") & "' AND " & _
			"AdminPass = '" & Session("adminPass") & "'"
		Set tmprs = db.Execute(SQL)
		If tmprs.EOF =False Then
			ser = tmprs("SERIAL")
			Set tmprs = Nothing

			upSQL = "SELECT * FROM admin_settings WHERE " & _
				"[SERIAL] = " & ser
			rs_set.Open upSQL,db,3,3

			'�ݒ�̔��f
			rs_set("AdminID") = Request.Form("adminID")
			rs_set("AdminPass") = Request.Form("adminPass")
			rs_set("AdminName") = Request.Form("adminName")
			rs_set("AdminMail") = Request.Form("adminMail")

			rs_set.Update

			rs_set.Close
			Set rs_set = Nothing

			'�Z�b�V�����Ǘ��ҏ����X�V
			Session("adminID") = Request.Form("adminID")
			Session("adminPass") = Request.Form("adminPass")
			Session("adminName") = Request.Form("adminName")
			Session("adminMail") = Request.Form("adminMail")


		Else
			%>���̊Ǘ��҂͌�����Ȃ��ׁA�X�V�ł��܂���ł����B<%
		End If


   End Select

%>
�ݒ肪�ύX����܂����B<br>
<a href="admin.asp" target="_top">[BBS Admin]</a>
<% End If %>
</body>
</html><%

Else
	Response.Redirect "admin.asp"
End If %>
