<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'�o�b�t�@�L��
Response.Buffer = True
%>
<html lang="ja">
<head>
<meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<title>Admin ���[�U�Ǘ�</title>
<style>
<!--
table,td{
	color:#000000;
	font-size:10pt;
	font-family:"MS UI Gothic";
	border-style:solid;
	border-width:0px;
	border-color:#cccccc;
}
-->
</style>
</head>
<body>
<font color="#ff0000" face="Times New Roman" size="+3"><b><i>Administrator user management</i></b></font><br>
<br>
<br>
<%
Set Fso = Server.CreateObject("Scripting.FileSystemObject")


'�g�т̏ꍇ�͌g�їp�̊Ǘ��y�[�W�ֈړ�
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'���O�C���Z�b�V�����͊m�����Ă��邩
If Session("login") = 1 Then





	If Session("AdminLevel") < 9 Then
%>���[�U�Ǘ�����������܂���B<%
	Else


		If Request("mode") = "" Then


		'���[�U�Ǘ����
%>
<form action="admusrmng.asp" method="post">
<input type="hidden" name="mode" value="add">
<b>[���[�U�̒ǉ�]</b><br>
�Ǘ���ID�F
<input type="text" name="uid" value="" maxlength="20">
<input type="submit" value="�A�J�E���g�쐬">�@�����p�p�����ő�20����<br>
�Ǘ��Җ��F
<input type="text" name="uname" value="" maxlength="50">
�@���ő�50����<br>
�p�X���[�h�F
<input type="password" name="pwd" value="" maxlength="20">
�@�����p�p�����ő�20����<br>
�Ǘ��҃����N�F
<select name="adminlevel">
<option value="8">���ō�����</option>
<option value="7">�f���Ǘ�+�쐬</option>
<option value="6" selected>�f���Ǘ���</option>
<option value="5">�i�����蓖�āj</option>
<option value="4">���f���Ǘ���</option>
<option value="3">��ʊǗ��҃��[�U</option>
<option value="2">�폜�l</option>
<option value="1">���O�{���̂�</option>
<option value="0">�A�J�E���g��~</option>
</select>
�@<a href="adminranklist.html">�Ǘ��҃����N�����ꗗ</a><br>
�@�������N9(�ō�����:root)�͎w��ł��܂���B<br>
�Ǘ��Ώیf���F<br>
<select name="bbs" multiple><%
SQL = "SELECT [bbs_table], [BBSName] FROM [settings] WHERE [bbs_table] Like 'bbs_%' ORDER BY [SERIAL] ASC"
Set rs = db.Execute(SQL)
Do While Not rs.EOF = True
	bbskey = Replace(rs("bbs_table"), "bbs_", "")
	bbsname = rs("BBSName")
%>
<option value="<%=bbskey %>"><%=bbsname %></option><%
	rs.MoveNext
Loop

Set rs = Nothing
%>
</select>
</form>
<hr size="1" bgcolor="<%=BorderColor %>">
<br>

<form action="admusrmng.asp" method="get">
<input type="hidden" name="mode" value="edit">
<b>[���[�U���̕ύX]</b><br>
�Ǘ��Җ��F<select name="adminID"><%
SQL = "SELECT [adminID], [adminName] FROM [admin_settings] ORDER BY [SERIAL] ASC"

Set rs = db.Execute(SQL)
Do While Not rs.EOF = True

	tmpAdminID = rs("adminID")
	tmpAdminName = rs("adminName")

%>
<option value="<%=tmpAdminID %>"><%=tmpAdminName %></option><%

	rs.MoveNext
Loop

Set rs = Nothing
%>
</select>
<input type="submit" value="�Ǘ��ҏ��ύX">
</form>
<%



		Else
			Select Case Request("mode")

				Case "add"
					'�|�X�g����
					If Request.Form <> "" Then

						If Request.Form("uid") = "" Then
							%>ID�����͂���Ă��܂���B<%
						ElseIf Request.Form("pwd") = "" Then
							%>�p�X���[�h�����͂���Ă��܂���B<%
						Else
							uid = Replace(Request.Form("uid"), "'", "''")
							uname = Replace(Request.Form("uname"), "'", "''")
							pwd = Request.Form("pwd")
							adminrank = Request.Form("adminlevel")

							tmpSQL = "SELECT * FROM admin_settings WHERE adminID = '" & uid & "'"
							Set tmprs = db.Execute(tmpSQL)
							If tmprs.EOF = True Then


								If adminrank >= 7 Then
									adminbbs = "allbbs"
								Else
									adminbbs = Replace(Request.Form("bbs"), ", ", ",")
								End If
								SQL = "INSERT INTO admin_settings (" & _
									"adminID, adminPass, adminName, adminRank, adminBBS" & _
									") VALUES(" & _
									"'" & uid & "', '" & pwd & "', '" & uname & "', " & _
									adminrank & ", '" & adminbbs & "')"
								db.Execute(SQL)

								SQL = "SELECT * FROM admin_settings WHERE adminID = '" & uid & "'"
								Set rs = db.Execute(SQL)
								If rs.EOF = True Then
									%>�Ǘ��ҁF<%=uid %> ��ǉ��ł��܂���ł����B<%
								Else
								%>
�Ǘ��ҁF<%=rs("adminID") %> ��ǉ����܂����B<br>
<br>
�Ǘ���ID�F<%=rs("adminID") %><br>
���O�F<%=rs("adminName") %><br>
�p�X���[�h�F<%=String(Len(rs("adminPass")), "*") %><br>
�Ǘ��҃����N�F<%=rs("adminRank") %><br>
�Ǘ��Ώیf���F<%=rs("adminBBS") %><br>
<%
								End If
							Else
								%>���̊Ǘ���ID�͊��ɑ��݂��܂��B<%
							End If
						End If
					End If

				Case "edit"
					If Request.Form = "" Then
						'�|�X�g�O

						admSQL = "SELECT * FROM [admin_settings] " & _
							"WHERE [adminID] = '" & Request.QueryString("adminID") & "'"
						Set admrs = db.Execute(admSQL)

						Session("tmpSer") = admrs("SERIAL")
						tmpAdminID = admrs("adminID")
						tmpAdminPass = admrs("adminPass")
						tmpAdminName = admrs("adminName")
						tmpAdminMail = admrs("adminMail")
						tmpAdminRank = admrs("adminRank")
						tmpAdminBBS = admrs("adminBBS")

						%>
<form action="admusrmng.asp" method="post" enctype="application/x-www-form-urlencoded">
<input type="hidden" name="mode" value="edit">
<table border="1" cellspacing="0">
<tr>
<td align="center" bgcolor="#88ff88">�ݒ荀��</td>
<td align="center" bgcolor="#ffaaaa">�ݒ�l</td>
<td align="center" bgcolor="#aaaaff">�����E���l</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�Ǘ���ID�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminID" value="<%=tmpAdminID %>" maxlength="20" size="20"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��҃��O�C���pID</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�p�X���[�h�F</td>
<td align="left" bgcolor="#ffcccc"><input type="password" name="adminPass" value="<%=tmpAdminPass %>" maxlength="10" size="16"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��p�p�X���[�h</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�Ǘ��Җ��F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminName" value="<%=tmpAdminName %>" maxlength="20" size="30"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��Җ�</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">���[���A�h���X�F</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminMail" value="<%=tmpAdminMail %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">�Ǘ��҃��[���A�h���X</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�Ǘ��҃����N�F</td>
<td align="left" bgcolor="#ffcccc">
<select name="adminLevel"><%
						Select Case tmpAdminRank
							Case 9 : tmpRank9 = " selected"
							Case 8 : tmpRank8 = " selected"
							Case 7 : tmpRank7 = " selected"
							Case 6 : tmpRank6 = " selected"
							Case 5 : tmpRank5 = " selected"
							Case 4 : tmpRank4 = " selected"
							Case 3 : tmpRank3 = " selected"
							Case 2 : tmpRank2 = " selected"
							Case 1 : tmpRank1 = " selected"
							Case 0 : tmpRank0 = " selected"
						End Select
%>
<option value="9"<%=tmpRank9 %>>�ō�����</option>
<option value="8"<%=tmpRank8 %>>���ō�����</option>
<option value="7"<%=tmpRank7 %>>�f���Ǘ�+�쐬</option>
<option value="6"<%=tmpRank6 %>>�f���Ǘ���</option>
<option value="5"<%=tmpRank5 %>>�i�����蓖�āj</option>
<option value="4"<%=tmpRank4 %>>���f���Ǘ���</option>
<option value="3"<%=tmpRank3 %>>��ʊǗ��҃��[�U</option>
<option value="2"<%=tmpRank2 %>>�폜�l</option>
<option value="1"<%=tmpRank1 %>>���O�{���̂�</option>
<option value="0"<%=tmpRank0 %>>�A�J�E���g��~</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">�Ǘ��Ҍ����@��<a href="adminranklist.html">�����ꗗ</a></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">�Ǘ��Ώیf���F</td>
<td align="left" bgcolor="#ffcccc">
<select name="adminBBS" multiple><%
						SQL = "SELECT [bbs_table], [BBSName] FROM [settings] " & _
							"WHERE [bbs_table] Like 'bbs_%' " & _
							"ORDER BY [SERIAL] ASC"
						Set rs = db.Execute(SQL)

						Do While Not rs.EOF = True

							'�f�����[�h
							bbskey = Replace(rs("bbs_table"), "bbs_", "")
							bbsname = rs("BBSName")

							'�Ǘ��Ώیf���Ăяo��
							If tmpAdminBBS = "allbbs" Then
								%><option value="<%=bbskey %>" selected><%=bbsname %></option><%
							Else
								aryAdmRank = Split(tmpAdminBBS, ",")
								selctd = ""
								For i = 0 To UBound(aryAdmRank)
									If aryAdmRank(i) = bbskey then
										selctd = " selected"
									End If
								Next

								%><option value="<%=bbskey %>"<%=selctd %>><%=bbsname %></option><%

							End If

							rs.MoveNext

						Loop

						Set rs = Nothing
%>
</select>
</td>
<td align="left" bgcolor="#ccccff">�Ǘ��\�Ȍf����</td>
</tr>
</table>
<br>
<input type="submit" value="�@�ρ@�@�X�@">
</form>
<%
					Else
						'�|�X�g��
						tmpSQL = "SELECT * FROM [admin_settings] " & _
							"WHERE [SERIAL] = " & Session("tmpSer")
						Set rs_set = Server.CreateObject("ADODB.Recordset")
						Set tmprs = db.Execute(tmpSQL)
						If tmprs.EOF = False Then
							ser = Session("tmpSer")
							Set tmprs = Nothing

							upSQL = "SELECT * FROM admin_settings WHERE " & _
								"[SERIAL] = " & ser
							rs_set.Open upSQL,db,3,3

							'�ݒ�̔��f
							rs_set("AdminID") = Request.Form("adminID")
							rs_set("AdminPass") = Request.Form("adminPass")
							rs_set("AdminName") = Request.Form("adminName")
							rs_set("AdminMail") = Request.Form("adminMail")
							rs_set("AdminRank") = Request.Form("adminLevel")
							If Request.Form("adminLevel") >= 7 Then
								rs_set("AdminBBS") = "allbbs"
							Else
								rs_set("AdminBBS") = Replace(Request.Form("adminBBS"), " ", "")
							End If

							rs_set.Update

							rs_set.Close
							Set rs_set = Nothing

							%>
<%=Request.Form("adminName") %>�i<%=Request.Form("adminID") %>�j�̏����X�V���܂����B<br>
<a href="admin.asp" target="_top">[�f���Ǘ����]</a><%
						Else
							%>���̊Ǘ��҂͌�����Ȃ��ׁA�X�V�ł��܂���ł����B<%
						End If
					End If

				Case Else
						%>���̋@�\�͍쐬���ł��B<%
			End Select
		End If
	End If
Else
	'���O�C������Ă��Ȃ�
	Response.Redirect "admin.asp"
End If
%>
</body>
</html>