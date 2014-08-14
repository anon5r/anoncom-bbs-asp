<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'�o�b�t�@�L��
Response.Buffer = True

Set Fso = Server.CreateObject("Scripting.FileSystemObject")


'�g�т̏ꍇ�͌g�їp�̊Ǘ��y�[�W�ֈړ�
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'���O�C���Z�b�V�����͊m�����Ă��邩
If Session("login") = 1 Then

Select Case Request("mode")

	Case "create"
		title_j = "�V�K�f���쐬"
		title_e = "Create New BBS"

	Case "delete"
		title_j = "�f���폜"
		title_e = "Delete BBS"

	Case Else
		Response.Redirect "admin.asp"
End Select


Set rs_set = db.Execute("SELECT * FROM settings WHERE bbs_table = 'default_settings'")

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

Set rs_set = Nothing

%>
<html lang="ja">
<head>
<meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<title><%=title_j %></title>
<style>
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

table,td{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:"MS UI Gothic";
	border-style:solid;
	border-width:0px;
	border-color:<%=BorderColor %>;
}
-->
</style>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<font color="<%=TitleColor %>" face="Times New Roman" size="+3"><b><i><%=title_e %></i></b></font><br>
<br>
<%
	'�|�X�g�����O
	If Request.Form = "" Then

Select Case Request.QueryString("mode")
	Case "create"

	'�f���쐬�����͂��邩
	If CInt(Session("AdminLevel")) <= 6 Then
	%><b>���̃��[�U�͌f���쐬����������܂���B</b>
	<%
	Else
					'�f���쐬
	%>
<form action="setbbs.asp" method="post">
<input type="hidden" name="mode" value="create">
<table border="0">
<tr>
<td align="right" valign="top">�V�K�ɍ쐬������BBS ID����͂��Ă��������F</td>
<td align="left" valign="top">
<input type="text" name="bbsid" size="20" maxlength="16">
<input type="submit" value="�쐬">
</td>
<td align="left" valign="top"><font color="#ff0000">�����p�p�����@�ő�16����</font></td>
</tr>
</table>
</form>

<%
	End If

	Case "delete"
	'�f���폜�����͂��邩
	If CInt(Session("AdminLevel")) <= 7 Then
	%><b>���̃��[�U�͌f���폜����������܂���B</b>
	<%
	Else
		'�f���폜
%>
<script language="JavaScript"><!--
function DeleteBBS(){
  msgRet = confirm("�{���ɂ�낵���ł����H");
  if ( msgRet == true ){
	document.delform.submit();
  }
}
// --></script>

<form action="setbbs.asp" name="delform" method="post">
<input type="hidden" name="mode" value="delete">
<table border="0">
<tr>
<td align="right" valign="top">�폜����BBS ID��I�����Ă��������F</td>
<td align="left" valign="top">
<select name="bbsid">
<option value="">---�I�����Ă�������---</option><%

SQL = "SELECT bbs_table, BBSName FROM settings WHERE bbs_table Like 'bbs_%'"
Set rs = db.Execute(SQL)

If rs.EOF = False Then
	Do While Not rs.EOF

		bbstable = Replace(rs("bbs_table"), "bbs_", "")
		BBSName = rs("BBSName")
%>
<option value="<%=bbstable %>"><%=BBSName %>(<%=bbstable %>)</option><%
		rs.MoveNext
	Loop
Else
%>
<option value="">�f��������܂���</option><%
End If
%>
</select>
<input type="button" value="�폜" onclick="DeleteBBS()">
</td>
</tr>
</table>
</form>
	<%
	End If


	End Select
Else
		'�|�X�g��

	BBSID = Request.Form("bbsid")

Select Case Request.Form("mode")
	'�f���쐬����
	Case "create"
		'Set Rs = db.OpenSchema(20)	'�e�[�u�����擾

SQLs = "SELECT bbs_table, BBSName FROM settings WHERE bbs_table = 'bbs_" & BBSID & "'"

		Set RecSet = db.Execute(SQLs)
		If RecSet.EOF = False Then
			%>���̌f���͊��ɍ쐬����Ă��܂��B<br>
�f����ID�F<b><%=Replace(RecSet("bbs_table"), "bbs_", "") %></b><br>
�f�����F<b><%=RecSet("BBSName") %></b><br><%
		Else
		'�f���쐬����

		'�e�[�u���쐬

SQL = "SELECT * FROM settings WHERE bbs_table = 'default_settings'"
			Set Rs = db.Execute(SQL)

SQL = "CREATE TABLE bbs_" & BBSID & " (" & _
	"[Num] COUNTER NOT NULL, " & _
	"[abone] BIT, " & _
	"[from] TEXT(255), " & _
	"[mail] TEXT(255), " & _
	"[title] TEXT(255), " & _
	"[message] LONGTEXT, " & _
	"[url] TEXT(255), " & _
	"[sdat] DATETIME, " & _
	"[IP] TEXT(255), " & _
	"[Host] TEXT(255), " & _
	"[UserAgent] TEXT(255), " & _
	"[UA] TEXT(5), " & _
	"CONSTRAINT AutoInc PRIMARY KEY([Num]), " & _
	"UNIQUE([Num]))"
			db.Execute(SQL)

						'�ݒ�e�[�u���ɐV�K�̐ݒ����������
SQL = "INSERT INTO settings(" & _
	"[bbs_table], [SiteName], [SiteURL], [BaseURL], " & _
	"[BBSName], [BBSComment]," & _
	"[BGColor], [TextColor], [LinkColor], [aLinkColor], " & _
	"[BorderColor], [TitleColor], [ViewCount], [CountFile], " & _
	"[Tag], [TagSourceView], [MailSend], [Mailserver], " & _
	"[SendToAddr], [MailFromAddr], [MailBBSBodyCut], " & _
	"[NotFoundName], [DelMailAddr], [DelName], [DelTitle], " & _
	"[DelBody], [DelDevType], [DelTitleColor], [DelBodyColor]) "
SQL = SQL & "VALUES(" & _
	"'bbs_" & BBSID & "', '" & Rs("SiteName") & "', " & _
	"'" & Rs("SiteURL") & "', '" & Rs("BaseURL") & "', " & _
	"'" & Rs("BBSName") & "', '���ł������Ăˁ`��'," & _
	"'" & Rs("BGColor") & "', '" & Rs("TextColor") & "', " & _
	"'" & Rs("LinkColor") & "', '" & Rs("aLinkColor") & "', " & _
	"'" & Rs("BorderColor") & "', '" & Rs("TitleColor") & "', " & _
	Rs("ViewCount") & ", 'bbscnt_" & BBSID & ".dat', " & _
	"False, False, False, '" & Rs("MailServer") & "', " & _
	"'" & Rs("SendToAddr") & "', '" & Rs("MailFromAddr") & "', " & _
	"False, '" & Rs("NotFoundName") & "', " & _
	"'" & Rs("DelMailAddr") & "', '" & Rs("DelName") & "', " & _
	"'" & Rs("DelTitle") & "', '" & Rs("DelBody") & "', " & _
	"'" & Rs("DelDevType") & "', '" & Rs("DelTitleColor") & "', " & _
	"'" & Rs("DelBodyColor") & "')"

			db.Execute(SQL)

			Fso.CreateTextFile Server.MapPath("./count/bbscnt_" & BBSID & ".dat")
			Set Txt = Fso.OpenTextFile(Server.MapPath("./count/bbscnt_" & BBSID & ".dat"), 2)
			Txt.Write "0"
			Txt.Close
			Set Txt = Nothing

			Session("BBSQuery") = BBSID
%>
�f���F<%=BBSID %> ���쐬���܂����B<br>
<a href="<%=BBSURL %>bbs.asp?bbs=<%=BBSID %>" target="_blank"><%=BBSURL %>bbs.asp?bbs=<%=BBSID %></a><br>
<br>
<a href="<%=BBSURL %>admin.asp" target="_top">[�f���ݒ���]</a><br>
<%

		End If

	'�f���폜����
	Case "delete"

		If BBSID = "" Then
			%>�f�����I������Ă��܂���B<%
		Else
			If BBSID = "default" Then
				%><font size="+2">���̌f���͍폜�ł��܂���B</font><br>
default�f���͍폜���邱�Ƃ͂ł��܂���B<br>
�g�p�������Ȃ��ꍇ��<b>�{���E�������ݕs��</b>�ݒ�ɂ��Ă��������B<%
			Else

				SQLs = "SELECT bbs_table, BBSName FROM settings WHERE bbs_table = 'bbs_" & BBSID & "'"

				Set RecSet = db.Execute(SQLs)
				If RecSet.EOF = True Then
					%>���̌f���͂���܂���B<br>
		�f����ID�F<b><%=Replace(RecSet("bbs_table"), "bbs_", "") %></b><br><%

				Else


					'�f���폜����

					'�e�[�u���폜

					SQL = "DROP TABLE bbs_" & BBSID
					db.Execute(SQL)

					'�ݒ�e�[�u���̐ݒ���폜
					SQL = "DELETE FROM settings WHERE bbs_table = 'bbs_" & BBSID & "'"

					db.Execute(SQL)
					Session("BBSQuery") = "default"

					If Fso.FileExists(Server.MapPath("./count/bbscnt_" & BBSID & ".dat")) = True Then
						Fso.DeleteFile(Server.MapPath("./count/bbscnt_" & BBSID & ".dat"))
					End If

		%>ID: <b><%=BBSID %></b> �f�����폜���܂����B<br>
		<a href="admin.asp" target="_top">[BBS Administrator]</a><%

				End If
			End If

		End If

End Select

End If

Else
	'���O�C������Ă��Ȃ�
	Response.Redirect "admin.asp"
End If
%>
