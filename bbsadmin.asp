<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'�g�т̏ꍇ�͌g�їp�̊Ǘ��y�[�W�ֈړ�
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

If CBool(BBSBlank) = True Then
	Response.Redirect "nobbs.html"
End If


'���O�C���Z�b�V�����͊m�����Ă��邩
If Session("login") = 1 Then

Select Case Request.QueryString("mode")

	'�A�N�Z�X���
	Case"access"
%>
<html>
<head>
<title>�������݉��</title>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/css; charset=shift_jis">
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

table#bbs{
	border-style:solid;
	border-color:<%=BorderColor %>;
	border-width:1px;
}

td{
	color:<%=TextColor %>;
	font-size:12pt;
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
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" vlink="<%=LinkColor %>" alink="<%=ActiveLinkColor %>">
<font color="#0000ff" size="+2">�������݉�� for <%=BBSName %></font><hr>
<%
	If Request("cnt") = "" Then
	  cnt = 1
	  page = 1
	Else
	  cnt = CInt(Request("cnt"))
	  tmpCnt = cnt - 5
	  page = CInt(Request("page"))
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM bbs_" & BBSQuery & " WHERE Num ORDER BY sdat DESC",db,3,2

	If rs.EOF = True Then

		'  IP&Host���̎擾���\��
		Rmt_Addr  =  IP
		Rmt_Host  =  Perlget(Rmt_Addr)         ' ���̎擾


		'  IP&Host���̎擾	by CGI Perl Script
		%>
		<script language="PerlScript" runat="Server">
		sub Perlget{
		  local($addr) = @_;
		    $host = gethostbyaddr(pack("C4", split(/\./, $addr)), 2);

		    if ($host eq "") { $host = $addr; }
		    return $host;
		}
		</script>
0:��[anoncomBBS]<br>
[No Script]<br>
<br>(�L���̏������݂�����܂���)<br>
<br>
[abone: (�o���܂���)]<br>
[system]<br>
[anoncomBBS ver.1.3]<br>
[IP: <%=Request.ServerVariables("REMOTE_ADDR") %>]<br>
[Host: <%=Rmt_Host %>]<br>
[<%=Now %>]<br>
<hr>
<%
	Else

		rs.AbsolutePosition=cnt
		Do While Not rs.EOF
		    If rs("from")<>"" Then
		      If rs("mail")<>"" Then
		        Response.Write rs("Num") & ":��[<a href=""mailto:" & rs("mail") &""">" & rs("from") & "</a>]<br>" & vbCrLf
		      Else
		        Response.Write rs("Num") & ":��[" & rs("from") & "]<br>" & vbCrLf
		      End If
		    Else
		      If rs("mail")<>"" Then
		        Response.Write rs("Num") & ":��[<a href=""mailto:" & rs("mail") &""">" & rs("mail") & "</a>]<br>" & vbCrLf
		      Else
		        Response.Write rs("Num") & ":��[" & NotFoundName & "]<br>" & vbCrLf
		      End If
		    End If
		    If rs("title")<>"" Then
		      Response.Write "[" & rs("title") & "]<br>" & vbCrLf
		    End If

		    If Len(rs("message")) > 80 Then
		        message = "<br>" & Left(rs("message") ,80) & "...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & "&del=view"">����</a>" & vbCrLf
		    Else
			message = "<br>" & rs("message") & vbCrLf
		    End If
			message = Replace(message,vbCrLf,"<br>" & vbCrLf)
		        message = message & "<br><br>" & vbCrLf
		    Response.Write message
		    If rs("url")<>"" Then
		      If rs("url")<>"http://" Then
		        Response.Write "<a href=""http://anoncom.net/jump.asp?url=" & rs("url") & """ target=""_blank"">" & rs("url") & "</a><br>" & vbCrLf
		      End If
		    End If
		    Response.Write "[abone: " & rs("abone") & "]<br>" & vbCrLf
		    Response.Write "[" & rs("UA") & "]<br>" & vbCrLf
		    Response.Write "[" & rs("UserAgent") & "]<br>" & vbCrLf
		    Response.Write "[IP: " & rs("IP") & "]<br>" & vbCrLf
		    Response.Write "[Host: " & rs("Host") & "]<br>" & vbCrLf
		    Response.Write "[" & FormatDateTime(rs("sdat"),0) & "]<hr>" & vbCrLf
		  rs.MoveNext
		  cnt=cnt+1
		  If cnt=page*5+1 Then Exit Do
		Loop

		If Not rs.EOF Then
%>
<form action="bbsadmin.asp?mode=access" method="post">
<input type="hidden" name="id" value="<%=AdminID %>">
<input type="hidden" name="pw" value="<%=AdminPass %>">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="type" value="access">
<input type="hidden" name="cnt" value="<%=cnt %>">
<input type="hidden" name="page" value="<%=page+1 %>">
<input type="submit" value="����">
<%
		End If

	End If
%>
</body>
</html>
<%
	'�폜�Ǘ����
	Case"delete"
	If Request.Form("edit")="" Then
%>
<html lang="ja">
<head>
<title>���X�폜�Ǘ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/css; charset=shift_jis">
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

table#bbs{
	border-style:solid;
	border-color:<%=BorderColor %>;
	border-width:1px;
}

td{
	color:<%=TextColor %>;
	font-size:12pt;
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
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" vlink="<%=LinkColor %>" alink="<%=ActiveLinkColor %>">
<font color="#ff0000" size="5">�폜�Ǘ� for <%=BBSName %></font>
<hr>
<%
	If Request("cnt") = "" Then
	  cnt = 1
	  page = 1
	Else
	  cnt = CInt(Request("cnt"))
	  tmpCnt = cnt - 5
	  page = CInt(Request("page"))
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM bbs_" & BBSQuery & " WHERE Num ORDER BY sdat DESC",db,3,2


	If rs.EOF = True Then

		'  IP&Host���̎擾���\��
		Rmt_Addr  =  IP
		Rmt_Host  =  Perlget(Rmt_Addr)         ' ���̎擾


		'  IP&Host���̎擾	by CGI Perl Script
		%>
		<script language="PerlScript" runat="Server">
		sub Perlget{
		  local($addr) = @_;
		    $host = gethostbyaddr(pack("C4", split(/\./, $addr)), 2);

		    if ($host eq "") { $host = $addr; }
		    return $host;
		}
		</script>
0:��[anoncomBBS]<br>
[No Script]<br>
<br>(�L���̏������݂�����܂���)<br>
<br>
[system]<br>
[anoncomBBS ver.1.3]<br>
RemoteHost: <%=Rmt_Host %><br>
[<%=Now %>]<br>
<input type="submit" value="�폜" disabled><hr>
<%
	Else

		rs.AbsolutePosition=cnt
		Do While Not rs.EOF

%><form action="bbsadmin.asp?mode=delete" method="post">
<%
		    If rs("from")<>"" Then
		      If rs("mail")<>"" Then
		        Response.Write rs("Num") & ":��[<a href=""mailto:" & rs("mail") &""">" & rs("from") & "</a>]<br>" & vbCrLf
		      Else
		        Response.Write rs("Num") & ":��[" & rs("from") & "]<br>" & vbCrLf
		      End If
		    Else
		      If rs("mail")<>"" Then
		        Response.Write rs("Num") & ":��[<a href=""mailto:" & rs("mail") &""">" & rs("mail") & "</a>]<br>" & vbCrLf
		      Else
		        Response.Write rs("Num") & ":��[" & NotFoundName & "]<br>" & vbCrLf
		      End If
		    End If
		    If rs("title")<>"" Then
		      Response.Write "[" & rs("title") & "]<br>" & vbCrLf
		    End If

		    If Len(rs("message")) > 80 Then
		        message = "<br>" & Left(rs("message") ,80) & "...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & "&del=view"">����</a>" & vbCrLf
		    Else
			message = "<br>" & rs("message") & vbCrLf
		    End If
		    message = Replace(message,vbCrLf,"<br>" & vbCrLf)
		    message = message & "<br><br>" & vbCrLf
		    Response.Write message
		    If rs("url")<>"" Then
		      If rs("url")<>"http://" Then
		        Response.Write "<a href=""" & SiteURL & "jump.asp?url=" & rs("url") & """ target=""_blank"">" & rs("url") & "</a><br>" & vbCrLf
		      End If
		    End If

		    If rs("abone")="True" Then
		      Abone="no"
		      AboneValue="����"
		    ElseIf rs("abone")="False" Then
		      Abone="yes"
		      AboneValue="�폜"
		    End If

		    If Len(rs("UserAgent")) > 35 Then
			usragent = Left(rs("UserAgent"),35) & "..."
		    Else
			usragent = rs("UserAgent")
		    End If

%>
[<%=usragent %>]<br>
[<%=FormatDateTime(rs("sdat"),0) %>]<br>
<input type="hidden" name="edit" value="on">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="no" value="<%=rs("Num") %>">
<input type="hidden" name="abone" value="<%=Abone %>">
<input type="hidden" name="cnt" value="<%=Request("cnt") %>">
<input type="hidden" name="page" value="<%=Request("page") %>">
<input type="submit" value="<%=AboneValue %>">
</form>
<hr>
<%
		  rs.MoveNext
		  cnt=cnt+1
		  If cnt=page*5+1 Then Exit Do
		  Loop
		  If Not rs.EOF Then
%>
<form action="bbsadmin.asp?mode=delete" method="post">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="cnt" value="<%=cnt %>">
<input type="hidden" name="page" value="<%=page+1 %>">
<input type="submit" value="����">
<%
		   End If
		   rs.Close

	End If

	db.Close
%>
</body>
</html>
<%
	Else				'edit=on�̎�


		anoncomBBS = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & BBSDBFileName
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open anoncomBBS

		If Request.Form("abone")="yes" Then
			db_abone = "True"
		ElseIf Request.Form("abone")="no" Then
			db_abone = "False"
		End If

		SQL = "UPDATE bbs_" & BBSQuery & " SET abone=" & db_abone
		SQL = SQL & " WHERE Num=" & Request.Form("no")
		conn.Execute(SQL)
		conn.Close

		Response.Redirect "bbsadmin.asp?bbs=" & BBSQuery & "&mode=delete&cnt=" & Request.Form("cnt") & "&page=" & Request.Form("page")

	End If

	'�f�����X��������
	Case "clear"
%><html>
<head>
<title>�������ݑS�폜</title>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/css; charset=shift_jis">
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
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" vlink="<%=LinkColor %>" alink="<%=ActiveLinkColor %>">
<font face="Times New Roman" size="+3"><b><i>BBS Response Clear</i></i></font><br>
<script language="JavaScript"><!--
function BBSClear(){
  msgRet = confirm("�{���ɂ�낵���ł����H");
  if ( msgRet == true ){
	document.delform.submit();
  }
}
// --></script>
<center>
<table border="0" height="40%" width="70%">
<tr>
<td align="center" bgcolor="#ddffdd">
<table border="0">
<tr>
<td align="center" bgcolor="#ddffdd" colspan="2">
<b><%=BBSName %>�i<%=BBSQuery %>�j�̏������ݓ��e�����ׂď��������܂��B��낵���ł����H</b><br>
<br>
<u>�f���̐ݒ�͏���������܂���B</u><br>
<span style="backgroud-color:#ffff00;">�����̍�Ƃ͎��������Ƃ��o���܂���B</span>
</td>
</tr><tr>
<td align="center" height="50">&nbsp;</td>
</tr><tr>
<td align="center" valign="bottom">
<form action="bbsclear.asp" method="post">
<input type="checkbox" name="cntreset">�J�E���^��������<br>
<input type="hidden" name="setup" value="execute">
<input type="submit" value="�͂�" onclick="BBSClear()">
</td>
<td align="center" valign="bottom">
<input type="button" value="������" onClick="javascript:history.back(-1)">
</td>
</td>
</tr>
</table>
</tr>
</table>
</center>
</body>
</html>
<%
	'�o�b�N�A�b�v����
	Case "backup"


	'BBS Redirect
	Case "bbs"
		Response.Redirect "bbs.asp?bbs=" & BBSQuery

	'�f���ݒ�
	Case "setting"
		Response.Redirect "setting.asp?bbs=" & BBSQuery & "&set=bbs"

	'�ݒ�̏�����
	Case "settingclear"
		Response.Redirect "setup.asp?bbs=" & BBSQuery

	'�Ǘ��Ґݒ�
	Case "adminsetting"
		Response.Redirect "setting.asp?bbs=" & BBSQuery & "&set=admin"

	'�V�K�Ǘ��Ғǉ�
	Case "adminuserreg"
		Response.Redirect "admusrmng.asp"

	'�f�[�^�x�[�X�ړ�
	Case "dbmove"
		Response.Redirect "dbmove.asp"

	'�N�G�����Ȃ��ꍇ
	Case Else
		Response.Redirect "admin.asp?main"

End Select

Else

'���O�C���Z�b�V�����؂�̏ꍇ
	Response.Redirect "admin.asp"

End If
%>