<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="removehtml.asp" -->
<!-- #Include file="devtype.asp" -->
<%

'###########################################
'####					####
'####	     a n o n c o m B B S	####
'####					####
'###########################################
'				by anoncom.net


If CInt(BBSStatus) < 3 Then
	'�{���s�ݒ�̌f����forbidden.html�ɔ�΂�
	Response.Redirect "forbidden.html"
ElseIf CInt(BBSStatus) < 2 Then
	'�������f����closebbs.html�ɔ�΂�
	Response.Redirect "closebbs.html"
End If


cntpath = Server.MapPath("./count/" & CountFileName)
Set Fso = Server.CreateObject("Scripting.FileSystemObject")

If Fso.FileExists(cntpath) = True Then
	Set File = Fso.OpenTextFile(cntpath,1,True,False)
	Set objFile = Fso.GetFile(cntpath)
	If objFile.Size > 0 Then
		cnthit = File.ReadLine
		cnthit = cnthit + 1
		Set File = Fso.OpenTextFile(cntpath,2,False,False)
		File.Write cnthit
	Else
		Set File = Fso.OpenTextFile(cntpath,2,False,False)
		File.Write "0"
		cnthit = "�J�E���^�t�@�C�������Ă����̂ŏC�����܂����B<br>" & vbCrLf & _
			"0"
	End If
	File.Close
Else
	Fso.CreateTextFile Server.MapPath("./count/" & CountFileName)
	Set Txt = Fso.OpenTextFile(cntpath, 2)
	Txt.Write "0"
	Txt.Close
	Set Txt = Nothing
	cnthit = "�J�E���^�t�@�C����������܂���ł����B<br>" & _
	"�J�E���^�t�@�C�����쐬���܂����B<br>" & vbCrLf & "0"
End If

%>

<html>
<head>
<title><%=BBSName %> - anoncomBBS</title>
<% If MobileType = "" Then %>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<meta http-equiv="Content-Style-Type" content="text/css; charset=shift_jis">
<style type="text/css">
<!--
a:link { text-decoration:none; color:<%=LinkColor %>; }
a:visited { text-decoration:none; color:<%=LinkColor %>; }
a:active { text-decoration:none; color:<%=ActiveLinkColor %>; }
a:hover { text-decoration:underline; color:<%=HoverLinkColor %>; cursor:default; }

body{
	bakground-color:<%=BGColor %>;
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
	font-size:10pt;
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
}

td#tdname{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
	border-left-style:solid;
	border-left-width:10px;
	border-left-color:<%=BorderColor %>;
	border-bottom-style:solid;
	border-bottom-width:1px;
	border-bottom-color:<%=BorderColor %>;
	border-top-style:solid;
	border-top-width:1px;
	border-top-color:<%=BorderColor %>;
}

/*
td#bottom{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','�l�r�S�V�b�N';
	border-bottom-style:solid;
	border-bottom-width:1px;
	border-bottom-color:<%=BorderColor %>;
}
*/

input{
	color:<%=TextColor %>;
	background-color:<%=BGColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic';
	border-style:solid;
	border-width:1px;
	border-color:<%=BorderColor %>;
}
span.button{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic';
	border-style:solid;
	border-width:1px;
	border-color:<%=BorderColor %>;
	line-height:12pt;
	margin:1px;
	padding:1px;
}

-->
</style>
<% End If %>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<%
'PC���猩��ꍇ�̂݃e�[�u�����g�p���ĕ\��
Select Case BrowserType

   Case "Mobile"
etd = "<br>"
etd_bottom = "<br>" & vbCrLf & "<hr>"

   Case Else
 %>
<center>
<table border="0" width="90%" cellspacing="0">
<%
'�ϐ����g�p���ăe�[�u����`��
tr = "<tr>" & vbCrLf
etr = "</tr>"
td = "<td align=""left"" valign=""top"">" & vbCrLf
td_Name = "<td align=""left"" valign=""top"" id=""tdname"">" & vbCrLf
td_bottom = "<td align=""right"" valign=""top"" id=""bottom"">" & vbCrLf
td_r = "<td align=""right"" valign=""top"">" & vbCrLf
td_c = "<td align=""center"" valign=""top"">" & vbCrLf
etd = "</td>" & vbCrLf
etd_bottom = "</td>" & vbCrLf

End Select
%><%=tr %>
<%=td %><font size="5" color="<%=TitleColor %>"><i><%=BBSName %></i></font><%=etd %>
<%=etr & tr %>
<%= td_r%><font size="2">by anoncomBBS</font><%=etd %>
<%=etr & tr %>
<%=td & "&nbsp;" & etd %>
<%=etr & tr %>
<%=td_c %><%=Replace(BBSComment, vbCrLf, "<br>" & vbCrLf) %><%=etd %>
<%=etr & tr %>
<%=td & "&nbsp;" & etd %>
<%=etr & tr %>
<%

 Select Case BrowserType

'	�g�т̏ꍇ
   Case "Mobile"

	If CInt(BBSStatus) > 3 Then
%>
<a href="res.asp?bbs=<%=BBSQuery %>">����</a>
<hr>
<%
	End If

'	�p�\�R���A���̑�����̏ꍇ
   Case Else

	If BBSStatus > 3 Then
%>
<%=td %>
<span class="button"><a href="res.asp?bbs=<%=BBSQuery %>">��������</a></span>
<%=etd %><%
	Else %>
<%=td %>&nbsp;<%=etd %><%
	End If
%>
</tr><tr>
<td align="left" valign="top">&nbsp;</td>
</tr>
<td align="center">
<%

End Select


'PC���猩��ꍇ�̂݃e�[�u�����g�p���ĕ\��
Select Case BrowserType
   Case "Mobile"
etd = "<br>"
   Case Else
 %>
<table border="0" width="100%" id="bbs">
<%

End Select



'	�f���y�[�W�̃J�E���g
If Request.QueryString("cnt") = "" Then
  cnt = 1
  page = 1
Else
  cnt = CInt(Request.QueryString("cnt"))
  tmpCnt = cnt - CntNum
  page = CInt(Request.QueryString("page"))
End If


Set rs = Server.CreateObject("ADODB.Recordset")

''�e�[�u�������݂��邩�m�F
'Set TableExists = db.OpenSchema(12, TABLE_NAME)

'If TableExists("TABLE_NAME") = BBSQuery & "_bbs" Then
If BBSBlank = True Then
	rs.Open "SELECT * FROM blankbbs WHERE Num ORDER BY sdat DESC",db,3,2
Else
	rs.Open "SELECT * FROM bbs_" & BBSQuery & " WHERE Num ORDER BY sdat DESC",db,3,2
End If

Set TableExists = Nothing

If rs.EOF = True Then

'���X���Ȃ��ꍇ

ResCnt = "0"

%>
<%=tr %>
<%=td_name %>0:[anoncomBBS]<%=etd %>
<%=etr & tr %>
<%=td %>[No Script]<%=etd %>
<%=etr & tr %>
<%=td %>(�L���̏������݂�����܂���)<%=etd %>
<%=etr & tr %>
<%=td_bottom %>[system]<br>
[<%=WriteTime(Now) %>]<%=etd %>
<%=etr %>
<%
Else

'�������݂�����ꍇ
ResCnt = rs.RecordCount
rs.AbsolutePosition = cnt

Do While Not rs.EOF
  If rs("abone")="True" Then
'�폜�ς݃��X�̕\��
%><%=tr%>
<%=td_name %>
<%=rs("Num") %>:[<a href="mailto:<%=DeleteMailAddr %>"><%=DeleteName %></a>]
<%=etd %>
<%=etr & tr %>
<%=td %>
[<font color="<%=DelTitleColor %>"><%=DeleteTilte %></font>]
<%=etd %>
<%=etr & tr %>
<%=td %>
<font color="<%=DelBodyColor %>"><%=DeleteBody %></font>
<%=etd %>
<%=etr & tr %>
<%=td_bottom %>
[<%=DeleteDeviceType %>]<br>
[<%=WriteTime(rs("sdat")) %>]
<%=etd_bottom %>
<%=etr %><%
  Else
'�ʏ�̃��X�̕\��
    If rs("from")<>"" Then
      If rs("mail")<>"" Then
%><%=tr %>
<%=td_name %>
<%=rs("Num") %>:[<a href="mailto:<%=rs("mail") %>"><%=rs("from") %></a>]
<%=etd %>
<%      Else
%><%=tr %>
<%=td_name %>
<%=rs("Num") %>:[<%=rs("from") %>]
<%=etd %>
<%
      End If
    Else
      If rs("mail")<>"" Then
%><%=tr %>
<%=td_name %>
<%=rs("Num") %>:[<a href="mailto:<%=rs("mail") %>"><%=rs("mail") %></a>]
<%=etd %>
<%      Else
%><%=tr %>
<%=td_name %>
<%=rs("Num") %>:[<%=NotFoundName %>]
<%=etd %>
<%
      End If
    End If
    If rs("title")<>"" Then
%>
<%=etr & tr %>
<%=td %>
[<%=rs("title") %>]
<%=etd %><%

    End If


	If TagUse = 1 Then
		'�^�O�L��

	    If LenB(rs("message")) > 500 Then
	        message = "<br>" & LeftB(rs("message") ,500) & "...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & """>����</a>" & vbCrLf
	    Else
		message = "<br>" & rs("message") & vbCrLf
	    End If
		message = Replace(message,vbCrLf,"<br>" & vbCrLf)
	        message = message & "<br>" & vbCrLf

	Else
		'�^�O����

		'�^�O�\��
		If TagSourceView = 1 Then
			Set bsp = Server.CreateObject("basp21")	'BASP��ǂݍ���
			message = bsp.RepTagChar(rs("message"))
			Set bsp = Nothing
		Else
		'�^�O��\��
			'�^�O�����u������
			message = RemoveHTML(rs("message"))
		End If
		
	    If LenB(rs("message")) > 500 Then
	        message = Replace(LeftB(message,500),vbCrLf,"<br>" & vbCrLf) & _
			"...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & _
			""">����</a><br>" & vbCrLf
	    Else
	        message = Replace(message,vbCrLf,"<br>" & vbCrLf)
	    End If
	End If
%>
<%=etr & tr %>
<%=td %>
<%=message %>
<%=etd %><%

    If rs("url")<>"" Then
      If rs("url")<>"http://" Then
		HP_URL = "<a href=""" & rs("url")  & """ target=""_blank"">Homepage</a><br>" & vbCrLf
      Else
		HP_URL = ""
      End If
    Else
	HP_URL = ""
    End If
%>
<%=etr & tr %>
<%=td_bottom %>
<%=HP_URL %>[<%=rs("UA") %>]<br>
[<%=WriteTime(rs("sdat")) %>]
<%=etd_bottom %>
<%=etr %><%

  End If
  rs.MoveNext
  cnt = cnt + 1
  If cnt = page * CntNum + 1 Then Exit Do
Loop

'�I��(�e�[�u���^�O��)
Select Case Provider
   Case "DoCoMo"
   Case "J-PHONE"
   Case "au"
   Case "DDIPocket"
   Case Else

  Response.Write vbCrLf
%>
</table>
</td>
</tr><tr>
<%
End Select


'�������ݓ��t����
Function WriteTime(dtmNow)

Dim strDate
strDate = Right(String(4,"0") & Year(dtmNow),4) & "/" & Right(String(2,"0") & Month(dtmNow),2) & "/" & Right(String(2,"0") & Day(dtmNow),2) & " " & Right(String(2,"0") & Hour(dtmNow),2) & ":" & Right(String(2,"0") & Minute(dtmNow),2)
WriteTime = strDate

End Function

'�������ݓ��t�����I��

End If


 Select Case Provider

'	NTT DoCoMo i-mode �̏ꍇ
   Case "DoCoMo"
	If BBSStatus > 3 Then
%>
<a href="res.asp?bbs=<%=BBSQuery %>">����</a><br>
<%
	End If

If page>1 Then %>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=tmpCnt %>&page=<%=page - 1 %>">&lt;�߂�</a>/
<%
End If
If Not rs.EOF Then
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page + 1 %>">����&gt;</a>
<% End If %><br>
Find <%=cnthit %> ccess<br>
Write <%=ResCnt %>response<br><br>
<hr color="<%=BorderColor %>">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
<%
'	Vodafone Vodafone live! �̏ꍇ
   Case "Vodafone"
	If BBSStatus > 3 Then
%>
<a href="res.asp?bbs=<%=BBSQuery %>">����</a><br>
<%
	End If

If page > 1 Then %>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=tmpCnt %>&page=<%=page - 1 %>">$F[�߂�</a>/
<%
End If
If Not rs.EOF Then
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page + 1 %>">����$FZ</a>
<% End If %><br>
Find <%=cnthit %>access<br>
Write <%=ResCnt %>response<br><br>
<hr color="<%=BorderColor %>">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
<%
'	au Ez-web �̏ꍇ
   Case "au"

	If BBSStatus > 3 Then
%>
<a href="res.asp?bbs=<%=BBSQuery %>">����</a><br>
<%
	End If

If page>1 Then %>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=tmpCnt %>&page=<%=page - 1 %>">&lt;�߂�</a>/
<%
End If
If Not rs.EOF Then
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page + 1 %>">����&gt;</a>
<% End If %><br>
Find <%=cnthit %>access<br>
Write <%=ResCnt %>response<br><br>
<hr color="<%=BorderColor %>">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
<%
'	�p�\�R���A���̑�����̏ꍇ
   Case Else

	If BBSStatus > 3 Then
%>
<td align="left" valign="top">
<span class="button"><a href="res.asp?bbs=<%=BBSQuery %>">��������</a></span>
</td>
</tr><tr>
<%
	End If

If Not rs.EOF Then
%>
<td align="left" valign="top">
<span class="button"><a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page+1 %>">����</a></span>
</td>
</tr><tr>
<% End If %>
<td align="left">&nbsp;</td>
</tr><tr>
<td align="left" valign="top" id="bottom">
<font size="2"><i>
<%=cnthit %>access<br>
<%=ResCnt %>response<br>
</i></font>
</td>
</tr><tr>
<td align="right">
<i>system by <a href="http://anoncom.net/" target="_top">anoncom.net</a></i>
</td>
</tr>
</table>
<%
End Select
'	************UA�ɂ��\���`���̕ύX�I��***************


%>
</body>
</html>
