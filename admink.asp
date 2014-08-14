<%@ Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<%

If Request("id") = ""  And Request("pw") = "" Then
%><html>
<head>
<title>BBS Administrator</title>
</head>
<body link="#ff0000" vlink="#ff0000" alink="#ff0000">
<font color="#0000ff">BBS Administrator</font>
<form action="admink.asp" method="get">
管理者ID：<input type="text" name="id" size="8"><br>
Password：<input type="password" name="pw" size="8"><br>
BBS ID：
<select name="bbs">
<option value="">(掲示板...)</option><%

SQL = "SELECT bbs_table, BBSName FROM settings WHERE bbs_table Like 'bbs_%'"

Set rs = db.Execute(SQL)

If rs.EOF = False Then

	Do While rs.EOF = False
		BBS_ID = Replace(rs("bbs_table"), "bbs_", "")
		BBS_Name = rs("BBSName")
		%>
<option value="<%=BBS_ID %>"><%=BBS_Name %></option><%
		rs.MoveNext
	Loop
Else
%>
<option value="">掲示板がありません</option><%
End If

Set rs = Nothing


%>
</select><br>
<input type="submit" value="ログイン">
</form>
<br><br>
<hr size="1">
system by <a href="http://anoncom.net/">anoncom.net</a>
</body>
</html><%

Else

	'ログイン処理
	'管理情報読み込み
	Set rs_admin = Server.CreateObject("ADODB.Recordset")
	rs_admin.Open "SELECT * FROM admin_settings " & _
		"WHERE adminID = '" & Request("id") & "' AND " & _
		"adminPass = '" & Request("pw") & "'", db, 3, 2

	'認証完了
	If  rs_admin.EOF = False Then

		'初期設定
		AdminID = rs_admin("adminID")
		AdminPass = rs_admin("adminPass")
		AdminName = Server.URLEncode(rs_admin("adminName"))
		AdminMail = Server.URLEncode(rs_admin("adminMail"))
		AdminLevel = CInt(rs_admin("adminRank"))
		AdminBBS = rs_admin("adminBBS")
		rs_admin.Close
		AdminBBSSession = False

		'BBSQuery = Request("bbs")
	Else
		Response.Redirect "admink.asp"
	End If

	Set rs_admin = Nothing

	'Response.Redirect "admink.asp?id=" & AdminID & "&pw=" & AdminPass


	If BBSBlank = True Then
		Response.Redirect "nobbs.html"
	End If

	If AdminBBS = "allbbs" Then
		AdminBBSSession = True
	Else
		aryAdminBBS = Split(AdminBBS, ",")
		For i = 0 To UBound(aryAdminBBS)
			If BBSQuery = aryAdminBBS(i) Then
				AdminBBSSession = True
				Exit For
			End If
		Next
	End If

If AdminBBSSession = True Then

    Select Case Request("type")


'==============================================================================
'	設定画面
	Case "setting"


If AdminLevel >= 4 Then

%>
<html>
<head>
<title>掲示板設定</title>
</head>
<body link="#ff0000" vlink="#ff0000">
<font color="#0000ff" size="+2">掲示板設定 for <%=BBSName %></font><hr><%
		If Request.Form("edit") <> "on" Then %>

<form action="admink.asp" method="post">
<input type="hidden" name="id" value="<%=AdminID %>">
<input type="hidden" name="pw" value="<%=AdminPass %>">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="type" value="setting">
<input type="hidden" name="edit" value="on">
掲示板名:<br>
<input type="text" name="BBSName" value="<%=BBSName %>" maxlength="100" size="20"><br>
掲示板コメント:<br>
<textarea name="BBSComment" rows="2" cols="20"><%=BBSComment %></textarea>
<br>
掲示板のURL:<br>
<input type="text" name="baseurl" value="<%=BBSURL %>" maxlength="255" size="20"><br>
掲示板の状態:
<%
Select Case BBSStatus
	Case 9: status9 = " selected"
	Case 3: status3 = " selected"
	Case 0: status0 = " selected"
End Select %><select name="bbsstatus">
<option value="9"<%=status9 %>>ReadWrite</option>
<option value="3"<%=status3 %>>Read Only</option>
<option value="0"<%=status0 %>>Forbidden</option>
</select><br>
デバッグモード:
<%
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
<option value="0"<%=debug0 %>>OFF</option>
<option value="3"<%=debug3 %>>簡易Debug</option>
<option value="5"<%=debug5 %>>通常Debug</option>
<option value="9"<%=debug9 %>>完全Debug</option>
</select><br>
背景色:
<input type="text" name="bgcolor" value="<%=BGColor %>" maxlength="20" size="10"><br>
文字の色:
<input type="text" name="textcolor" value="<%=TextColor %>" maxlength="20" size="10"><br>
リンクの色:
<input type="text" name="linkcolor" value="<%=LinkColor %>" maxlength="20" size="10"><br>
掲示板名の色:
<input type="text" name="titlecolor" value="<%=TitleColor %>" maxlength="20" size="10"><br>
表示件数:
<input type="text" name="cntnum" value="<%=CntNum %>" maxlength="3" size="3"><br>
カウンタファイル名:
<input type="text" name="countfilename" value="<%=CountFileName %>" size="10"><br>
タグの使用:
<%
Select Case TagUse
	Case 1: TUon = " selected"
	Case 0: TUoff = " selected"
End Select
%><select name="tag">
<option value="1"<%=TUon %>>有効</option>
<option value="0"<%=TUoff %>>無効</option>
</select><br>
タグ表示:
<%
Select Case TagSourceView
	Case 1: TSVon = " selected"
	Case 0: TSVoff = " selected"
End Select
%><select name="tagsourceview">
<option value="1"<%=TSVon %>>表示</option>
<option value="0"<%=TSVoff %>>非表示</option>
</select><br>
通知配信:
<%
Select Case BBSMailSend
	Case 1: BMSon = " selected"
	Case 0: BMSoff = " selected"
End Select
%><select name="bbsmailsend">
<option value="1"<%=BMSon %>>有効</option>
<option value="0"<%=BMSoff %>>無効</option>
</select><br>
メールサーバー:<br>
<input type="text" name="mailserver" value="<%=MailServer %>" maxlength="255" size="20"><br>
<!--送信先アドレス:<br>
<input type="text" name="sendtoaddr" value="<%=SendToAddr %>" maxlength="255" size="20"><br>-->
送信先グループ:<br>
<input type="text" name="SendGroup" value="<%=UserGroup %>" maxlength="255" size="20"><br>
送信元アドレス:<br>
<input type="text" name="mailfromaddr" value="<%=MailFromAddr %>" maxlength="255" size="20"><br>
通知本文カット:
<%
	Select Case MailBodyCut
		Case 1: MBCon = " selected"
		Case 0: MBCoff = " selected"
	End Select
%><select name="mailbbsbodycut">
<option value="1"<%=MBCon %>>有効</option>
<option value="0"<%=MBCoff %>>無効</option>
</select><br>
<br>
<hr>
<input type="submit" value=" 変更 ">
</form>
<%
		Else

			Set rs_set = Server.CreateObject("ADODB.Recordset")
			upSQL = "SELECT * FROM settings WHERE bbs_table = 'bbs_" & BBSQuery & "'"
			rs_set.Open upSQL,db,3,3

			'設定の反映
			rs_set("BBSName") = Request.Form("BBSName")
			rs_set("BBSComment") = Request.Form("BBSComment")
			rs_set("BaseURL") = Request.Form("BaseURL")
			rs_set("act_flag") = CInt(Request.Form("BBSStatus"))
			rs_set("debug_flag") = CInt(Request.Form("DebugMode"))
			rs_set("BGColor") = Request.Form("BGColor")
			rs_set("TextColor") = Request.Form("TextColor")
			rs_set("LinkColor") = Request.Form("LinkColor")
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

			'項目の更新
			rs_set.Update
			'データベースとの接続を閉じる
			rs_set.Close
			'メモリの開放
			Set rs_set = Nothing
			%>

設定が変更されました。<br>
<a href="admink.asp?id=<%=AdminID %>&pw=<%=AdminPass %>&bbs=<%=BBSQuery %>">[管理画面]</a><%
		End If %>
</body>
</html>
<%

Else
%>権限がありません。<%
End If

'==============================================================================
'	書き込み解析画面
	Case "access"

If AdminLevel >= 1 Then
%>
<html>
<head>
<title>書き込み解析</title>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
</head>
<body link="#ff0000" vlink="#ff0000" alink="#ff0000">
<font color="#0000ff" size="+2">書き込み解析 for <%=BBSName %></font><hr>
<%
If Request.QueryString("cnt")="" Then
  cnt = 1
  page = 1
Else
  cnt = CInt(Request.QueryString("cnt"))
  tmpCnt = cnt - 5
  page = CInt(Request.QueryString("page"))
End If
Set db = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Mode = 1
db.ConnectionString = BBSDBFileName
db.Open
rs.Open "SELECT * FROM bbs_" & BBSQuery & " WHERE Num ORDER BY sdat DESC",db,3,2

If rs.EOF = True Then

'  IP&Host名称取得＆表示
Rmt_Addr  =  IP
Rmt_Host  =  Perlget(Rmt_Addr)         ' 名称取得


'  IP&Host名称取得	by CGI Perl Script
%>
<script language="PerlScript" runat="Server">
sub Perlget{
  local($addr) = @_;
    $host = gethostbyaddr(pack("C4", split(/\./, $addr)), 2);

    if ($host eq "") { $host = $addr; }
    return $host;
}
</script>
0:■[anoncomBBS]<br>
[No Script]<br>
<br>(記事の書き込みがありません)<br>
<br>
[abone: (出来ません)]<br>
[system]<br>
[anoncomBBS ver.1.3]<br>
[IP: <%=Request.ServerVariables("REMOTE_ADDR") %>]<br>
[Host: <%=Rmt_Host %>]<br>
[通知配信: False]<br>
[<%=Now %>]<br>
<hr>
<%
Else

rs.AbsolutePosition=cnt
Do While Not rs.EOF
    If rs("from")<>"" Then
      If rs("mail")<>"" Then
        Response.Write rs("Num") & ":■[<a href=""mailto:" & rs("mail") &""">" & rs("from") & "</a>]<br>" & vbCrLf
      Else
        Response.Write rs("Num") & ":■[" & rs("from") & "]<br>" & vbCrLf
      End If
    Else
      If rs("mail")<>"" Then
        Response.Write rs("Num") & ":■[<a href=""mailto:" & rs("mail") &""">" & rs("mail") & "</a>]<br>" & vbCrLf
      Else
        Response.Write rs("Num") & ":■[" & NotFoundName & "]<br>" & vbCrLf
      End If
    End If
    If rs("title")<>"" Then
      Response.Write "[" & rs("title") & "]<br>" & vbCrLf
    End If

    If Len(rs("message")) > 80 Then
        message = "<br>" & Left(rs("message") ,80) & "...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & "&del=view"">続き</a>" & vbCrLf
    Else
	message = "<br>" & rs("message") & vbCrLf
    End If
	message = Replace(message,vbCrLf,"<br>" & vbCrLf)
        message = message & "<br><br>" & vbCrLf
    Response.Write message
    If rs("url")<>"" Then
      If rs("url")<>"http://" Then
        Response.Write "<a href=""jump.asp?url=" & rs("url") & """ target=""_blank"">" & rs("url") & "</a><br>" & vbCrLf
      End If
    End If
    Response.Write "[abone: " & rs("abone") & "]<br>" & vbCrLf
    Response.Write "[" & rs("UA") & "]<br>" & vbCrLf
    Response.Write "[" & rs("UserAgent") & "]<br>" & vbCrLf
    Response.Write "[IP: " & rs("IP") & "]<br>" & vbCrLf
    Response.Write "[Host: " & rs("Host") & "]<br>" & vbCrLf
    Response.Write "[通知配信: " & rs("sendchk") & "]<br>" & vbCrLf
    Response.Write "[" & FormatDateTime(rs("sdat"),0) & "]<hr>" & vbCrLf
  rs.MoveNext
  cnt=cnt+1
  If cnt=page*5+1 Then Exit Do
Loop

If Not rs.EOF Then
%>
<form action="admink.asp" method="get">
<input type="hidden" name="id" value="<%=AdminID %>">
<input type="hidden" name="pw" value="<%=AdminPass %>">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="type" value="access">
<input type="hidden" name="cnt" value="<%=cnt %>">
<input type="hidden" name="page" value="<%=page+1 %>">
<input type="submit" value="次へ">
<%
End If

End If
%>
</body>
</html>
<%
Else
%>権限がありません。<%
End If

'==============================================================================
'	削除管理画面
	Case "del"

If AdminLevel >= 2 Then

  If Request("edit") = "" Then
%>
<html>
<head>
<title>削除管理</title>
</head>
<body link="#ff0000" vlink="#ff0000" alink="#ff0000">
<font color="#ff0000" size="5">削除管理 for <%=BBSName %></font>
<hr>
<%
If Request.QueryString("cnt")="" Then
  cnt=1
  page=1
Else
  cnt=CInt(Request.QueryString("cnt"))
  tmpCnt=cnt-5
  page=CInt(Request.QueryString("page"))
End If
Set db=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.Recordset")
db.Provider="Microsoft.Jet.OLEDB.4.0"
db.Mode=1
db.ConnectionString=BBSDBFileName
db.Open
rs.Open "SELECT * FROM bbs_" & BBSQuery & " WHERE Num ORDER BY sdat DESC",db,3,2


If rs.EOF = True Then

'  IP&Host名称取得＆表示
Rmt_Addr  =  IP
Rmt_Host  =  Perlget(Rmt_Addr)         ' 名称取得


'  IP&Host名称取得	by CGI Perl Script
%>
<script language="PerlScript" runat="Server">
sub Perlget{
  local($addr) = @_;
    $host = gethostbyaddr(pack("C4", split(/\./, $addr)), 2);

    if ($host eq "") { $host = $addr; }
    return $host;
}
</script>
0:■[anoncomBBS]<br>
[No Script]<br>
<br>(記事の書き込みがありません)<br>
<br>
[system]<br>
[anoncomBBS ver.1.3]<br>
RemoteHost: <%=Rmt_Host %><br>
[<%=Now %>]<br>
<input type="submit" value="削除" disabled><hr>
<%
Else

rs.AbsolutePosition=cnt
Do While Not rs.EOF

%><form action="admink.asp" method="get">
<%
    If rs("from")<>"" Then
      If rs("mail")<>"" Then
        Response.Write rs("Num") & ":■[<a href=""mailto:" & rs("mail") &""">" & rs("from") & "</a>]<br>" & vbCrLf
      Else
        Response.Write rs("Num") & ":■[" & rs("from") & "]<br>" & vbCrLf
      End If
    Else
      If rs("mail")<>"" Then
        Response.Write rs("Num") & ":■[<a href=""mailto:" & rs("mail") &""">" & rs("mail") & "</a>]<br>" & vbCrLf
      Else
        Response.Write rs("Num") & ":■[" & NotFoundName & "]<br>" & vbCrLf
      End If
    End If
    If rs("title")<>"" Then
      Response.Write "[" & rs("title") & "]<br>" & vbCrLf
    End If

    If Len(rs("message")) > 80 Then
        message = "<br>" & Left(rs("message") ,80) & "...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & "&del=view"">続き</a>" & vbCrLf
    Else
	message = "<br>" & rs("message") & vbCrLf
    End If
	message = Replace(message,vbCrLf,"<br>" & vbCrLf)
        message = message & "<br><br>" & vbCrLf
    Response.Write message
    If rs("url")<>"" Then
      If rs("url")<>"http://" Then
        Response.Write "<a href=""jump.asp?url=" & rs("url") & """ target=""_blank"">" & rs("url") & "</a><br>" & vbCrLf
      End If
    End If

    If rs("abone")="True" Then
      Abone="no"
      AboneValue="解除"
    ElseIf rs("abone")="False" Then
      Abone="yes"
      AboneValue="削除"
    End If

	If Len(rs("UserAgent")) > 35 Then
		usragent = Left(rs("UserAgent"),35) & "..."
	Else
		usragent = rs("UserAgent")
	End If

%>
[<%=usragent %>]<br>
[<%=FormatDateTime(rs("sdat"),0) %>]<br>
<input type="hidden" name="id" value="<%=AdminID %>">
<input type="hidden" name="pw" value="<%=AdminPass %>">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="type" value="del">
<input type="hidden" name="edit" value="on">
<input type="hidden" name="no" value="<%=rs("Num") %>">
<input type="hidden" name="abone" value="<%=Abone %>">
<input type="hidden" name="cnt" value="<%=Request.QueryString("cnt") %>">
<input type="hidden" name="page" value="<%=Request.QueryString("page") %>">
<input type="submit" value="<%=AboneValue %>">
</form>
<hr>
<%
  rs.MoveNext
  cnt = cnt + 1
  If cnt = page * 5 + 1 Then Exit Do
Loop
If Not rs.EOF Then
%>
<form action="admink.asp" method="get">
<input type="hidden" name="id" value="<%=AdminID %>">
<input type="hidden" name="pw" value="<%=AdminPass %>">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<input type="hidden" name="type" value="del">
<input type="hidden" name="cnt" value="<%=cnt %>">
<input type="hidden" name="page" value="<%=page+1 %>">
<input type="submit" value="次へ">
<%
End If
rs.Close

End If

db.Close
%>
</body>
</html>
<%
  Else						'edit=onの時


	anoncomBBS = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & BBSDBFileName
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open anoncomBBS

      If Request.QueryString("abone") = "yes" Then
          db_abone = "True"
      ElseIf Request.QueryString("abone") = "no" Then
          db_abone = "False"
      End If

  SQL="UPDATE bbs_" & BBSQuery & " SET abone=" & db_abone
  SQL=SQL & " WHERE Num=" & Request.QueryString("no")
	conn.Execute(SQL)
	conn.Close

      Response.Redirect "admink.asp?id=" & AdminID & "&pw=" & AdminPass & "&bbs=" & BBSQuery & "&type=del&cnt=" & Request("cnt") & "&page=" & Request("page")

      End If

Else
%>権限がありません。<%
End If

'==============================================================================
'管理画面メニュー

	Case Else
%>
<html>
<head>
<title>BBS Administrator</title>
</head>
<body link="#ff0000" vlink="#ff0000" alink="#ff0000">
<font color="#0000ff">BBS Administrator for <%=BBSName %></font>
<br><br><br><%
If AdminLevel >= 1 Then %>
<a href="admink.asp?bbs=<%=BBSQuery %>&id=<%=AdminID %>&pw=<%=AdminPass %>&type=access">書き込み解析</a><br><%
End If

If AdminLevel >= 2 Then %>
<a href="admink.asp?bbs=<%=BBSQuery %>&id=<%=AdminID %>&pw=<%=AdminPass %>&type=del">削除管理</a><br><%
End If

If AdminLevel >= 4 Then %>
<a href="admink.asp?bbs=<%=BBSQuery %>&id=<%=AdminID %>&pw=<%=AdminPass %>&type=setting">掲示板設定</a><br><%
End If %>
<br><br>
<hr size="1">
system by <a href="http://anoncom.net/">anoncom.net</a>
</body>
</html>
<%

End Select

Else
%>この掲示板の管理はできません。<%
End If

End If
%>
