<% @Language = "VBScript" %>
<!-- #Include file="removehtml.asp" -->
<!-- #Include file="config.asp" -->
<%
'***********	レスフォーム	************


If CInt(BBSStatus) < 3 Then
	Response.Redirect "forbidden.html"
ElseIf CInt(BBSStatus) < 2 Then
	Response.Redirect "closebbs.html"
End If

If CInt(BBSStatus) < 4 Then
%>
<html lang="ja">
<head>
<meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/html; charset=shift_jis">
<title>Response Forbidden</title>
<style>
<!--
a:link { text-decoration:none; color:#ff0000; }
a:visited { text-decoration:none; color:#ff0000; }
a:active { text-decoration:none; color:#0000ff; }
a:hover { text-decoration:underline; color:#0000ff; cursor:default; }

body{
	background-color:#ffffff;
	color:#000000;
	font-size:10pt;
	font-family:'MS UI Gothic','ＭＳゴシック';
	overflow-y:auto;
	scrollbar-base-color:#ffffff;
	scrollbar-face-color:#ffffff;
	scrollbar-arrow-color:#00000;
	scrollbar-highlight-color:#ffffff;
	scrollbar-3dlight-color:#00000;
	scrollbar-shadow-color:#00000;
	scrollbar-darkshadow-color:#ffffff;
}
-->
</style>
</head>
<body>
<font color="#ff0000" face="Times New Roman" size="+3"><b><i>Response Forbidden</i></b></font><br>
<br>
<br>
この掲示板は、現在書込み禁止設定になっています。<br>
閲覧のみでお楽しみください。
<br>
<br>
<br>
<a href="bbs.asp?bbs=<%=Request.QueryString("bbs") %>">[掲示板]</a><br>
<br>
<br>
</body>
</html>
<%
Else


If BBSBlank = True Then
	Response.Redirect BBSURL & "bbs.asp?bbs=" & BBSQuery
End If

If Request.Form="" Then


'カウンタ(表示のみ)
cntpath = Server.MapPath("./count/" & CountFileName)
Set FSObj = Server.CreateObject("Scripting.FileSystemObject")

If FSObj.FileExists(cntpath) = True Then
	Set File = FSObj.OpenTextFile(cntpath,1,False,0)
	cnthit = File.ReadLine
	cnthit = cnthit*1
	Set File = FSObj.OpenTextFile(cntpath,2,False,0)
	File.Write cnthit
	File.Close
Else
	cnthit = "カウンタファイルが見つかりません。<br>" & "0"
End If

%><!-- #Include file="devtype.asp" --><%


 Select Case Provider
'	DoCoMo i-mode の場合
   Case "DoCoMo" :
%>
<html>
<head>
<title>書く[<%=BBSName %>]</title>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>">
<font color="<%=TitleColor %>"><%=BBSName %></font><br>
<font color="<%=TitleColor %>">―レスを書く―</font><br>
<form method="post" action="res.asp">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
名前:<br>
<input type="text" name="from" maxlength="64"><br>
E-mail:<br>
<input type="text"  name="mail" maxlength="128"><br>
題名:<br>
<input type="text" name="title" maxlength="128"><br>
本文:<br>
<textarea rows="10" cols="20" wrap="off" name="message"></textarea><br>
<br>
<input type="hidden" name="select" value="url" checked>
URL:<br>
<input type="text" name="url" value="http://" maxlength="256"><br>
<input type="submit" value="書く">
</form>
<hr>
<%=cnthit %>access<br>
<div align="right">
system by <a href="http://anoncom.net/">anoncom.net</a>
</div>
</body>
</html>
<%

'	Vodafone Vodafone live! の場合
   Case "Vodafone" :
%>
<html>
<head>
<title>書く[<%=BBSName %>]</title>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>">
<font color="<%=TitleColor %>"><%=BBSName %></font><br>
<font color="<%=TitleColor %>">―レスを書く―</font><br>
<form method="post" action="res.asp">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
名前:<br>
<input type="text" name="from" maxlength="64"><br>
E-mail:<br>
<textarea  rows="1" name="mail"></textarea><br>
題名:<br>
<textarea rows="1" name="title"></textarea><br>
本文:<br>
<textarea rows="10" cols="20" wrap="off" name="message"></textarea><br>
<br>
<input type="hidden" name="select" value="url" checked>
URL:<br>
<textarea rows="1" name="url">http://</textarea><br>
<input type="submit" value="書く">
</form>
<hr>
<%=cnthit %>access<br>
by <a href="http://anoncom.net/">anoncom.net</a>
</body>
</html>
<%

'	au EzWeb の場合
   Case "au" :
%>
<html>
<head>
<title>書く[<%=BBSName %>]</title>
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>">
<font color="<%=TitleColor %>"><%=BBSName %></font><br>
<font color="<%=TitleColor %>">―レスを書く―</font><br>
<form method="post" action="res.asp">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
名前:<br>
<input type="text" name="from" maxlength="64"><br>
E-mail:<br>
<input type="text" name="mail" maxlength="128"><br>
題名:<br>
<input type="text" name="title" maxlength="128"><br>
本文:<br>
<textarea rows="10" cols="20" wrap="off" name="message"></textarea><br>
<br>
<input type="hidden" name="select" value="url" checked>
URL:<br>
<input type="text" name="url" value="http://" maxlength="256"><br>
<input type="submit" value="書く">
</form>
<hr>
<%=cnthit %>access<br>
<div align="right">
system by <a href="http://anoncom.net/">anoncom.net</a>
</div>
</body>
</html>
<%
'	PCの場合
   Case Else :

If Request.Cookies("anoncom.BBS")("URL")="" Then
	URLVal = "http://"
Else
	URLVal = Request.Cookies("anoncom.BBS")("URL")
End If
 %><html lang="ja">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=x-sjis;">
<meta http-equiv="Content-Style-Type" content="text/css; charset=x-sjis;">
<title><%=BBSName %> by anoncomBBS</title>
<style type="text/css">
<!--//
a:link { text-decoration:none; color:<%=LinkColor %>; }
a:visited { text-decoration:none; color:<%=LinkColor %>; }
a:active { text-decoration:none; color:<%=ActiveLinkColor %>; }
a:hover { text-decoration:none; color:<%=HoverLinkColor %>; cursor:default; }

body{
	bakground-color:<%=BGColor %>;
	color:<%=TextColor %>
	overflow-y:auto;
	scrollbar-base-color:<%=BGColor %>;
	scrollbar-face-color:<%=BGColor %>;
	scrollbar-arrow-color:<%=BorderColor %>;
	scrollbar-highlight-color:<%=BGColor %>;
	scrollbar-3dlight-color:<%=BorderColor %>;
	scrollbar-shadow-color:<%=BorderColor %>;
	scrollbar-darkshadow-color:<%=BGColor %>;
	font-size:12px;
}

input,textarea,select,option{
	background:<%=BGColor %>;
	color:<%=TextColor %>;
	border-color:<%=BorderColor %>;
	border-width:1px;
	border-style:solid;
	font-size:12px;
	overflow-y:auto;

}

table#border{
	border-style:solid;
	border-width:1px;
	border-color:<%=BorderColor %>;
}

td{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','ＭＳゴシック';
}

td#tdname{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','ＭＳゴシック';
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

hr{
	color:<%=BorderColor %>;
}

//-->
</style>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<%
'正規表現によって、外から来た人の戻るページのURLを書き換える
'REFERER

HTTP_Referer = Request.ServerVariables("HTTP_REFERER")

Set RegObj = New RegExp

	RegObj.Global = True
	RegObj.IgnoreCase = False

	'DoCoMo i-mode
	RegObj.Pattern = "^(" & BBSURL & ")+"
	If RegObj.Test(HTTP_Referer) = True Then
		Ref_BackURL = "Site"
	Else
		Ref_BackURL = "other"
	End If

If Ref_BackURL = "Site" Then
%><a href="javascript:history.back()">&lt;&lt; back</a><br>
<%
Else
%><a href="<%=SiteURL %>" target="_top">&lt;&lt; back</a><br>
<%
End If
%>
<table border="0" id="border">
<tbody>
<tr>
<td align="left" valign="top">

<table border="0" id="border" cellspacing="0">
<tbody>
<tr>
<td align="left" valign="top">
<font color="<%=TitleColor %>" face="Times New Roman" size="+3"><b><i>Response</i></b></font><br>
<font color="<%=TitleColor %>" size="1">―――書き込む―――</font>
</td>
</tr><tr>
<td align="left" valign="top">
<form method="post" action="res.asp">
<input type="hidden" name="bbs" value="<%=BBSQuery %>">
<table border="0" cellspacing="0">
<tbody>
	<tr>
		<td align="right">名前：</td>
		<td align="left"><input type="text" name="from" size="45" value="<%=Request.Cookies("anoncom.BBS")("FROM") %>"></td>
	</tr><tr>
		<td align="right">E-mail：</td>
		<td align="left"><input type="text" name="mail" size="45" value="<%=Request.Cookies("anoncom.BBS")("MAIL") %>"></td>
	</tr><tr>
		<td align="right">タイトル：</td>
		<td align="left"><input type="text" name="title" size="45" value=""></td>
	</tr><tr>
<td align="left" valign="top" colspan="2">
本文：<br>
<textarea rows="14" cols="60" wrap="on" name="message"></textarea>
</td>
</tr><tr>
<td align="left" valign="top" colspan="2">
<input type="radio" name="select" value="id" id="id"><label for="id">ID:
<select name="site">
<option value="iland">魔法のｉらんど</option>
<option value="hamq">ハムスター島</option>
<option value="ihome">ihome</option>
<option value="freepe">ふり～ぺ</option>
<option value="bakudan">バクダンネット</option>
<option value="m-page">M-PAGE</option>
</select>
<input type="text" name="id" size="10">
</label>
</td>
</tr><tr>
<td align="left" valign="top" colspan="2">
<input type="radio" name="select" value="url" id="url" checked><label for="url">URL：
<input type="text" name="url" size="45" value="<%=URLVal %>"></label>
</td>
</tr><tr>
<td align="left" valign="top" colspan="2">
<input type="submit" value="書き込む"> 
<input type="reset" value="消去">
</td>
</tr>
</tbody>
</table>
</form>
<hr size="1" color="<%=BorderColor %>">
</td>
</tr><tr>
<td align="right" valign="top">
<i>system by <a href="http://anoncom.net/">anoncom.net</a></i>
</td>
</tr>
</tbody>
</table>

</td>

<td align="left" valign="top">
<%

'日付処理
Function WriteTime(dtmNow)

Dim strDate
strDate = Right(String(4,"0") & Year(dtmNow),4) & "/" & Right(String(2,"0") & Month(dtmNow),2) & "/" & Right(String(2,"0") & Day(dtmNow),2) & " " & Right(String(2,"0") & Hour(dtmNow),2) & ":" & Right(String(2,"0") & Minute(dtmNow),2)
WriteTime = strDate

End Function

'日付処理終了

cnt = 1
page = 1
Set rs = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM bbs_" & BBSQuery & " WHERE 'Num' ORDER BY sdat DESC"
rs.Open SQL, db, 3, 2

If rs.EOF = True Then
%>
<table border="0" width="100%" cellspacing="0" id="border">
<tbody>
<tr>
<td align="left" valign="top">
<font size="+3">最新レス数 <%=CntNum %>件</font>　/ 0件<br>
<font color="<%=TitleColor %>"><i><%=BBSName %></i></font>
</td>
</tr><tr>
<td align="right" valign="top">
<font size="-1">by anoncomBBS</font>
</td>
</tr><tr>
<td align="left" valign="top">
<table border="0" width="100%">
<tbody>
<tr>
<td align="left" valign="top" id="tdname">
0:■[anoncomBBS]
</td>
</tr><tr>
<td align="left" valign="top">
[No Script]
</td>
</tr><tr>
<td align="left" valign="top">
(記事の書き込みがありません)
</td>
</tr><tr>
<td align="right" valign="top">
[system]<br>
[<%=WriteTime(Now) %>]
</td>
</tr><tr>
<%
Else

CntSQL = "SELECT count(*) As RsCnt FROM bbs_" & BBSQuery
Set ResCount = db.Execute(CntSQL)
'総レスカウント数
RsCnt = ResCount("RsCnt")

rs.AbsolutePosition=cnt
%>
<table border="0" width="100%" cellspacing="0" id="border">
<tbody>
<tr>
<td align="left" valign="top">
<font <%=TextColor %> size="+2">最新レス数 <%=CntNum %>件</font>　/ <%=RsCnt %><br>
<font color="<%=TitleColor %>"><i><%=BBSName %></i></font>
</td>
</tr><tr>
<td align="right" valign="top">
<font size="2">by anoncomBBS</font>
</td>
</tr><tr>
<td align="left" valign="top">
<table border="0" width="100%" cellspacing="0">
<tbody>
<%
Do While Not rs.EOF
  If rs("abone")="True" Then
%><tr>
<td align="left" valign="top" id="tdname">
<%=rs("Num") %>:[<a href="mailto:<%=DeleteMailAddr %>"><%=DeleteName %></a>]
</td>
</tr><tr>
<td align="left" valign="top">
[<font color="<%=DelTitleColor %>"><%=DeleteTilte %></font>]
</td>
</tr><tr>
<td align="left" valign="top">
<font color="<%=DelBodyColor %>"><%=DeleteBody %></font>
</td>
</tr><tr>
<td align="right" valign="top">
[<%=DeleteDeviceType %>]<br>
[<%=WriteTime(rs("sdat")) %>]
</td>
</tr><%

  Else
    If rs("from")<>"" Then
      If rs("mail")<>"" Then
%><tr>
<td align="left" valign="top" id="tdname">
<%=rs("Num") %>:[<a href="mailto:<%=rs("mail") %>"><%=rs("from") %></a>]
<td>
</tr><%
      Else
%><tr>
<td align="left" valign="top" id="tdname">
<%=rs("Num") %>:[<%=rs("from") %>]
</td>
</tr><%

      End If
    Else
      If rs("mail")<>"" Then

%><tr>
<td align="left" valign="top" id="tdname">
<%=rs("Num") %>:[<a href="mailto:<%=rs("mail") %>"><%=rs("mail") %></a>]
</td>
</tr><%
      Else
%><tr>
<td align="left" valign="top" id="tdname">
<%=rs("Num") %>:[<%=NotFoundName %>]
</td>
</tr><%
      End If
    End If
    If rs("title")<>"" Then
%><tr>
<td align="left" valign="top">
[<%=rs("title") %>]
</td>
</tr><%

    End If


	If TagUse = 1 Then
		'タグ有効

	    If LenB(rs("message")) > 500 Then
	        message = "<br>" & LeftB(rs("message") ,500) & "...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & """>続き</a>" & vbCrLf
	    Else
		message = "<br>" & rs("message") & vbCrLf
	    End If
	message = Replace(message,vbCrLf,"<br>" & vbCrLf)
        message = message & "<br>" & vbCrLf

	Else
		'タグ無効

		'タグ表示
		If TagSourceView = 1 Then
			Set bsp = Server.CreateObject("basp21")	'BASPを読み込み
			message = bsp.RepTagChar(rs("message"))
			Set bsp = Nothing
		Else
		'タグ非表示
			'タグ部分置き換え
			message = RemoveHTML(rs("message"))
		End If
		
	    If LenB(rs("message")) > 500 Then
	        message = Replace(LeftB(message,500),vbCrLf,"<br>" & vbCrLf) & _
			"...<a href=""resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num") & _
			""">続き</a><br>" & vbCrLf
	    Else
	        message = Replace(message,vbCrLf,"<br>" & vbCrLf)
	    End If
	End If

%><tr>
<td align="left" valign="top">
<%=message %>
</td>
</tr><%

    If rs("url")<>"" Then
      If rs("url")<>"http://" Then
%><tr>
<td align="right" valign="top">
<a href="<%=rs("url") %>" target="_blank">Homepage</a>
</td>
</tr><%

      End If
    End If

%><tr>
<td align="right" valign="top">[<%=rs("UA") %>]</td>
</tr>
<tr>
<td align="right" valign="top">[<%=WriteTime(rs("sdat")) %>]</td>
</tr><%


  End If
  rs.MoveNext
  cnt = cnt + 1
  If cnt = page * 10 + 1 Then Exit Do
Loop
End If
%>
</tbody>
</table>
</td>
</tr><tr>
<td align="right" valign="top">
<hr size="1">
<font size="-1"><i>Find <%=cnthit %>access</i></font>
</td>
</tr><tr>
<td align="right" valign="top">
<font size="-1"><i>Write <%=RsCnt %>response</i></font>
</td>
</tr>
</tbody>
</table>

</td>
</tr>
</tbody>
</table>
</body>
</html>
 <%
End Select


Else						'フォームが""でない場合

'処理開始時間取得
ExStart = Timer

BBSQuery = Request.Form("bbs")

If BBSQuery = "" Then
	BBSQuery = "default"
End If


'--------------------------------------------------------------------------
    If Request.Form("message")="" Then		'本文未記入の場合
      ExecuteResponseMsg = "<font color=""" & TitleColor & """<b>本文</b></font>"
      ExecuteResponseMsg = ExecuteResponseMsg & "に何も書いてないです。<br>" & vbCrLf
      ExecuteResponseMsg = ExecuteResponseMsg & "前の画面に戻り、必要箇所を記入してください。"
						'以下は本文に記入がある場合
    Else


'変数定義
         from = Request.Form("from")		'from=名前
         title = Request.Form("title")		'title=題名
         mail = Request.Form("mail")		'mail=メールアドレス
						'message=本文
     	message = Request.Form("message")


	'[']シングルクォーテーションを['']に置き換える(SQLミス防止のため)
	message = Replace(message,"'","''")
	''["]ダブルクォーテーションを[""]に置き換える(SQLミス防止のため)
	'message = Replace(message,"""","""""")

If Request.Form("select")="id" Then

	ID = Request.Form("id")

	If ID = "" Then
		url = ""
	Else

		Select Case Request.Form("site")

							'魔法のｉらんど
		Case "iland" :
			url = "http://ip.tosp.co.jp/i.asp?i=" & ID

							'ハムスター島
		Case "hamq" :
			url = "http://www.hamq.jp/i.cfm?i=" & ID

							'ｉｈｏｍｅ
		Case "ihome" :
			url = "http://ihome.to/" & ID

							'ふり～ぺ
		Case "freepe" :
			url = "http://www.freepe.com/ii.cgi?" & ID

							'バクダンネット
		Case "bakudan" :
			url = "http://bakudan.net/" & ID

							'M-PAGE
		Case "m-page" :
			url = "http://k.m-page.jp/m.asp?U=" & ID


		End Select

	End If


ElseIf Request.Form("select")="url" Then

         url = Request.Form("url")		'url=ＵＲＬ
End If


'端末振り分け
%><!-- #Include file="devtype.asp" --><%



	Select Case Provider
		Case "DoCoMo" : UA="i"			'DoCoMo i-modeは"UA=i"
		Case "Vodafone" : UA="v"		'Vodafone Vodafone live!は"UA=v"
		Case "au" : UA="ez"			'KDDI EzWebは"UA=ez"
		Case "PC" : UA="Pc"			'PC Mozillaは"UA=Pc"
		Case Else : UA="??"			'その他エージェントは"UA=??"
	End Select
						'UserAgent=ユーザーエージェント
	UserAgent = Request.ServerVariables("HTTP_USER_AGENT")
						'IP=リモートIP
	IP = Request.ServerVariables("REMOTE_ADDR")
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
<%
			'   Host=リモートホスト(IP)
	Host = Rmt_Host


'--------------------------------------------------------------------------

'データベースへの書き込み作業

	DB_FILE = BBSDBFileName	'データベースファイル


	anoncomBBS = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & DB_FILE
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open anoncomBBS



				'BBSデータベースに
					'[名前(from)]
					'[メール(mail)]
					'[題名(title)]
					'[本文(message)]
					'[URL(url)]
					'[書き込み日付(sdat)]
					'[リモートIP(IP)]
					'[リモートホスト(Host)]
					'[ユーザーエージェント(UserAgent)]
					'[エージェントタイプ(UA)]
				'の順に書き込む
	SQL = "INSERT INTO bbs_" & BBSQuery & "([from],[mail],[title],[message],[url],[sdat],[IP],[Host],[UserAgent],[UA]) "
	SQL = SQL & "VALUES('" & from & "','" & mail & "','" & title & "','" & message & "','" & url & "',#" & Now & "#,'" & IP & "','" & Host & "','" & UserAgent & "','" & UA & "')"
	
	conn.Execute(SQL)
	conn.Close


'通知配信処理
If BBSMailSend = 1 Then
message = Replace(message,"<br>", vbCrLf)

'---------------------------------------------------------------------

'message = RemoveHTML(message)

If MailBBSBodyCut = 1 Then
	If LenB(message) >= 350 Then

		Set db = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		db.Provider = "Microsoft.Jet.OLEDB.4.0"
		db.Mode = 1
		db.ConnectionString = BBSDBFileName
		db.Open
		SQL = "SELECT * FROM bbs_" & BBSQuery & " WHERE 'Num' ORDER BY sdat DESC"
		Set rs = db.Execute(SQL)

		message = LeftB(message,350) & "..." & vbCrLf
		If BBSQuery = "" Or BBSQuery = "default" Then
			ResURL = "続きは" & BBSURL & "resview.asp?no=" & rs("Num")
		Else
			ResURL = "続きは" & BBSURL & "resview.asp?bbs=" & BBSQuery & "&no=" & rs("Num")
		End If

		db.Close
		Set rs = Nothing
		Set db = Nothing

		message = message & ResURL

	End If
End If




  Set bsp = Server.CreateObject("basp21")
  strSrv = MailServer


'データベースから一人ずつアドレスを読み取り、1通ずつ送信する

Set rs_user = Server.CreateObject("ADODB.Recordset")
'掲示板の対応グループごとに配信先を分ける

If UserGroup = "all_user" Then

	userSQL = "SELECT * FROM users_tb WHERE [act_flag] >= 9 ORDER BY [SERIAL] ASC"
	

	Set rs_user = db.Execute(userSQL)

	'送信開始時刻を取得
	SendStartTime = Timer
	'エラー件数カウントリセット
	ErrCnt = 0
	ErrCnt = CInt(ErrCnt)

	'== ループ開始
	Do While Not rs_user.EOF = True


		'strTo = "bcc" & vbTab & "<" & SendToAddr & ">" & vbTab & ">Return-Path: <" & MailFromAddr & ">"
 		strTo = "bcc" & vbTab & "<" & rs_user("mail") & ">" & vbTab & ">Return-Path: <" & MailFromAddr & ">"

		'アドレスが入力されている場合はそのアドレスから送信する
		'If mail <> "" Then
		'	strFrm = "<" & mail & ">"
		'Else
			strFrm = BBSName & "<" & MailFromAddr & ">"
		'End If
		strSbj = BBSName
		If from = "" Then
			strBdy = "■" & NotFoundName & vbCrLf
		Else
			strBdy = "■" & from & vbCrLf
		End If
		If title <> "" Then
			strBdy = strBdy & "[" & title & "]" & vbCrLf
		End If

		If BBSQuery = "" Or BBSQuery = "default" Then
			strBdy = strBdy & "---" & vbCrLf & _
				message & vbCrLf & vbCrLf & _
				"" & BBSURL
		Else
			strBdy = strBdy & "---" & vbCrLf & _
				message & vbCrLf & vbCrLf & _
				"" & BBSURL & "?" & BBSQuery
		End If

		strFL = ""

		If debug_flag >= 8 Then
			'Debugモード8以上は送信せず、送信先の書き出し
			SndAddr = SndAddr & rs_user("mail") & vbCrLf
		Else
			'Debugモード8未満は通常送信
			lngRst = bsp.SendMail(strSrv, strTo, strFrm, strSbj, strBdy, strFl)
			ErrorStatus = ErrorStatus & lngRst & vbCrLf

			'エラー件数カウント
			If lngRst <> "" Then
				ErrCnt = ErrCnt + 1
			End If
		End If

		'レコードの移動
		rs_user.MoveNext

	Loop


Else
	AryGroup = Split(UserGroup, ",", -1, vbTextCompare)
	For i = 0 To UBound(AryGroup)
		userSQL = "SELECT * FROM users_tb WHERE [act_flag] = 9 AND [GroupID] = '" & AryGroup(i) & "' ORDER BY [SERIAL] ASC"


		Set rs_user = db.Execute(userSQL)

		'送信開始時刻を取得
		SendStartTime = Timer
		'エラー件数カウントリセット
		ErrCnt = 0
		ErrCnt = CInt(ErrCnt)

		'== ループ開始
		Do While Not rs_user.EOF = True


			'strTo = "bcc" & vbTab & "<" & SendToAddr & ">" & vbTab & ">Return-Path: <" & MailFromAddr & ">"
	 		strTo = "bcc" & vbTab & "<" & rs_user("mail") & ">" & vbTab & ">Return-Path: <" & MailFromAddr & ">"

			'アドレスが入力されている場合はそのアドレスから送信する
			'If mail <> "" Then
			'	strFrm = "<" & mail & ">"
			'Else
				strFrm = BBSName & "<" & MailFromAddr & ">"
			'End If
			strSbj = BBSName
			If from = "" Then
				strBdy = "■" & NotFoundName & vbCrLf
			Else
				strBdy = "■" & from & vbCrLf
			End If
			If title <> "" Then
				strBdy = strBdy & "[" & title & "]" & vbCrLf
			End If

			If BBSQuery = "" Or BBSQuery = "default" Then
				strBdy = strBdy & "---" & vbCrLf & _
					message & vbCrLf & vbCrLf & _
					"" & BBSURL
			Else
				strBdy = strBdy & "---" & vbCrLf & _
					message & vbCrLf & vbCrLf & _
					"" & BBSURL & "?" & BBSQuery
			End If

			strFL = ""

			If debug_flag >= 8 Then
				'Debugモード8以上は送信せず、送信先の書き出し
				SndAddr = SndAddr & rs_user("mail") & vbCrLf
			Else
				'Debugモード8未満は通常送信
				lngRst = bsp.SendMail(strSrv, strTo, strFrm, strSbj, strBdy, strFl)
				ErrorStatus = ErrorStatus & lngRst & vbCrLf

				'エラー件数カウント
				If lngRst <> "" Then
					ErrCnt = ErrCnt + 1
				End If
			End If

			'レコードの移動
			rs_user.MoveNext

		Loop

	'== ループ終了

	Set rs_user= Nothing

	'グループのループ終了
	Next
End If



'送信終了時間
SentTime = Timer


'Debugモード3以上は送信結果を表示
If debug_flag >= 3 Then


  If ErrCnt > 0 Then
	Response.Write lngRst
	'If sent to mail failed msg
	SendErrMsg = "エラーにより一部配信出来ませんでした。<br>" & vbCrLf & _
			"エラーメッセージは " & ErrCnt & "件です。<br>" & vbCrLf

	'Debugモード5以上時はエラーの詳細を表示する
	If debug_flag >= 5 Then
		ErrDebugMsg = 	"詳細は以下のとおりです：<br>" & vbCrLf  &_
		Replace(ErrorStatus, vbCrLf, "<br>" & vbCrLf)
		SendErrMsg = SendErrMsg & "<br>" & vbCrLf & ErrDebugMsg & "<br>" & vbCrLf
	End If


  Else
	'If sent to mail successed msg
	SendErrMsg = "メッセージは正しく送信されました。<br>" & vbCrLf
  End If

End If


'Debugモード8以上は送信せず、送信内容の書き出し
If debug_flag >= 8 Then

	'送信先を書き出す
	'改行ごとに区切る
	arySndTo = Split(SndAddr, vbCrLf)
	'配列の最大値
	arySndToUBound = UBound(arySndTo)
	'配列の数だけ吐き出す
	For i = 0 To arySndToUBound
		SndTo = SndTo & "<a href=""mailto:" & arySndTo(i) & """>" & _
			arySndTo(i) & "</a><br>" & vbCrLf
	Next

	DebugMsg = "送信元: <a href=""mailto:" & MailFromAddr & """>" & MailFromAddr & _
		"</a><br>" & vbCrLf & _
		"送信先:<br>" & vbCrLf & _
		SndTo & _
		"<br>" & vbCrLf & _
		"[件名]<br>" & vbCrLf & _
		strSbj & "<br>" & vbCrLf & _
		"<br>" & vbCrLf & _
		"[本文]" & "<br>" & vbCrLf & _
		Replace(strBdy, vbCrLf, "<br>" & vbCrLf) & _
		"<br>" & vbCrLf

End If

'送信処理時間を計算
SendingTime = SentTime - SendStartTime
SendTimeMsg = "<br>うち送信処理時間 " & Round(CDbl(SendingTime), 3) & "秒"

End If


'クッキーに記録
 Response.Cookies("anoncom.BBS")("From") = from
 Response.Cookies("anoncom.BBS")("Mail") = mail
 Response.Cookies("anoncom.BBS")("URL") = url
 Response.Cookies("anoncom.BBS").Expires = Date + 365

'書き込み終了時
 ExecuteResponseMsg = "書き込みしました。"

    End If

'処理終了時間取得
ExEnd = Timer

'処理経過時間計算
ExTime = ExEnd - ExStart
ExTimeMsg = "処理時間 " & Round(CDbl(ExTime), 3) & "秒"

'デバッグモード5以上で処理実行時間を表示
If debug_flag >= 5 Then
	DebugTimeMsg = "<i>" & ExTimeMsg & vbCrLf & SendTimeMsg & "</i><br>" & vbCrLf
End If

'処理結果の文字列での出力
%>
<html lang="ja">
<head>
<meta http-equiv="cache-control" content="nocache">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis;">
<meta http-equiv="Content-Style-Type" content="text/css;">
<title><%=BBSName %></title>
<style>
<!--//
a:link { text-decoration:none; color:<%=LinkColor %> }
a:visited { text-decoration:none; color:<%=LinkColor %> }
a:active { text-decoration:none; color:<%=ActiveLinkColor %> }
a:hover { color:<%=ActiveLinkColor %>; cursor:default; }

body{
	bakground-color:<%=BGColor %>;
	color:<%=TextColor %>
	font-size:11px;
	font-family:"MS UI Gothic","ＭＳ ゴシック"
	overflow-y:auto;
	scrollbar-base-color:<%=BGColor %>;
	scrollbar-face-color:<%=BGColor %>;
	scrollbar-arrow-color:<%=BorderColor %>;
	scrollbar-highlight-color:<%=BGColor %>;
	scrollbar-3dlight-color:<%=BorderColor %>;
	scrollbar-shadow-color:<%=BorderColor %>;
	scrollbar-darkshadow-color:<%=BGColor %>;
}

//-->
</style>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<%=ExecuteResponseMsg %><br><%
Response.Write SendErrMsg %><%=DebugMsg %><%=DebugTimeMsg %>
<div align="right"><%
If Request.Form("bbs") = "" Then
%>
<a href="<%=BBSURL %>">掲示板</a><br><%
Else
%>
<a href="<%=BBSURL %>?<%=BBSQuery %>">掲示板</a><br><%
End If
%>
<a href="<%=SiteURL %>">サイトトップ</a>
</div>
</body>
</html>
<%
End If


End If

%>

