<%@ Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'携帯の場合は携帯用の管理ページへ移動
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'バッファを有効にする
Response.Buffer = True


'ログイン前はカラー情報などは標準設定を使う

If Session("login") <> 1 And Request.Form = "" Then	'ログイン前のみ適応

Set rs_set = Server.CreateObject("ADODB.Recordset")
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'default_settings'",db,3,2

'標準の設定を読み込む
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
	font-family:'MS UI Gothic','ＭＳゴシック';
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
	font-family:'MS UI Gothic','ＭＳゴシック';
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

	If Request.Form = "" Then	'ログインポスト前

'ログイン画面

'標準の設定を読み込む
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
管理者ID：<input type="text" name="id" size="8"><br>
Password：<input type="password" name="pw" size="8"><br>
<input type="submit" value="ログイン">
</form>
<br>
<a href="bbs.asp">掲示板へ</a><br>
<hr size="1">
system by <a href="http://anoncom.net/">anoncom.net</a>
</body><%

	Else
	'ログイン処理
		'管理情報読み込み
		Set rs_admin = Server.CreateObject("ADODB.Recordset")
		rs_admin.Open "SELECT * FROM admin_settings " & _
			"WHERE adminID = '" & Request.Form("id") & "' AND " & _
			"adminPass = '" & Request.Form("pw") & "'", db, 3, 2

		'認証完了
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
			Session.TimeOut = 30	'タイムアウトは30分
		End If

		Set rs_admin = Nothing

		Response.Redirect "admin.asp"

	End If

Else
'セッションを確認し、ログインしている状態
	If Request("bbs") <> "" Then
		Session("BBSQuery") = Request("bbs")
		Session("BBSNo") = CInt(0)
		Response.Redirect "admin.asp"
	End If

	BBSQuery = Session("BBSQuery")
	BBSSelectNo = Session("BBSNo")


	If Session("AdminLevel") = 0 Then
		'AdminRank=0は、ログイン時のセッションをすべて破棄し、ログイン排除
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
このIDは現在停止処置を施されているため、管理画面を表示することはできません。<br>
原因は次のものが考えられます：<br>
<br>
・管理者の利用制限によるアカウントの一次利用停止<br>
・悪質なユーザーからの権限剥奪<br>
・掲示板管理機能のデバッグのための一時的な停止処置<br>
<br>
<br>
なお、このページは15秒後に掲示板のトップへ移動します。<br>
</body>
<%
	Else

		Select Case Request.QueryString

			'メニュー画面
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
<b>掲示板：</b><br>
<select name="bbs" onchange="this.form.submit()">
<option value="">(掲示板...)</option>
<%

SQL = "SELECT [bbs_table], [BBSName] FROM [settings] WHERE [bbs_table] Like 'bbs_%' ORDER BY [SERIAL] ASC"
Set rs = db.Execute(SQL)

'ループカウンタ初期値
i = 1

If rs.EOF = False Then
	Do While Not rs.EOF
		bbstable = Replace(rs("bbs_table"), "bbs_", "")
		BBSName = rs("BBSName")

		If Session("AdminBBS") = "allbbs" Then

			'カウンタと掲示板シリアルNoが一致すればselect
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
<option value="">掲示板が作成されていません</option><%
End If
%>
</select>
<noscript><input type="submit" value="選択"></noscript></form></td>
</tr><%
	If BBSQuery <> "" Then

%><tr>
<td align="center">
<b><a href="admin.asp?main" target="adminmain">トップ</a></b>
</td>
</tr><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=bbs" target="adminmain">掲示板</a></b>
</td>
</tr><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=access" target="adminmain">書き込み解析</a></b>
</td>
</tr><%
		If Session("AdminLevel") > 1 Then
%><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=delete" target="adminmain">削除管理</a></b>
</td>
</tr><%
		End If
%><%
		If Session("AdminLevel") > 3 Then
%><tr>
<td align="center">
<b><a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=setting" target="adminmain">掲示板設定</a></b>
</td>
</tr><%
		End If
%><%
		If Session("AdminLevel") > 4 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=clear" target="adminmain"><font color="<%=TitleColor %>">書き込み全消去</font></a>
</td>
</tr><tr>
<td align="center">&nbsp;</td>
</tr><tr>
<td align="center">
<a href="bbsadmin.asp?bbs=<%=BBSQuery %>&mode=settingclear" target="adminmain">
<font color="<%=TitleColor %>">設定の初期化</font></a>
</td>
</tr><%
		End If



	End If

%><tr>
<td align="center">&nbsp;</td>
</tr><%
	If Session("AdminLevel") > 6 Then
%><tr>
<td align="center"><a href="setbbs.asp?mode=create" target="adminmain">新規掲示板作成</a></td>
</tr><%
	End If
%><%
	If Session("AdminLevel") > 7 Then
%><tr>
<td align="center"><a href="setbbs.asp?mode=delete" target="adminmain">掲示板削除</a></td>
</tr><%
	End If
%><%
	If Session("AdminLevel") > 2 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=adminsetting" target="adminmain">管理者情報変更</a>
</td>
</tr><%
		End If
%><%
	If Session("AdminLevel") >= 9 Then
%><tr>
<td align="center">
<a href="bbsadmin.asp?mode=adminuserreg" target="adminmain">管理者ユーザ管理</a>
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
<font color="<%=TitleColor %>">データベース移動</font></a>
</td>
</tr><%
	End If

%><tr>
<td align="center">
<a href="admin.asp?logout" target="_top">ログアウト</a>
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
'トップメイン画面
Case "main"
%>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="<%=LinkColor %>" alink="<%=ActiveLinkColor %>" vlink="<%=LinkColor %>">
<font color="<%=TitleColor %>" face="Times New Roman"><b><i>anoncomBBS</i></b></font><br>
<font color="<%=TitleColor %>" size="+3" face="Times New Roman"><b><i>BBS Administrator for <%=BBSName %></i></b></font><br>
<br>
<br>
<br>
左のメニューから、設定する項目をクリックしてください。
</body>
<%
'ログアウト処理
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

'URLにpageクエリがない場合
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
