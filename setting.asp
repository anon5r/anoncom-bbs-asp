<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'携帯の場合は携帯用の管理ページへ移動
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'バッファを有効にする
Response.Buffer = True

If BBSBlank = True Then
	Response.Redirect "blankbbs.html"
End If

'ログインセッションが確立しているか確認
If Session("login") = 1 Then

%>
<html lang="ja">
<head>
<meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/html; charset=shift_jis">
<title>設定変更ツール</title>
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
<td align="center" bgcolor="#88ff88">設定項目</td>
<td align="center" bgcolor="#ffaaaa">設定値</td>
<td align="center" bgcolor="#aaaaff">説明・備考</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">サイト名：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="sitename" value="<%=SiteName %>" maxlength="100" size="20"></td>
<td align="left" bgcolor="#ccccff">※100文字まで</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">サイトURL：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="siteurl" value="<%=SiteURL %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">※半角英数字</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">掲示板名：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="bbsname" value="<%=BBSName %>" maxlength="100" size="20"></td>
<td align="left" bgcolor="#ccccff">※100文字まで</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">掲示板コメント：</td>
<td align="left" bgcolor="#ffcccc"><textarea name="bbscomment" rows="2" cols="20"><%=BBSComment %></textarea></td>
<td align="left" bgcolor="#ccccff">※100文字まで</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">掲示板のURL：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="baseurl" value="<%=BBSURL %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">※半角英数字</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">掲示板の状態：</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case BBSStatus
	Case 9: status9 = " selected"
	Case 3: status3 = " selected"
	Case 0: status0 = " selected"
End Select %><select name="bbsstatus">
<option value="9"<%=status9 %>>閲覧・書き込み可</option>
<option value="3"<%=status3 %>>書き込み不可</option>
<option value="0"<%=status0 %>>閲覧・書き込み不可</option>
</select></td>
<td align="left" bgcolor="#ccccff">掲示板の運営状態を変更します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">デバッグモード：</td>
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
<option value="0"<%=debug0 %>>デバッグなし</option>
<option value="3"<%=debug3 %>>簡易デバッグ</option>
<option value="5"<%=debug5 %>>デバッグ</option>
<option value="9"<%=debug9 %>>完全デバッグ</option>
</select></td>
<td align="left" bgcolor="#ccccff">デバッグモードでの実行可否を変更します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">背景色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="bgcolor" value="<%=BGColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">例：<i>#ffffff</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">文字の色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="textcolor" value="<%=TextColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">例：<i>#000000</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">リンクの色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="linkcolor" value="<%=LinkColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">例：<i>#0000ff</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">アクティブリンクの色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="alinkcolor" value="<%=ActiveLinkColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">例：<i>#ffff00</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">線の色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="bordercolor" value="<%=BorderColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">例：<i>#888888</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">掲示板名の色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="titlecolor" value="<%=TitleColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">例：<i>#ff0000</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">1ページの表示件数：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="cntnum" value="<%=CntNum %>" maxlength="3" size="3"></td>
<td align="left" bgcolor="#ccccff">※半角数字</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">カウンタファイル名：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="countfilename" value="<%=CountFileName %>"></td>
<td align="left" bgcolor="#ccccff">カウンタファイル名を指定します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">掲示板でのタグの使用：</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case TagUse
	Case 1: TUon = " selected"
	Case 0: TUoff = " selected"
End Select
%><select name="tag">
<option value="1"<%=TUon %>>有効</option>
<option value="0"<%=TUoff %>>無効</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">無効にすると書き込みされたタグがすべて無効になります。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">タグの表示：</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case TagSourceView
	Case 1: TSVon = " selected"
	Case 0: TSVoff = " selected"
End Select
%><select name="tagsourceview">
<option value="1"<%=TSVon %>>表示</option>
<option value="0"<%=TSVoff %>>非表示</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">表示にするとタグのソースが表示されます。（タグ使用が「無効」の場合のみ有効）</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">書き込み通知配信：</td>
<td align="left" bgcolor="#ffcccc"><%
Select Case BBSMailSend
	Case 1: BMSon = " selected"
	Case 0: BMSoff = " selected"
End Select
%><select name="bbsmailsend">
<option value="1"<%=BMSon %>>有効</option>
<option value="0"<%=BMSoff %>>無効</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">有効にすると書き込みが指定のアドレスへ通知配信されます。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">メールサーバ：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="mailserver" value="<%=MailServer %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">通知配信を利用する場合は必ず指定してください！</td>
</tr><!--<tr>
<td align="right" bgcolor="#ccffcc">送信先アドレス：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="sendtoaddr" value="<%=SendToAddr %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">※通知配信の送信先アドレス</td>
</tr>--><tr>
<td align="right" bgcolor="#ccffcc">送信先グループ：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="SendGroup" value="<%=UserGroup %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">※通知配信の送信先グループ</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">送信元アドレス：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="mailfromaddr" value="<%=MailFromAddr %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">通知配信利用時に、ここで指定したアドレスからメールが送られてきます。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">通知メールカット：</td>
<td align="left" bgcolor="#ffcccc"><%
	Select Case MailBodyCut
		Case 1: MBCon = " selected"
		Case 0: MBCoff = " selected"
	End Select
%><select name="mailbbsbodycut">
<option value="1"<%=MBCon %>>有効</option>
<option value="0"<%=MBCoff %>>無効</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">有効にすると書き込みの本文が長い場合、本文省略して配信されます。i-modeで分割設定していない場合に特に便利です。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">名無しの名前：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="notfoundname" value="<%=NotFoundName %>" maxlength="40" size="20"></td>
<td align="left" bgcolor="#ccccff">名前がない場合、名前部分に表示する文字列を指定します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">名前：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="delname" value="<%=DeleteName %>" maxlength="50" size="20"></td>
<td align="left" bgcolor="#ccccff">削除レスの名前部分に表示する文字列を指定します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">メールアドレス：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="delmailaddr" value="<%=DeleteMailAddr %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">削除レスのメールアドレスに表示する文字列を指定します。指定しなくてもかまいません。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">タイトル：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="deltitle" value="<%=DeleteTilte %>" maxlength="100" size="40"></td>
<td align="left" bgcolor="#ccccff">削除レスの代わりのタイトル部分に表示する文字列を指定します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">本文：</td>
<td align="left" bgcolor="#ffcccc"><textarea name="delbody" rows="2" cols="20"><%=DeleteBody %></textarea></td>
<td align="left" bgcolor="#ccccff">削除レスの本文部分に表示する文字列を指定します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">端末：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="deldevtype" value="<%=DeleteDeviceType %>"></td>
<td align="left" bgcolor="#ccccff">削除レスの端末部分に表示する文字列を指定します。</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">タイトルの色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="deltitlecolor" value="<%=DelTitleColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">削除レスのタイトルの色。例：<i>#FF99AA</i></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">本文の色：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="delbodycolor" value="<%=DelBodyColor %>" maxlength="20" size="10"></td>
<td align="left" bgcolor="#ccccff">削除レスの本文の色。例：<i>#ff0000</i></td>
</tr>
</tbody>
</table>
<br>
<input type="submit" value="　変　　更　">
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
<td align="center" bgcolor="#88ff88">設定項目</td>
<td align="center" bgcolor="#ffaaaa">設定値</td>
<td align="center" bgcolor="#aaaaff">説明・備考</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">管理者ID：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminid" value="<%=Session("AdminID") %>" maxlength="20" size="20"></td>
<td align="left" bgcolor="#ccccff">管理者ログイン用ID</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">パスワード：</td>
<td align="left" bgcolor="#ffcccc"><input type="password" name="adminpass" value="<%=Session("AdminPass") %>" maxlength="10" size="16"></td>
<td align="left" bgcolor="#ccccff">管理用パスワード</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">管理者名：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminname" value="<%=Session("AdminName") %>" maxlength="20" size="30"></td>
<td align="left" bgcolor="#ccccff">管理者名</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">メールアドレス：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminmail" value="<%=Session("AdminMail") %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">管理者メールアドレス</td>
</tr>
</tbody>
</table>
<br>
<input type="submit" value="　変　　更　">
</form>
<%

	End Select


'ポスト後 ****************************************************************
Else

   Select Case Request.QueryString("set")

	Case "bbs"
		'掲示板設定

		Set rs_set = Server.CreateObject("ADODB.Recordset")
		upSQL = "SELECT * FROM settings WHERE bbs_table = 'bbs_" & BBSQuery & "'"
		rs_set.Open upSQL,db,3,3

		'設定の反映
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
		'管理者設定

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

			'設定の反映
			rs_set("AdminID") = Request.Form("adminID")
			rs_set("AdminPass") = Request.Form("adminPass")
			rs_set("AdminName") = Request.Form("adminName")
			rs_set("AdminMail") = Request.Form("adminMail")

			rs_set.Update

			rs_set.Close
			Set rs_set = Nothing

			'セッション管理者情報も更新
			Session("adminID") = Request.Form("adminID")
			Session("adminPass") = Request.Form("adminPass")
			Session("adminName") = Request.Form("adminName")
			Session("adminMail") = Request.Form("adminMail")


		Else
			%>その管理者は見つからない為、更新できませんでした。<%
		End If


   End Select

%>
設定が変更されました。<br>
<a href="admin.asp" target="_top">[BBS Admin]</a>
<% End If %>
</body>
</html><%

Else
	Response.Redirect "admin.asp"
End If %>
