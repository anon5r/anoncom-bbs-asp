<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'バッファ有効
Response.Buffer = True

Set Fso = Server.CreateObject("Scripting.FileSystemObject")


'携帯の場合は携帯用の管理ページへ移動
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'ログインセッションは確立しているか
If Session("login") = 1 Then

Select Case Request("mode")

	Case "create"
		title_j = "新規掲示板作成"
		title_e = "Create New BBS"

	Case "delete"
		title_j = "掲示板削除"
		title_e = "Delete BBS"

	Case Else
		Response.Redirect "admin.asp"
End Select


Set rs_set = db.Execute("SELECT * FROM settings WHERE bbs_table = 'default_settings'")

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
	'ポストされる前
	If Request.Form = "" Then

Select Case Request.QueryString("mode")
	Case "create"

	'掲示板作成権限はあるか
	If CInt(Session("AdminLevel")) <= 6 Then
	%><b>このユーザは掲示板作成権限がありません。</b>
	<%
	Else
					'掲示板作成
	%>
<form action="setbbs.asp" method="post">
<input type="hidden" name="mode" value="create">
<table border="0">
<tr>
<td align="right" valign="top">新規に作成したいBBS IDを入力してください：</td>
<td align="left" valign="top">
<input type="text" name="bbsid" size="20" maxlength="16">
<input type="submit" value="作成">
</td>
<td align="left" valign="top"><font color="#ff0000">※半角英数字　最大16文字</font></td>
</tr>
</table>
</form>

<%
	End If

	Case "delete"
	'掲示板削除権限はあるか
	If CInt(Session("AdminLevel")) <= 7 Then
	%><b>このユーザは掲示板削除権限がありません。</b>
	<%
	Else
		'掲示板削除
%>
<script language="JavaScript"><!--
function DeleteBBS(){
  msgRet = confirm("本当によろしいですか？");
  if ( msgRet == true ){
	document.delform.submit();
  }
}
// --></script>

<form action="setbbs.asp" name="delform" method="post">
<input type="hidden" name="mode" value="delete">
<table border="0">
<tr>
<td align="right" valign="top">削除するBBS IDを選択してください：</td>
<td align="left" valign="top">
<select name="bbsid">
<option value="">---選択してください---</option><%

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
<option value="">掲示板がありません</option><%
End If
%>
</select>
<input type="button" value="削除" onclick="DeleteBBS()">
</td>
</tr>
</table>
</form>
	<%
	End If


	End Select
Else
		'ポスト後

	BBSID = Request.Form("bbsid")

Select Case Request.Form("mode")
	'掲示板作成処理
	Case "create"
		'Set Rs = db.OpenSchema(20)	'テーブル情報取得

SQLs = "SELECT bbs_table, BBSName FROM settings WHERE bbs_table = 'bbs_" & BBSID & "'"

		Set RecSet = db.Execute(SQLs)
		If RecSet.EOF = False Then
			%>その掲示板は既に作成されています。<br>
掲示板ID：<b><%=Replace(RecSet("bbs_table"), "bbs_", "") %></b><br>
掲示板名：<b><%=RecSet("BBSName") %></b><br><%
		Else
		'掲示板作成処理

		'テーブル作成

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

						'設定テーブルに新規の設定を書き込む
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
	"'" & Rs("BBSName") & "', '何でも書いてね～♪'," & _
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
掲示板：<%=BBSID %> を作成しました。<br>
<a href="<%=BBSURL %>bbs.asp?bbs=<%=BBSID %>" target="_blank"><%=BBSURL %>bbs.asp?bbs=<%=BBSID %></a><br>
<br>
<a href="<%=BBSURL %>admin.asp" target="_top">[掲示板設定画面]</a><br>
<%

		End If

	'掲示板削除処理
	Case "delete"

		If BBSID = "" Then
			%>掲示板が選択されていません。<%
		Else
			If BBSID = "default" Then
				%><font size="+2">その掲示板は削除できません。</font><br>
default掲示板は削除することはできません。<br>
使用したくない場合は<b>閲覧・書き込み不可</b>設定にしてください。<%
			Else

				SQLs = "SELECT bbs_table, BBSName FROM settings WHERE bbs_table = 'bbs_" & BBSID & "'"

				Set RecSet = db.Execute(SQLs)
				If RecSet.EOF = True Then
					%>その掲示板はありません。<br>
		掲示板ID：<b><%=Replace(RecSet("bbs_table"), "bbs_", "") %></b><br><%

				Else


					'掲示板削除処理

					'テーブル削除

					SQL = "DROP TABLE bbs_" & BBSID
					db.Execute(SQL)

					'設定テーブルの設定を削除
					SQL = "DELETE FROM settings WHERE bbs_table = 'bbs_" & BBSID & "'"

					db.Execute(SQL)
					Session("BBSQuery") = "default"

					If Fso.FileExists(Server.MapPath("./count/bbscnt_" & BBSID & ".dat")) = True Then
						Fso.DeleteFile(Server.MapPath("./count/bbscnt_" & BBSID & ".dat"))
					End If

		%>ID: <b><%=BBSID %></b> 掲示板を削除しました。<br>
		<a href="admin.asp" target="_top">[BBS Administrator]</a><%

				End If
			End If

		End If

End Select

End If

Else
	'ログインされていない
	Response.Redirect "admin.asp"
End If
%>
