<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" --><%

'バッファ有効
Response.Buffer = True
%>
<html lang="ja">
<head>
<meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Content-Style-Type" content="text/html; charset=shift_jis">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<title>Admin ユーザ管理</title>
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


'携帯の場合は携帯用の管理ページへ移動
If BrowserType = "Mobile" Then
	Response.Redirect BBSURL & "admink.asp"
End If

'ログインセッションは確立しているか
If Session("login") = 1 Then





	If Session("AdminLevel") < 9 Then
%>ユーザ管理権限がありません。<%
	Else


		If Request("mode") = "" Then


		'ユーザ管理画面
%>
<form action="admusrmng.asp" method="post">
<input type="hidden" name="mode" value="add">
<b>[ユーザの追加]</b><br>
管理者ID：
<input type="text" name="uid" value="" maxlength="20">
<input type="submit" value="アカウント作成">　※半角英数字最大20文字<br>
管理者名：
<input type="text" name="uname" value="" maxlength="50">
　※最大50文字<br>
パスワード：
<input type="password" name="pwd" value="" maxlength="20">
　※半角英数字最大20文字<br>
管理者ランク：
<select name="adminlevel">
<option value="8">準最高権限</option>
<option value="7">掲示板管理+板作成</option>
<option value="6" selected>掲示板管理者</option>
<option value="5">（未割り当て）</option>
<option value="4">準掲示板管理者</option>
<option value="3">一般管理者ユーザ</option>
<option value="2">削除人</option>
<option value="1">ログ閲覧のみ</option>
<option value="0">アカウント停止</option>
</select>
　<a href="adminranklist.html">管理者ランク権限一覧</a><br>
　※ランク9(最高権限:root)は指定できません。<br>
管理対象掲示板：<br>
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
<b>[ユーザ情報の変更]</b><br>
管理者名：<select name="adminID"><%
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
<input type="submit" value="管理者情報変更">
</form>
<%



		Else
			Select Case Request("mode")

				Case "add"
					'ポスト後作業
					If Request.Form <> "" Then

						If Request.Form("uid") = "" Then
							%>IDが入力されていません。<%
						ElseIf Request.Form("pwd") = "" Then
							%>パスワードが入力されていません。<%
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
									%>管理者：<%=uid %> を追加できませんでした。<%
								Else
								%>
管理者：<%=rs("adminID") %> を追加しました。<br>
<br>
管理者ID：<%=rs("adminID") %><br>
名前：<%=rs("adminName") %><br>
パスワード：<%=String(Len(rs("adminPass")), "*") %><br>
管理者ランク：<%=rs("adminRank") %><br>
管理対象掲示板：<%=rs("adminBBS") %><br>
<%
								End If
							Else
								%>その管理者IDは既に存在します。<%
							End If
						End If
					End If

				Case "edit"
					If Request.Form = "" Then
						'ポスト前

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
<td align="center" bgcolor="#88ff88">設定項目</td>
<td align="center" bgcolor="#ffaaaa">設定値</td>
<td align="center" bgcolor="#aaaaff">説明・備考</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">管理者ID：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminID" value="<%=tmpAdminID %>" maxlength="20" size="20"></td>
<td align="left" bgcolor="#ccccff">管理者ログイン用ID</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">パスワード：</td>
<td align="left" bgcolor="#ffcccc"><input type="password" name="adminPass" value="<%=tmpAdminPass %>" maxlength="10" size="16"></td>
<td align="left" bgcolor="#ccccff">管理用パスワード</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">管理者名：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminName" value="<%=tmpAdminName %>" maxlength="20" size="30"></td>
<td align="left" bgcolor="#ccccff">管理者名</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">メールアドレス：</td>
<td align="left" bgcolor="#ffcccc"><input type="text" name="adminMail" value="<%=tmpAdminMail %>" maxlength="255" size="40"></td>
<td align="left" bgcolor="#ccccff">管理者メールアドレス</td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">管理者ランク：</td>
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
<option value="9"<%=tmpRank9 %>>最高権限</option>
<option value="8"<%=tmpRank8 %>>準最高権限</option>
<option value="7"<%=tmpRank7 %>>掲示板管理+板作成</option>
<option value="6"<%=tmpRank6 %>>掲示板管理者</option>
<option value="5"<%=tmpRank5 %>>（未割り当て）</option>
<option value="4"<%=tmpRank4 %>>準掲示板管理者</option>
<option value="3"<%=tmpRank3 %>>一般管理者ユーザ</option>
<option value="2"<%=tmpRank2 %>>削除人</option>
<option value="1"<%=tmpRank1 %>>ログ閲覧のみ</option>
<option value="0"<%=tmpRank0 %>>アカウント停止</option>
</select>
</td>
<td align="left" bgcolor="#ccccff">管理者権限　※<a href="adminranklist.html">権限一覧</a></td>
</tr><tr>
<td align="right" bgcolor="#ccffcc">管理対象掲示板：</td>
<td align="left" bgcolor="#ffcccc">
<select name="adminBBS" multiple><%
						SQL = "SELECT [bbs_table], [BBSName] FROM [settings] " & _
							"WHERE [bbs_table] Like 'bbs_%' " & _
							"ORDER BY [SERIAL] ASC"
						Set rs = db.Execute(SQL)

						Do While Not rs.EOF = True

							'掲示板ロード
							bbskey = Replace(rs("bbs_table"), "bbs_", "")
							bbsname = rs("BBSName")

							'管理対象掲示板呼び出し
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
<td align="left" bgcolor="#ccccff">管理可能な掲示板</td>
</tr>
</table>
<br>
<input type="submit" value="　変　　更　">
</form>
<%
					Else
						'ポスト後
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

							'設定の反映
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
<%=Request.Form("adminName") %>（<%=Request.Form("adminID") %>）の情報を更新しました。<br>
<a href="admin.asp" target="_top">[掲示板管理画面]</a><%
						Else
							%>その管理者は見つからない為、更新できませんでした。<%
						End If
					End If

				Case Else
						%>その機能は作成中です。<%
			End Select
		End If
	End If
Else
	'ログインされていない
	Response.Redirect "admin.asp"
End If
%>
</body>
</html>