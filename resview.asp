<% @Language="VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="removehtml.asp" -->
<%

If CInt(BBSStatus) < 3 Then
	Response.Redirect "forbidden.html"
ElseIf CInt(BBSStatus) < 2 Then
	Response.Redirect "closebbs.html"
End If


If BBSBlank = True Then
	Response.Redirect "nobbs.html"
End If

If Request.QueryString("bbs") = "" Then
	BBSQuery = "default"
Else
	BBSQuery = Request.QueryString("bbs")
End If

'	************UAによって表示の仕方を少し変える************

%>
<!-- #Include file="devtype.asp" -->
<html>
<head>
<title><%=BBSName %> - anoncomBBS</title>
<% If MobileType = "" Then %>
<meta http-equiv="Content-Type" content="text/html; charset=Shift_JIS">
<style type="text/css">
<!--
a:link { text-decoration:none; color:<%=LinkColor %>; }
a:visited { text-decoration:none; color:<%=LinkColor %>; }
a:active { text-decoration:none; color:<%=ActiveLinkColor %>; }
a:hover { text-decoration:underline; color:#<%=HoverLinkColor %>; cursor:default; }

body{
	font-size:10pt;
	font-family:'MS UI Gothic';
	overflow-y:auto;
	scrollbar-base-color:<%=BGColor %>;
	scrollbar-face-color:<%=BGColor %>;
	scrollbar-arrow-color:<%=BorderColor %>;
	scrollbar-highlight-color:<%=BGColor %>;
	scrollbar-3dlight-color:<%=BorderColor %>;
	scrollbar-shadow-color:<%=BorderColor %>;
	scrollbar-darkshadow-color:<%=BGColor %>;
}

input{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic';
	background-color:<%=BGColor %>;
	border-style:solid;
	border-width:1px;
	border-color:<%=BorderColor %>;
}

-->
</style>
<% End If %>
</head>
<body bgcolor="<%=BGColor %>" text="<%=TextColor %>" link="#ff0000" vlink="#ff0000">
<font size="5" color="<%=TitleColor %>"><i><%=BBSName %></i></font>
<div align="right"><font size="2">by anoncomBBS</font></div>
<br>
<br>
<center><%=BBSComment %></center>
<br>
<hr size="1">
<%
If Request.QueryString("no") = "" Then
'レス番号指定なし
%>
0:■[anoncomBBS]<br>
[No Script]<br>
<br>書き込みがないか、レス番号が不正です。
<div align="right">[system]<br>
[<%=WriteTime(Now) %>]</div><hr>
<%
Else
'レス番号の指定あり

	ResNo = Request.QueryString("no")
	'レス番号がハイフンで区切られているか判定
	'Split("対象文字列", "区切り文字列(省略で半角スペース)", 返す配列の要素数(-1で全て), 評価方法)
	aryResNo = Split(ResNo, "-")

	If UBound(aryResNo) >= 1 Then	'配列のインデックスが1以上あるか
		'ある場合

		If aryResNo(0) = "" Then

			'指定表示件数を上回らないか
			If aryResNo(1) - 1 >= CntNum Then
				'表示中断フラグ
				ExitView = True
			End If

			SQL = "SELECT * FROM bbs_" & BBSQuery & " WHERE Num " & _
				"BETWEEN 1 AND " & aryResNo(1) & _
				" ORDER BY Num DESC"

		ElseIf aryResNo(1) = "" Then

			'総レス数をカウント
			CntSQL = "SELECT count(*) As RsCnt FROM bbs_" & BBSQuery
			Set ResCount = db.Execute(CntSQL)
			MaxCnt = ResCount("RsCnt")

			'指定表示件数を上回らないか
			If MaxCnt - aryResNo(0) >= CntNum Then
				'表示中断フラグ
				ExitView = True
			End If

			SQL = "SELECT * FROM bbs_" & BBSQuery & " WHERE Num " & _
				"BETWEEN " & aryResNo(0) & " AND " & MaxCnt & _
				" ORDER BY Num DESC"
		Else

			'指定表示件数を上回らないか
			If aryResNo(0) - aryResNo(1) >= CntNum  Or aryResNo(1) - aryResNo(0) >= CntNum Then
				'表示中断フラグ
				ExitView = True
			End If
			SQL = "SELECT * FROM bbs_" & BBSQuery & " WHERE Num " & _
				"BETWEEN " & aryResNo(0) & " AND " & aryResNo(1) & _
				" ORDER BY Num DESC"
		End If
	Else
		'ない場合
		SQL = "SELECT * FROM bbs_" & BBSQuery & " WHERE Num=" & ResNo
	End If

	Set rs = db.Execute(SQL)

	If rs.EOF = True Then

		'レスがない場合
%>
0:■[anoncomBBS]<br>
[No Script]<br>
<br>書き込みがないか、レス番号が不正です。
<div align="right">[system]<br>
[<%=WriteTime(Now) %>]</div><hr>
<%
	Else

		'1ページの指定表示件数が多くないか？
		If ExitView = True Then
			%>表示件数が多すぎます！<br><br><%
		Else


		Do While Not rs.EOF = True

			If rs("abone") = "True" Then
				If Request.QueryString("del") <> "view" Then
%>
<%=rs("Num") %>:■[<a href="mailto:<%=DeleteMailAddr %>"><%=DeleteName %></a>]<br>
[<font color="#FF99AA"><%=DeleteTilte %></font>]<br>
<br><font color="#ff0000"><%=DeleteBody %></font><br><br>
<div align="right">[<%=DeleteDeviceType %>]<br>
[<%=WriteTime(rs("sdat")) %>]</div>
<hr size="1">
<%
				Else
					'削除レスも強制表示

					If rs("from")<>"" Then
						If rs("mail")<>"" Then
%>
<%=rs("Num") %>:■[<a href="mailto:<%=rs("mail") %>"><%=rs("from") %></a>]<br><%
						Else %>
<%=rs("Num") %>:■[<%=rs("from") %>]<br><%
						End If
					Else
						If rs("mail")<>"" Then
%>
<%=rs("Num") %>:■[<a href="mailto:<%=rs("mail") %>"><%=rs("mail") %></a>]<br><%
						Else %>
<%=rs("Num") %>:■[<%=NotFoundName %>]<br>
<%
						End If
					End If

					If rs("title")<>"" Then
%>
[<%=rs("title") %>]<br><%
					End If %>
<br><%=rs("message") %><br><br><%
					If rs("url") <> "" Then
						If rs("url") <> "http://" Then
%>
<div align="right"><a href="<%=rs("url") %>" target="_blank">Homepage</a></div><%
						End If
					End If
%>
<div align="right">[<%=rs("UA") %>]<br>
[<%=WriteTime(rs("sdat")) %>]</div>
<hr size="1"><%
				End If

			Else
				'あぼーん以外

				If rs("from")<>"" Then
					If rs("mail")<>"" Then
%>
<%=rs("Num") %>:■[<a href="mailto:<%=rs("mail") %>"><%=rs("from") %></a>]<br><%
					Else %>
<%=rs("Num") %>:■[<%=rs("from") %>]<br>
<%
					End If
				Else
					If rs("mail")<>"" Then
%>
<%=rs("Num") %>:■[<a href="mailto:<%=rs("mail") %>"><%=rs("mail") %></a>]<br><%
					Else %>
<%=rs("Num") %>:■[<%=NotFoundName %>]<br><%
					End If
				End If

				If rs("title")<>"" Then
%>
[<%=rs("title") %>]<br><%
				End If

				If TagUse = 1 Then
					'タグ有効時
					message = "<br>" & rs("message") & vbCrLf
					message = Replace(message, vbCrLf, "<br>" & vbCrLf)
					message = message & "<br>" & vbCrLf
				Else
					'タグ無効時

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

				        message = Replace(message, vbCrLf, "<br>" & vbCrLf)

				End If
%>
<br><%=message %><br><br><%

    If rs("url")<>"" Then
      If rs("url")<>"http://" Then
		HP_URL = "<a href=""jump.asp?url=" & rs("url")  & """ target=""_blank"">Homepage</a><br>" & vbCrLf
      Else
		HP_URL = ""
      End If
    Else
	HP_URL = ""
    End If

%>
<div align="right">
<%=HP_URL %>[<%=rs("UA") %>]<br>
[<%=WriteTime(rs("sdat")) %>]</div>
<hr size="1">
<%
			End If

			rs.MoveNext
		Loop
		End If	'ExitViewの判定

'書き込み日付処理
Function WriteTime(dtmNow)

Dim strDate
strDate = Right(String(4,"0") & Year(dtmNow),4) & "/" & Right(String(2,"0") & Month(dtmNow),2) & "/" & Right(String(2,"0") & Day(dtmNow),2) & " " & Right(String(2,"0") & Hour(dtmNow),2) & ":" & Right(String(2,"0") & Minute(dtmNow),2)
WriteTime = strDate

End Function
'書き込み日付処理終了

	End If

End If

If CInt(BBSStatus) > 3 Then
%>
<a href="res.asp?bbs=<%=BBSQuery %>">書く</a><br><%
End If
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>">[掲示板]</a><br>
<hr size="1">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
</body>
</html>
