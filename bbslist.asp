<%@ Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<!-- #Include file="devtype.asp" -->
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

table#bbs{
	border-style:solid;
	border-color:<%=BorderColor %>;
	border-width:1px;
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

/*
td#bottom{
	color:<%=TextColor %>;
	font-size:10pt;
	font-family:'MS UI Gothic','ＭＳゴシック';
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
'PCから見る場合のみテーブルを使用して表示
Select Case BrowserType

   Case "Mobile"
etd = "<br>"
etd_bottom = "<br>" & vbCrLf & "<hr>"

   Case Else
 %>
<center>
<table border="0" width="75%" cellspacing="0">
<%
'変数を使用してテーブルを描写
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

'PCから見る場合のみテーブルを使用して表示
Select Case BrowserType
   Case "Mobile"
etd = "<br>"
   Case Else
 %>
<table border="0" width="75%" id="bbs">
<%

End Select



'	掲示板ページのカウント
If Request.QueryString("cnt") = "" Then
  cnt = 1
  page = 1
Else
  cnt = CInt(Request.QueryString("cnt"))
  tmpCnt = cnt - CntNum
  page = CInt(Request.QueryString("page"))
End If


Set rs = Server.CreateObject("ADODB.Recordset")


rs.Open "SELECT * FROM settings WHERE [bbs_table] Like 'bbs_%' AND [act_flag] > 0 ORDER BY [SERIAL] ASC", db, 3, 2


If rs.EOF = True Then

'掲示板がない場合

ResCnt = "0"

%>
<%=tr %>
<%=td %>現在有効な掲示板がありません。<%=etd %>
<%=etr %>
<%
Else

'掲示板がある場合
ResCnt = rs.RecordCount
rs.AbsolutePosition = cnt

Do While Not rs.EOF
%><%=tr %>
<%=td_name %><a href="<%=rs("BaseURL") %>?<%=Replace(rs("bbs_table"), "bbs_", "") %>"><%=rs("BBSName") %></a><%=etd %>
<%=etr %>
<%
  rs.MoveNext
  cnt = cnt + 1
  If cnt = page * CntNum + 1 Then Exit Do
Loop

'終了(テーブルタグ閉じ)
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


End If


 Select Case Provider

'	NTT DoCoMo i-mode の場合
   Case "DoCoMo"

If page>1 Then %>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=tmpCnt %>&page=<%=page - 1 %>">&lt;戻る</a>/
<%
End If
If Not rs.EOF Then
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page + 1 %>">次へ&gt;</a>
<% End If %><br>
<hr color="<%=BorderColor %>">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
<%
'	Vodafone Vodafone live! の場合
   Case "Vodafone"

If page > 1 Then %>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=tmpCnt %>&page=<%=page - 1 %>">$F[戻る</a>/
<%
End If
If Not rs.EOF Then
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page + 1 %>">次へ$FZ</a>
<% End If %><br>
<hr color="<%=BorderColor %>">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
<%
'	au Ez-web の場合
   Case "au"

If page>1 Then %>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=tmpCnt %>&page=<%=page - 1 %>">&lt;戻る</a>/
<%
End If
If Not rs.EOF Then
%>
<a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page + 1 %>">次へ&gt;</a>
<% End If %><br>
<hr color="<%=BorderColor %>">
system by<br>
<a href="http://anoncom.net/">anoncom.net</a>
<%
'	パソコン、その他からの場合
   Case Else

If Not rs.EOF Then
%>
<td align="left" valign="top">
<span class="button"><a href="bbs.asp?bbs=<%=BBSQuery %>&cnt=<%=cnt %>&page=<%=page+1 %>">次へ</a></span>
</td>
</tr><tr>
<% End If %>
<td align="left">&nbsp;</td>
</tr><tr>
<td align="right">
<i>system by <a href="http://anoncom.net/" target="_top">anoncom.net</a></i>
</td>
</tr>
</table>
<%
End Select
'	************UAによる表示形式の変更終了***************


%>
</body>
</html>
