<%@Language="VBScript" %>
<!-- #Include file="config.asp" -->
<%
If Request.QueryString = "" Then
	Response.Redirect BBSURL & "bbs.asp"
ElseIf Request.QueryString("bbs") <> "" Then
	Response.Redirect BBSURL & "bbs.asp?" & Request.QueryString
Else
	Response.Redirect BBSURL & "bbs.asp?bbs=" & Request.QueryString
End If
%>
