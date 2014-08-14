<% @Language = "VBScript" %>
<!-- #Include file="config.asp" -->
<%
If Session("login") <> 1 Then
	Response.Redirect "admin.asp"
End If

If Request.Form = "" Then
	Response.Redirect "bbsadmin.asp?main"
Else

'初期化実行


'データベース接続
Set db = Server.CreateObject("ADODB.Connection")

db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Mode = 3
db.ConnectionString = BBSDBFileName
db.Open


'BBS テーブルを消去

SQL = "DROP TABLE bbs_" & BBSQuery

db.Execute SQL


'テーブルを新規作成

SQL = "CREATE TABLE bbs_" & BBSQuery & " (" & _
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
db.Execute SQL

db.Close


'書き込み内容の初期化処理終了


If Request.Form("cntreset") = "on" Then
	'カウンタのリセット処理
	Set Fso = Server.CreateObject("Scripting.FileSystemObject")
	Set Txt = Fso.OpenTextFile(Server.MapPath("./count/bbscnt_" & BBSQuery & ".dat"), 2)
	Txt.Write "0"
	Txt.Close
	Set Txt = Nothing
	Set Fso = Nothing
End If


%>
<b>掲示板の書き込み内容の初期化処理が完了しました。</b><br>
<br><%

End If
%>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<font align="center" size="-2">&copy;2004 anoncom.net</font>
</body>
</html>