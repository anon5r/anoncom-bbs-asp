<!-- #Include file="bbsdb.asp" -->
<%
'###########################################
'####					####
'####	     a n o n c o m B B S	####
'####		ver. 1.8		####
'###########################################
'				by anoncom.net

bbs_ver = "1.8"


'データベース接続
Set db = Server.CreateObject("ADODB.Connection")

db.Provider = "Microsoft.Jet.OLEDB.4.0"
db.Mode = 3
db.ConnectionString = BBSDBFileName
db.Open

If Request("bbs") = "" Then
	If Request("mode") = "" Then
		BBSQuery = Request.QueryString	'bbs.asp?default用
	Else
		BBSQuery = Session("bbsquery")
	End If
Else
	BBSQuery = Request("bbs")	'bbs.asp?bbs=default用
End If

'クエリでのbbs指定がない場合はdefaultを指定
If BBSQuery = "" Then
	If Session("bbsquery") = "" Then
		BBSQuery = "default"
	Else
		BBSQuery = Session("bbsquery")
	End If
End If

'BBS Setting の読み込み
Set rs_set = Server.CreateObject("ADODB.Recordset")
rs_set.Open "SELECT * FROM settings WHERE bbs_table = 'bbs_" & _
		BBSQuery & "'", db, 3, 2

If rs_set.EOF = True Then

'bbsが存在しない場合
	ScriptNameNum = InStrRev(Request.ServerVariables("SCRIPT_NAME"), "/", -1)
	ScriptNameTemp = Mid(Request.ServerVariables("SCRIPT_NAME"), 1, CInt(ScriptNameNum))
	ScriptName = 	Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptNameTemp, "")

	BBSBlank = True
	SiteName = "anoncomBBS"
	SiteURL = "http://" & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptName, "")
	BBSName = "BBS Not Found"
	BBSComment = "そんな掲示板は存在しないです。。。"
	BBSURL = "http://" & Request.ServerVariables("HTTP_HOST") & Replace(Request.ServerVariables("SCRIPT_NAME"), ScriptName, "")
	BBSStatus = 3
	BGColor = "#ffffff"
	TextColor = "#000000"
	LinkColor = "#0000ff"
	ActiveLinkColor = "#ff0000"
	HoverLinkcolor = ActiveLinkColor
	BorderColor = "#888888"
	TitleColor = "#ff0000"
	CntNum = "10"
	CountFileName = "blankcnt.dat"
	TagUse = 0
	TagSourceView = 0
	BBSMailSend = 0
	MailServer = "mail.example.com"
	SendToAddr = "you@example.com"
	MailFromAddr = "bbs@example.com"
	MailBBSBodyCut = 0
	NotFoundName = "名無しさん"
	DeleteMailAddr = "あぼーん"
	DeleteName = "あぼーん"
	DeleteTilte = "あぼーん"
	DeleteBody = "あぼーん"
	DeleteDeviceType = "あぼーん"
	DelTitleColor = "#0000ff"
	DelBodyColor = "#000000"
	
	'ユーザ管理部分
	UserGroup = "Admin"

	rs_set.Close

Else

'bbsが存在する場合
	BBSBlank = False
	SiteName = rs_set("SiteName")
	SiteURL = rs_set("SiteURL")
	BBSName = rs_set("BBSName")
	BBSComment = rs_set("BBSComment")
	BBSURL = rs_set("BaseURL")
	BBSStatus = rs_set("act_flag")
	debug_flag = rs_set("debug_flag")
	BGColor = rs_set("BGColor")
	TextColor = rs_set("TextColor")
	LinkColor = rs_set("LinkColor")
	ActiveLinkColor = rs_set("aLinkColor")
	HoverLinkcolor = ActiveLinkColor
	BorderColor = rs_set("BorderColor")
	TitleColor = rs_set("TitleColor")
	CntNum = rs_set("ViewCount")
	CountFileName = rs_set("CountFile")

	If rs_set("Tag") = True Then
		TagUse = 1
	Else
		TagUse = 0
	End If

	If rs_set("TagSourceView") = True Then
		TagSourceView = 1
	Else
		TagSourceView = 0
	End If

	If rs_set("MailSend") = True Then
		BBSMailSend = 1
	Else
		BBSMailSend = 0
	End If


	MailServer = rs_set("MailServer")
	SendToAddr = rs_set("SendToAddr")
	MailFromAddr = rs_set("MailFromAddr")

	If rs_set("MailBBSBodyCut") = True Then
		MailBBSBodyCut = 1
	Else
		MailBBSBodyCut = 0
	End If

	NotFoundName = rs_set("NotFoundName")


	DeleteMailAddr = rs_set("DelMailAddr")
	DeleteName = rs_set("DelName")
	DeleteTilte = rs_set("DelTitle")
	DeleteBody = rs_set("DelBody")
	DeleteDeviceType = rs_set("DelDevType")

	DelTitleColor = rs_set("DelTitleColor")
	DelBodyColor = rs_set("DelBodyColor")

	'ユーザ管理部分
	If IsNull(rs_set("groups")) = True Or rs_set("groups") = "" Then
		UserGroup = "all_user"
	Else
		UserGroup = rs_set("groups")
	End If

	rs_set.Close

End If

Set rs_set = Nothing

%>
