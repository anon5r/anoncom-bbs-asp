<%

'####################################################
'#####						#####
'#####		Device Type Checker		#####
'#####						#####
'####################################################
'					ver.1.2
' 2003/12/11 UpDate
'				Created by ���̂�
'Copyright(C) 2003 anoncom.net
'Support => http://anoncom.net/script/




'���K�\����p����USER AGENT�ɂ��y�[�W�U�蕪��

'User Agent�����o��
AgentType = Request.ServerVariables("HTTP_USER_AGENT")

'���K�\��
Set RegObj = New RegExp

	RegObj.Global = True
	RegObj.IgnoreCase = False

	'DoCoMo i-mode
	RegObj.Pattern = "^DoCoMo\/[0-9]\.[0-9]\/[A-Z]{1,}[0-9]{3}(i|iS)+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mobile"
		MobileType = "i-mode"
		Provider = "DoCoMo"
	End If

	'DoCoMo FOMA
	RegObj.Pattern = "^DoCoMo\/[0-9]\.[0-9]\s[A-Z]{1,}[0-9]+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mobile"
		MobileType = "FOMA"
		Provider = "DoCoMo"
	End If

	'Vodafone Vodafone live!
	RegObj.Pattern = "^J\-PHONE\/[0-9]\.[0-9]\/\w+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mobile"
		MobileType = "Vodafone live!"
		Provider = "Vodafone"
	End If

	'au EZweb ���[��
	RegObj.Pattern = "^UP\.Browser\/\w+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mobile"
		MobileType = "EzWeb"
		Provider = "au"
	End If

	'au EZweb WAP2.0 �Ή��[��
	RegObj.Pattern = "^KDDI\-[A-Za-z]{1,2}[0-9]{2,}\sUP\.Browser\/+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mobile"
		MobileType = "EzWeb"
		Provider = "au"
	End If

	'H"�[��
	RegObj.Pattern = "PDXGW\/[0-9]\.[0-9]\s\(+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mobile"
		MobileType = "H"
		Provider = "DDIPocket"
	End If

	'PC Mozilla
	RegObj.Pattern = "^Mozilla\/[0-9]\.[0-9]+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Mozilla"
		MobileType = ""
		Provider = "PC"
	End If

	'PC Mozilla
	RegObj.Pattern = "^Opera\/[0-9]\.[0-9]+"
	If RegObj.Test(AgentType) = True Then
		BrowserType = "Opera"
		MobileType = ""
		Provider = "PC"
	End If

%>
