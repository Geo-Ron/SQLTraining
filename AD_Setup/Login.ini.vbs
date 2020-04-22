if debugflag=False then On Error Resume Next
'On Error Resume Next
'**************************************************************
'Hieronder komt het bewerkbare gedeelte
' Welkomsttekst eerst,
' daarna definitie KABI gebruikers,
' daarna de groepen en share's
' Als laatste de mappings die altijd gelegd worden
'**************************************************************

'**************************************************************
'Bedrijfsspecifieke gegevens en
'tekstregels in welkomsttekst.
'**************************************************************
VDLCompany = "Ronsulting.net"
Tekstregel1 = "Your Service Provider is reachable by: " & VbCrLf & _ 
				"     Email at: servicedesk@ronsulting.net or"& VbCrLf & "     Phone at: 0123456789"
'Call SetTMGClient("Enable","EnableBrowserConfig")
'**************************************************************




'************************************************
'Groepen vs Mappings Template
'************************************************
	'Gebruiker lid van Domain Admin group 
	'wscript.echo isMember("Domain Admins") & " Is lid van domain admins"
	'wscript.echo Timer()
	'If isMember("Domain Admins") = True  Then ' -1 of True op XP werkt wel IPV "True"
	'	'Syntax: Call KoppelShare("Driveletter:","\\Hostname\Sharename$","Useraccount","Password","Domain")
	'End If
'************************************************

'************************************************
':::::::::::: Start EVERYONE Settings :::::::::::
'************************************************
If isMember("Domain Admins") = True  Then ' -1 of True op XP werkt wel IPV "True"
	Call KoppelShare("S:","\\SPTEST-az-ADC01\C$\_ICT-Beheer","","","")
End If
'************************************************
'::::::::::::: End EVERYONE Settings ::::::::::::
'************************************************

