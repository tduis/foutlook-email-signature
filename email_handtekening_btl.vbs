'--------------------------------------------------------------------------------------------------------------------------------------------------
'	Email_handtekening_btl.vbs
'--------------------------------------------------------------------------------------------------------------------------------------------------
' Door: T.Duis (E-genius)
'
'Datum		Wijzigingen
' 19-06-2008	(TD) Eerste werkende versie, getest met Meggie, Jeroen, Leon, Clemens
' 20-06-2008	(TD) Controle of Outlook al gestart is uitgeschakeld ivm Terminalserver gebruik
' 01-09-2008	(TD)Fout gevonden bij lege AD achternaam, opgelost
'			(TD)Fout voor text handtekening vbCrLf gebruikt ivp vbCr
' 03-09-2008	(TD) Extra groep controle Sec_Realisatie_Hoofdkantoor ingebouwd
' 24-10-2008	(TD) Script omgebouwd zodat: AD Default groep als basis email handtekening wordt ingesteld
' 28-10-2008	(TD) Groepen Sec_AdviesStein en Sec_Oisterwijk toegevoegd.
' 28-10-2008	(TD) RGB kleuren aanpassing voor 'denk aan milieu' regel
'			BTL: 179,213,151 (#b3d597) en Tuinstijl: 127,187,86 (#7fbb56)
' 02-03-2009	sec_boogaart toegevoegd
' 09-03-2009	Voor default groep onderverdeling voor FiveStar Grass gemaakt. Daarnaast als prigroup "Domain Users" is,
'			dan laatste groep als default instellen.
'			tuinstijl handtekeningen mogelijk gemaakt middels groepen tuinstijl_maarssen, tuinstijl_bruinisse, tuinstijl_roermond
'			tuinstijl_stein, tuinstijl_veldhoven, tuinstijl_hengelo, tuinstijl_roosendaal, tuinstijl_joure, tuinstijl_haaren
'			idem voor regio_directeur_zuid, regio_directeur_noord-oost, regio_directeur_midden-west.


Set objArgs = WScript.Arguments
Set WshShell = WScript.CreateObject("WScript.Shell") 
Set fso = CreateObject("Scripting.FileSystemObject") 

Dim aQuote
Dim theUser, Domain
aQuote = chr(34)

Domain = NTDomain ' Need NT domain for WinNT:
strPrimaryGroup = PrimaryGroup(objArgs(0))
Wscript.Echo "Primary Group: " & strPrimaryGroup


' Instellen pad voor handtekening bestand
'strUserProfile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")
'strWinDir = WshShell.ExpandEnvironmentStrings("%SYSTEMROOT%") 
'sig_pad = strUserProfile & "\Application Data\Microsoft\Signatures\"
strAppdata = WshShell.ExpandEnvironmentStrings("%APPDATA%")
sig_pad = strAppData & "\Microsoft\Handtekeningen\"

' Extra registry key schrijven indien deze niet bestaat
strWinLogon = "HKLM\SOFTWARE\Microsoft\Windows NT\currentVersion\Winlogon\"

' This section checks if the signature directory exits and if not creates one.
'====================
Dim objFS1
Set objFS1 = CreateObject("Scripting.FileSystemObject")
If (objFS1.FolderExists(sig_pad)) Then
Else
	Call objFS1.CreateFolder(sig_pad)
End if


milieu_nl="Denk aan het milieu. Print deze e-mail alleen als het noodzakelijk is."
milieu_uk="Please think about the environment. Only print this email if necessary."
disclaimer_nl="Aan dit bericht kunnen geen rechten worden ontleend. Dit bericht is alleen bestemd voor de geadresseerde. Indien dit bericht niet voor u bestemd is, verzoeken wij u vriendelijk dit onmiddellijk aan ons te melden en de inhoud van dit bericht te vernietigen. De afzender controleert al haar uitgaande e-mails op aanwezigheid van virussen. Wij zijn niet aansprakelijk voor en/of in verband met alle mogelijke gevolgen en/of schade voortvloeiende uit dit e-mailbericht, zoals schade door virussen." 
disclaimer_uk="No rights can be claimed from this electronic message. This e-mail message is private and confidential, and only intended for the addressee(s). If this e-mail was sent to you by mistake, would you please contact us immediately. In that case, we also request that you destroy the e-mail. This e-mail is scanned for all viruses known. We deny any responsibility for damages resulting from the use of this e-mail message."
groet_nl="Met vriendelijke groet,"
groet_uk="With kind regards,"
groet_de="Mit freundlichen Grüssen"
groet_fsg="Met vriendelijke groet,/ With kind regards,/ Mit freundlichen Grüssen,"
'Grüßen
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection

objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.Properties("Chase referrals") = 64 'ADS_CHASE_REFERRALS_EXTERNAL 
objCommand.Properties("Asynchronous") = True 
objCommand.Properties("Cache results") = False 

tmp="SELECT cn,Name,sAMAccountName,GivenName,Sn,displayName,mail,mailNickname,memberOf,mobile,description,department,title FROM 'LDAP://dc=btlgroep,dc=domein' WHERE objectCategory='user' " & _
        "AND sAMAccountName='" & objArgs(0) & "'"

objCommand.CommandText = tmp
Set objRecordSet = objCommand.Execute

objRecordSet.MoveFirst
Do Until objRecordSet.EOF
	'Wscript.Echo objRecordSet.Fields("cn").Value
    'Wscript.Echo "-------------------------------------------------------:"
	voornaam=objRecordSet.Fields("GivenName").Value
	' Extra controle omdat gebleken is dat als de voornaam niet is ingevult, het script niet loopt
	if IsNull(voornaam) then
		voornaam=" "
	End If
	voornaam=UCase(Left((voornaam),1)) & LCase(Right((voornaam), Len((voornaam))-1))
    'Wscript.Echo "Voornaam:" & objRecordSet.Fields("GivenName").Value & vbCr
	achternaam=objRecordSet.Fields("Sn").Value
	' Extra controle omdat gebleken is dat als de achternaam niet is ingevult, het script niet loopt
	if IsNull(achternaam) then
		achternaam=" "
	End If
	achternaam=UCase(Left((achternaam),1)) & LCase(Right((achternaam), Len((achternaam))-1))
	displaynaam=objRecordSet.Fields("displayName").Value
    'Wscript.Echo "Achternaam:" & objRecordSet.Fields("Sn").Value & vbCr
    'Wscript.Echo objRecordSet.Fields("Name").Value & vbCr
    'Wscript.Echo "Loginnaam:" & objRecordSet.Fields("sAMAccountName").Value & vbCr
    'Wscript.Echo "E-mail: " & objRecordSet.Fields("mail").Value & vbCr
	email=objRecordSet.Fields("mail").Value
    'Wscript.Echo "Mail Alias: " & objRecordSet.Fields("mailNickname").Value & vbCr
    'Wscript.Echo "Mobiel: " & objRecordSet.Fields("mobile").Value & vbCr
	mobiel=objRecordSet.Fields("mobile").Value
    'Wscript.Echo "Afdeling: " & objRecordSet.Fields("department").Value & vbCr
	kantoor=objRecordSet.Fields("department").Value
    'Wscript.Echo "Functie: " & objRecordSet.Fields("title").Value & vbCr
	functie=objRecordSet.Fields("title").Value

	'Nu kunnen we de primaire handtekening gaan maken indien de primarygroep geen "Domain Users" is.
	If LCase(strPrimaryGroup)="domain users" Then
		Wscript.Echo "Primary Group: Domain Users"
	Else
		strCN="LDAP://" & "CN=" & strPrimaryGroup & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
		set objGroup = GetObject(strCN)

		'WScript.Echo objGroup.info
		info=objGroup.info
		'WScript.Echo objGroup.Description
		bedrijf=objGroup.Description
						
		'Instellen handtekeningen bestandsnaam naar loginnaam gebruiker
		sig_file=objRecordSet.Fields("sAMAccountName").Value
		
		'Onderscheid maken voor subbase / fivestar grass
		If LCase(strPrimaryGroup)="sec_subbase" then
			'Email handtekening Fivestar Grass plaatsen
			call WriteHTMLFiveStarGrass(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_fsg)

			' Wegschrijven van TXT handtekening
			call WriteTXTFiveStarGrass(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_fsg)
		
		Else
			' Wegschrijven van HTML handtekening
			call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

			' Wegschrijven van TXT handtekening
			call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

		End If

		'Instellen van handtekening voor nieuwe berichten & antwoorden en doorsturen
		Call SetDefaultSignature(sig_file,"")  'this appears in the outlook file as the signature name, change it to desired
						
	End If

	'Nu controleren van welke groepen deze gebruiker lid is
	objmemberOf  = objRecordSet.Fields("memberOf")
	

	
	'Loop door alle groepen en bij juiste Sec_xxxx groepen instellingen ophalen
	For Each strGroep in objmemberOf

			'Nu op comma splitten (CN=Sec_Hoofdkantoor,OU=02 Security Groepen,dc=btlgroep,dc=domein)
			tmp = Split(strGroep, ",") 
			'Vervolgens op = splitsen
			strGrpName = Split(tmp(0), "=")
			
			'WScript.Echo "Groep: " & strGrpName(1) 
			Select case LCase(strGrpName(1))
				Case "sec_realisatie_hoofdkantoor"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					'WScript.Echo "\n" & strCN
					set objGroup = GetObject(strCN)

					'WScript.Echo objGroup.info
					info=objGroup.info
					'WScript.Echo objGroup.Description
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
					
					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If					
				Case "sec_hoofdkantoor"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					'WScript.Echo strCN
					set objGroup = GetObject(strCN)
					'WScript.Echo objGroup.whenCreated
					'WScript.Echo objGroup.info
					info=objGroup.info
					'WScript.Echo objGroup.Description
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_bedrijsbureau"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					'Script.Echo strCN
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_verhoeven"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					'Uitlezen van 
					info=objGroup.info					
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_haaren"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					'Uitlezen van 
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
					
					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_arnhem"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					'Uitlezen van 
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_bomendienst"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
					
					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_boogaart"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
					
					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_bruinisse"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_eindhoven"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

									Case "sec_denoo"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_directie"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description
					WScript.Echo "Bedrijf: " & bedrijf
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_emmen"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
				
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_hsluis"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_maarssen"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_oss"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_roermond"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_roosendaal"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_stein"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_subbase"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					'Wegschrijven Fivestargrass email handtekening
					call WriteHTMLFiveStarGrass(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_uk,disclaimer_nl,disclaimer_uk,groet_fsg)


					' Wegschrijven van TXT handtekening
					call WriteTXTFiveStarGrass(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_fsg)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_trias"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description
					
					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_ulrum"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_veldhoven"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_verhoeven"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_oisterwijk"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "sec_adviesstein"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
				'-------------------------------------------------------------------------------------------------------
				' T U I N S T I J L  H A N D T E K E N I N G E N
				'-------------------------------------------------------------------------------------------------------
				Case "tuinstijl_bruinisse"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "tuinstijl_haaren"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "tuinstijl_hengelo"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "tuinstijl_joure"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "tuinstijl_maarssen"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If	
					
				Case "tuinstijl_roermond"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "tuinstijl_roosendaal"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If	
					
				Case "tuinstijl_stein"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "tuinstijl_veldhoven"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
				'-------------------------------------------------------------------------------------------------------
				' R E G I O   D I R E C T E U R E N   H A N D T E K E N I N G E N
				'-------------------------------------------------------------------------------------------------------
				Case "regio_directeur_zuid"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "regio_directeur_midden-west"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If
					
				Case "regio_directeur_noord-oost"
					' NU moeten we het info veld uitlezen van deze groep
					strCN="LDAP://" & "CN=" & strGrpName(1) & "," & "OU=02 Security Groepen" & "," & "dc=btlgroep" & "," & "dc=domein"
					set objGroup = GetObject(strCN)
					info=objGroup.info
					bedrijf=objGroup.Description

					'Instellen handtekeningen bestandsnaam naar groepsnaam
					sig_file=strGrpName(1)
					
					' Wegschrijven van HTML handtekening
					call WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)
						
					' Wegschrijven van TXT handtekening
					call WriteTXTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf, functie,info,milieu_nl,disclaimer_nl,disclaimer_uk,groet_nl)

					' Als de Primary Group   "Domain Users" is, dan deze als default handtekening instellen
					If LCase(strPrimaryGroup)="domain users" Then
						Call SetDefaultSignature(sig_file,"")
					End If	
			End Select
	Next
	'Naar volgende
    objRecordSet.MoveNext
Loop



'Sub WriteTxTSignature(byVal sigpad,byVal sigfile,bestand, voornaam, achternaam,bedrijf,functie,info)
Function WriteTxTSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf,functie,info,milieu,disclaimer_nl,disclaimer_uk,groet)
	Set signature = fso.OpenTextFile(sig_pad & sig_file & ".txt", ForWriting, True) 
	signature.WriteLine "Met vriendelijke groet," 
	signature.WriteLine bedrijf & vbCr
	signature.WriteLine vbCr
	signature.WriteLine vbCr
	signature.WriteLine displaynaam & vbCr
	signature.WriteLine functie & vbCr
	signature.WriteLine "-------------------------------------------------------"	
	signature.WriteLine bedrijf & vbCr
	'Nu moeten we info nog over meerdere regels verdelen
	tekst=Split(info,vbCrLf)
	i=1
	For Each regel in tekst
		regel=ltrim(regel)
		Select Case i
			Case 1
				signature.WriteLine "Postadres    : " & regel
			Case 2
				signature.WriteLine "Bezoekadres  : " & regel
			Case 3
				signature.WriteLine "Telefoon     : " & regel
				if (mobiel <> "") Then
					signature.WriteLine "Mobiel       : " & mobiel
				End If
			Case 4
				signature.WriteLine "Fax          : " & regel
			Case 5
				signature.WriteLine "E-mail       : " & email
			Case 6
				signature.WriteLine "Website      : " & regel
		End Select
		'wscript.echo regel + Cstr(i) + "\n"
		i=i+1 'teller ophogen
	Next
	'signature.WriteLine "" & vbCr 
	'signature.WriteLine info & vbCr
	signature.WriteLine "" & vbCr 
	signature.WriteLine milieu & vbCr 
	signature.WriteLine "" & vbCr 
	signature.WriteLine "-------------------------------------------------------" & vbCr
	signature.WriteLine disclaimer_nl & vbCr
	signature.WriteLine "" & vbCr 
	signature.WriteLine disclaimer_uk & vbCr
	signature.WriteLine "-------------------------------------------------------" & vbCr	
	signature.Close
End Function

'Sub WriteTxTFiveStarGrass(byVal sigpad,byVal sigfile,bestand, voornaam, achternaam,bedrijf,functie,info)
Function WriteTxTFiveStarGrass(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf,functie,info,milieu,disclaimer_nl,disclaimer_uk,groet)
	Set signature = fso.OpenTextFile(sig_pad & sig_file & ".txt", ForWriting, True) 
	signature.WriteLine "Met vriendelijke groet," 
	signature.WriteLine bedrijf & vbCr
	signature.WriteLine vbCr
	signature.WriteLine vbCr
	signature.WriteLine displaynaam & vbCr
	signature.WriteLine functie & vbCr
	signature.WriteLine "-------------------------------------------------------"	
	signature.WriteLine bedrijf & vbCr
	'Nu moeten we info nog over meerdere regels verdelen
	tekst=Split(info,vbCrLf)
	i=1
	For Each regel in tekst
		regel=ltrim(regel)
		Select Case i
			Case 1
				signature.WriteLine "Postal addres    : " & regel
			Case 2
				signature.WriteLine "Visiting address : " & regel
			Case 3
				signature.WriteLine "Telephone        : " & regel
				if (mobiel <> "") Then
					signature.WriteLine "Mobile           : " & mobiel
				End If
			Case 4
				signature.WriteLine "Fax              : " & regel
			Case 5
				signature.WriteLine "E-mail           : " & email
			Case 6
				signature.WriteLine "Website          : " & regel
		End Select
		'wscript.echo regel + Cstr(i) + "\n"
		i=i+1 'teller ophogen
	Next
	'signature.WriteLine "" & vbCr 
	'signature.WriteLine info & vbCr
	signature.WriteLine "" & vbCr 
	signature.WriteLine milieu & vbCr 
	signature.WriteLine "" & vbCr 
	signature.WriteLine "-------------------------------------------------------" & vbCr
	signature.WriteLine disclaimer_nl & vbCr
	signature.WriteLine "" & vbCr 
	signature.WriteLine disclaimer_uk & vbCr
	signature.WriteLine "-------------------------------------------------------" & vbCr	
	signature.Close
End Function

' Email handtekening voor Fivestargrass
Function WriteHTMLFiveStarGrass(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf,functie,info,milieu,disclaimer_nl,disclaimer_uk,groet_uk)
	Set signature = fso.OpenTextFile(sig_pad & sig_file & ".htm", ForWriting, True) 
	signature.Write "<!DOCTYPE HTML PUBLIC " & aQuote & "-//W3C//DTD HTML 4.0 Transitional//EN" & aQuote & ">" & vbCrLf
	signature.write "<HTML><HEAD><TITLE>Microsoft Office Outlook Signature</TITLE>" & vbCrLf
	signature.write "<META http-equiv=Content-Type content=" & aQuote & "text/html; charset=windows-1252" & aQuote & ">" & vbCrLf
	signature.write "<META content=" & aQuote & "MSHTML 6.00.3790.186" & aQuote & " name=GENERATOR></HEAD>" & vbCrLf
	signature.write "<BODY>" & vbCrLf
	signature.WriteLine "<p style='color: #000; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px;'>" & groet_uk & "<br>"
	signature.WriteLine bedrijf & "</p>" & vbCr
	signature.WriteLine "<br>" & vbCr
	signature.WriteLine "<p style='color: #000; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px;'>" & displaynaam & "<br>" & vbCr
	signature.WriteLine functie & vbCr
	signature.WriteLine "</p>" & vbCr
	signature.WriteLine "-------------------------------------------------------" & vbCr
	signature.WriteLine "<div style='color: #919195; font-weight:bold; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 14px;margin-top: 5px; margin-bottom: 5px;'>" & vbCr
	signature.WriteLine bedrijf & vbCr
	signature.WriteLine "</div>"& vbCr
	'signature.WriteLine "<div style='color: #00688F; font-family: Verdana, Arial, Helvetica, sans-serif;  font-size: 12px;'>" & vbCr
	'temp=Split(info,vbCr)
	'signature.WriteLine "<span style='width: 100px;'>" & temp(1) & "</span>" & vbCr
	'signature.WriteLine "<span style='width: 100px;'>" & temp(1) & "</span>" & vbCr
		'Nu moeten we info nog over meerdere regels verdelen
	tekst=Split(info,vbCrLf)
	stijl="style='color: #919195; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px; margin-top: 2px; margin-bottom: 2px;'"
	stijl1="style='color: #919195; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px; margin-left: 35px; margin-right: 3px;'"
	stijl2="style='text-decoration: none; color: #919195; font-family: Arial,Verdana, Helvetica, sans-serif;  font-size: 12px; margin-left: 35px; margin-right: 3px;'"
	stijl3="style='text-decoration: none; color: #919195; font-family: Arial,Verdana, Helvetica, sans-serif;  font-size: 12px;'"
	

	i=1
	For Each regel in tekst
		regel=ltrim(regel)
		Select Case i
			Case 1
				signature.WriteLine "<table border='0' cellpadding='0' cellspacing='0'>"
				signature.WriteLine "<tr><td " & stijl & ">Postal address</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
			Case 2
				signature.WriteLine "<tr><td " & stijl & ">Visiting address</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
			Case 3
				signature.WriteLine "<tr><td " & stijl & ">Telephone</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
				if (mobiel <> "") Then
					signature.WriteLine "<tr><td " & stijl & ">Mobile</td><td " & stijl1 & ">:</td><td " & stijl & ">" & mobiel & "</td></tr>"
				End If
			Case 4
				signature.WriteLine "<tr><td " & stijl & ">Fax</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
				signature.WriteLine "<tr><td " & stijl & ">Email</td><td " & stijl1 & ">:</td><td " & stijl & ">" & email & "</td></tr>"
			Case 5
				signature.WriteLine "<tr><td " & stijl & ">Website</td><td " & stijl1 & ">:</td><td " & stijl3 & ">" & "<a " & stijl3 & " href='http://" & regel & "'>" & regel & "</a></td></tr>"
				signature.WriteLine "</table>"
		End Select
		
		i=i+1 'teller ophogen
	Next	
	
	'signature.WriteLine info & vbCr
	signature.WriteLine "<br><div style='font-weight:bold; color: #c2cd23; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 11px;'>" & vbCr
	signature.WriteLine milieu & vbCr
	signature.WriteLine "</div>" & vbCr 
	signature.WriteLine "" & vbCr 
	signature.WriteLine "<hr>"	& vbCr
	signature.WriteLine "<div style='color: #919195; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 9px;'>" & vbCr
	signature.WriteLine disclaimer_uk & vbCr
	signature.WriteLine "</div>" & vbCr
	signature.WriteLine "<hr>"	& vbCr
	signature.WriteLine "</BODY>" & vbCr
	signature.WriteLine "</HTML>" & vbCr
	signature.Close
End Function

'Sub WriteHTMLSignature(byVal sigpad,byVal sigfile,bestand, voornaam, achternaam,bedrijf,functie,info)
Function WriteHTMLSignature(sig_pad,sig_file,bestand, displaynaam, email, mobiel, bedrijf,functie,info,milieu,disclaimer_nl,disclaimer_uk,groet)
	Set signature = fso.OpenTextFile(sig_pad & sig_file & ".htm", ForWriting, True)

	' Melding tonen op scherm
	WScript.Echo "Bezig met opslaan e-mail handtekening in: " & sig_pad & sig_file & ".htm" 

	signature.Write "<!DOCTYPE HTML PUBLIC " & aQuote & "-//W3C//DTD HTML 4.0 Transitional//EN" & aQuote & ">" & vbCrLf
	signature.write "<HTML><HEAD><TITLE>Microsoft Office Outlook Signature</TITLE>" & vbCrLf
	signature.write "<META http-equiv=Content-Type content=" & aQuote & "text/html; charset=windows-1252" & aQuote & ">" & vbCrLf
	signature.write "<META content=" & aQuote & "MSHTML 6.00.3790.186" & aQuote & " name=GENERATOR></HEAD>" & vbCrLf
	signature.write "<BODY>" & vbCrLf
	signature.WriteLine "<p style='color: #000; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px;'>" & groet & "<br>"
	' Controleren of bedrijf een , (comma bevat), zoja dan splitsen op comma en 2e deel (Vestiging xxxxxx) op nieuwe regel plaatsen
	cc=","
	If (inStr(bedrijf,cc) > 0) Then
		'bedrijf bevat comma
		tmp=split(bedrijf,cc)
		signature.WriteLine tmp(0) & "<br>" & tmp(1) & vbCr
	Else
		'bedrijf bevat geen comma
		signature.WriteLine bedrijf &"</p>" & vbCr
	End If	

	'signature.WriteLine "<br>" & vbCr
	signature.WriteLine "<p style='color: #000; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px;'>" & displaynaam & "<br>" & vbCr
	signature.WriteLine functie & vbCr
	signature.WriteLine "</p>" & vbCr
	signature.WriteLine "-------------------------------------------------------" & vbCr
	signature.WriteLine "<div style='color: #005879; font-weight:bold; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 14px;'>" & vbCr
	' Controleren of bedrijf een , (comma bevat), zoja dan splitsen op comma
	cc=","
	If (inStr(bedrijf,cc) > 0) Then
		tmp=split(bedrijf,cc)
		signature.WriteLine tmp(0) & "<br>" & tmp(1) & vbCr
	Else
		'wscript.echo "geen comma"
		signature.WriteLine bedrijf & vbCr
	End If
	signature.WriteLine "</div>"& vbCr
	'Nu moeten we info nog over meerdere regels verdelen
	tekst=Split(info,vbCrLf)
	i=1
	
	stijl="style='color: #005879; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 12px; margin-top: 2px; margin-bottom: 2px;'"
	stijl1="style='color: #005879; font-family: Arial,Verdana, Helvetica, sans-serif;  font-size: 12px; margin-left: 35px; margin-right: 3px;'"
	stijl2="style='text-decoration: none; color: #005879; font-family: Arial,Verdana, Helvetica, sans-serif;  font-size: 12px; margin-left: 35px; margin-right: 3px;'"
	stijl3="style='text-decoration: none; color: #005879; font-family: Arial,Verdana, Helvetica, sans-serif;  font-size: 12px;'"
	
	i=1
	For Each regel in tekst
		regel=ltrim(regel)
		Select Case i
			Case 1
				signature.WriteLine "<table border='0' cellpadding='0' cellspacing='0'>"
				signature.WriteLine "<tr><td " & stijl & ">Postadres</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
			Case 2
				signature.WriteLine "<tr><td " & stijl & ">Bezoekadres</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
			Case 3
				signature.WriteLine "<tr><td " & stijl & ">Telefoon</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
				if (mobiel <> "") Then
					signature.WriteLine "<tr><td " & stijl & ">Mobiel</td><td " & stijl1 & ">:</td><td " & stijl & ">" & mobiel & "</td></tr>"
				End If
			Case 4
				signature.WriteLine "<tr><td " & stijl & ">Fax</td><td " & stijl1 & ">:</td><td " & stijl & ">" & regel & "</td></tr>"
				signature.WriteLine "<tr><td " & stijl & ">E-mail</td><td " & stijl1 & ">:</td><td " & stijl & ">" & email & "</td></tr>"
			Case 5
				signature.WriteLine "<tr><td " & stijl & ">Website</td><td " & stijl1 & ">:</td><td " & stijl3 & ">" & "<a " & stijl3 & " href='http://" & regel & "'>" & regel & "</a></td></tr>"
				signature.WriteLine "</table>"
		End Select
		
		i=i+1 'teller ophogen
	Next	


	'signature.WriteLine info & vbCr
	signature.WriteLine "<br><div style='font-weight:bold; color: #7fbb56; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 11px;'>" & vbCr
	signature.WriteLine milieu & vbCr
	signature.WriteLine "</div>" & vbCr 
	signature.WriteLine "" & vbCr 
	signature.WriteLine "<hr>"	& vbCr
	signature.WriteLine "<div style='color: #005879; font-family: Arial, Verdana, Helvetica, sans-serif;  font-size: 9px;'>" & vbCr
	signature.WriteLine disclaimer_nl & vbCr
	signature.WriteLine "<br><br>" & vbCr 
	signature.WriteLine disclaimer_uk & vbCr
	signature.WriteLine "</div>" & vbCr
	signature.WriteLine "<hr>"	& vbCr
	signature.WriteLine "</BODY>" & vbCr
	signature.WriteLine "</HTML>" & vbCr
	signature.Close
End Function


Sub SetDefaultSignature(strSigName, strProfile)
	Const HKEY_CURRENT_USER = &H80000001
	strComputer = "."

'	If Not IsOutlookRunning Then
		Set objreg = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
		strKeyPath = "Software\Microsoft\Windows NT\" & "CurrentVersion\Windows " & "Messaging Subsystem\Profiles\"

			'get default profile name if none specified
		If strProfile = "" Then
			objreg.GetStringValue HKEY_CURRENT_USER, strKeyPath, "DefaultProfile", strProfile
			'WScript.Echo strProfile
		End If

		'We moeten ervoor zorgen dat HKCU\Software\Microsoft\Office\11.0\Common\General\Signature de waarde "Handtekeningen"bevat, want dit is het pad naar
		'de handtekeningen map.
		objreg.GetStringValue HKEY_CURRENT_USER, "Software\Microsoft\Office\11.0\Common\General\", "Signatures", strSignaturePath
		If 	strSignaturePath <> "" Then
			'WScript.Echo strSignaturePath
		Else
			'WScript.Echo "Niet gevonden"
			' Set the string values
			strKey = "Signatures" ' New Key
			strWaarde = "Handtekeningen"
			strHandtekening = "HKCU\Software\Microsoft\Office\11.0\Common\General\"

			' Create the Shell object
			Set objShell = CreateObject("WScript.Shell")

			' These are the two crucial command in this script.
			objShell.RegWrite strHandtekening & strKey, 1, "REG_SZ"
			objShell.RegWrite strHandtekening & strKey, strWaarde, "REG_SZ"
		End If
		
		' Zorgen dat er geen briefpapier gebruikt wordt.
		objreg.GetStringValue HKEY_CURRENT_USER, "Software\Microsoft\Office\11.0\Common\Mailsettings\", "NewStationery", strNewStationery
		If 	strNewStationery <> "" Then
			WScript.Echo "Opruimen briefpapier instelling"
			Set objShell = CreateObject("WScript.Shell")
			objShell.RegDelete "HKCU\Software\Microsoft\Office\11.0\Common\Mailsettings\" & "NewStationery"
		End If
		
		' build array from signature name
		myArray = StringToByteArray(strSigName, True)
		strKeyPath = strKeyPath & strProfile & "\9375CFF0413111d3B88A00104B2A6676"
		objreg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrProfileKeys

		For Each subkey In arrProfileKeys
			'WScript.Echo "subkey:" & subkey
			strsubkeypath = strKeyPath & "\" & subkey
			objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "New Signature", myArray
			objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "Reply-Forward Signature", myArray
		Next
'	Else
'		strMsg = "Please shut down Outlook before " & "running this script." 
		' this is the message when error occurs, change it to desired.

'		MsgBox strMsg, vbExclamation, "SetDefaultSignature"
'	End If
End Sub

Function IsOutlookRunning()
	strComputer = "."
	strQuery = "Select * from Win32_Process " & "Where Name = 'Outlook.exe'"
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery(strQuery)
	For Each objProcess In colProcesses
		If UCase(objProcess.Name) = "OUTLOOK.EXE" Then
			IsOutlookRunning = True
		Else
			IsOutlookRunning = False
		End If
	Next
End Function

Public Function StringToByteArray (Data, NeedNullTerminator)
	Dim strAll
	strAll = StringToHex4(Data)
	If NeedNullTerminator Then
		strAll = strAll & "0000"
	End If
	intLen = Len(strAll) \ 2
	ReDim arr(intLen - 1)
	For i = 1 To Len(strAll) \ 2
		arr(i - 1) = CByte("&H" & Mid(strAll, (2 * i) - 1, 2))
	Next
	StringToByteArray = arr
End Function

Public Function StringToHex4(Data)
	' Input: normal text
	' Output: four-character string for each character,
	'e.g. "3204? for lower-case Russian B,
	'6500? for ASCII e
	'Output: correct characters
	'needs to reverse order of bytes from 0432
	Dim strAll
	For i = 1 To Len(Data)
		' get the four-character hex for each character
		strChar = Mid(Data, i, 1)
		strTemp = Right("00" & Hex(AscW(strChar)), 4)
		strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
	Next
	StringToHex4 = strAll
End Function


'
' Connect to the WinNT: provider and ask it for the user's group memberships.
' For each group, get the RID and compare it against the users PrimaryGroupID.
'
Public Function PrimaryGroup(btlUser)
   Dim objUser
   Dim Group, aGroup, PrimaryGroupRID
   Set objUser = GetObject("WinNT://" & Domain & "/" & btlUser & ",user")
   PrimaryGroupRID = objUser.Get("PrimaryGroupID")
   For Each Group in objUser.Groups
      aGroup = Group.Name
      If Rid(aGroup) = PrimaryGroupRID then
         PrimaryGroup = aGroup
         Exit Function
      End If
   Next
   Set objUser = Nothing
End Function
'
' Returns the RID, in decimal, of the specified group.
'
Public Function Rid(aGroup)
   Dim objGroup, Sid
   Dim sTmp, x, b
   Set objGroup = GetObject("WinNT://" & Domain & "/" & aGroup & ",group")
   Sid = objGroup.Get("objectSID")
   sTmp = ""
   For x = UBound(Sid) to UBound(Sid)-3 Step -1  ' Process last 4 bytes (the RID)
      b = AscB(MidB(SID, x + 1))                 ' Get a byte
      sTmp = sTmp & Hex(b \ 16) & Hex(b And 15)  ' Convert to hex
   Next
   Rid = Clng("&H" & sTmp)   ' Convert hex RID to long decimal
   Set objGroup = Nothing
End Function

'
' Get this users domain from the WshNetwork object
'
Function NTDomain
   Dim WshNetwork
   Set WshNetwork = WScript.CreateObject("WScript.Network")
   NTDomain = WshNetwork.UserDomain
   Set WshNetwork = Nothing
End Function