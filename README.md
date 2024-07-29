# Signatures-autofilled

## EXO settings
Create a new rule, apply dislaimer rules

%%displayname%% will fill "Name Surname" value <br />
%%Phone%% -> User's direct phone <br />
#XXXXXX change with company hex color <br />
Change website and office phone <br />

```
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head></head>
<body>
<div>&nbsp;</div>
<div><span style="font-family: 'arial'; font-size: 9pt;"><strong>%%displayname%%</strong><br />
<div><span style="font-family: 'arial'; font-size: 9pt;">------------------------------<br />
<div></span><span style="font-family: 'arial'; font-size: 9pt;">%%Phone%%</span></div>
<div>&nbsp;</div>
<div><img src="https://signatures.banner/logo.png" /></div>
<div><span style="font-family: 'arial';font-size: 7.5pt; color: #XXXXXX;"><strong>COMPANY NAME</strong> | Address </span><br />
<span style="font-family: 'arial';font-size: 7.5pt; color: #XXXXXX;">T <a href="callto:+123456789">+39 123456789</a></span><br />
<span style="font-family: 'arial';font-size: 7.5pt; color: #XXXXXX;"><span ><a href="https://www.website.company" style="color: #00722d;" target="_blank" rel="noopener">website.company</a></span></div>
```

## Powershell settings (us on Active Directory)

Download Signatures txt,rtf,htm <br />

```
timeout 3
#Variabili da compilare
$CompFolder=$args[0]
$area=$args[1]
$function=$args[2]
write-host $CompFolder
echo $CompFolder
write-host $function
echo $function
write-host $area
echo $area


if ($CompFolder -eq $null )
{
    echo "set paratemers:"
    echo "company LIS"
    echo "area like def, com"
    echo "function like empty, holiday, closed, exhibition_1, exchibition_2"
    echo "like .\scripts ame com holiday"
    exit
} 


$UPN = whoami /upn
$separator = "@"
$UPNparts = $UPN.split($separator)
$UPN_dom_uppercase = $UPNparts[1].Substring(0,1).ToUpper() + $UPNparts[1].Substring(1)

$ROAMING = "(" + $UPNparts[0] + $Separator + $UPN_dom_uppercase + ")"
$ROAMING_undercase = " (" + $UPN + ")"
$User = $env:UserName
$FileName = "Signature"
$FileExtension = "htm","rtf","txt"
$Path = "$env:TEMP"
$PathSignature = "$Path"
$PathSignatureUser = "$PathSignature\$User"
$AppSignatures =$env:APPDATA + "\Microsoft\Signatures"

#echo $UPNparts[0]
#echo $UPNparts[1]
#echo $UPN_dom_uppercase
echo $ROAMING
echo $ROAMING_undercase
#timeout 50

####################################Set AREA DEFAULT

$PathSignatureTemplates = "\\corp.companygroup.it\svc\group\app_deploy_group\M365\Signatures\$CompFolder"


####################################Set FUNCTION
#upper= comunicazioni testuali
#mid= fiere o eventi
#lower= banner azienda standard

if ($function -match "exhibition_1" )
{
$URL_lower = "https://signatures.companygroup.it/lower-banner/$CompFolder.png"
$URL_upper = "https://signatures.companygroup.it/upper-banner/$CompFolder.png"
$URL_mid = "https://signatures.companygroup.it/exh1/$CompFolder.png"
$URL_last = "https://signatures.companygroup.it/last-banner/$CompFolder.png"
}

if ($function -match "exhibition_2" )
{
$URL_lower = "https://signatures.companygroup.it/lower-banner/$CompFolder.png"
$URL_upper = "https://signatures.companygroup.it/upper-banner/$CompFolder.png"
$URL_mid = "https://signatures.companygroup.it/exh2/$CompFolder.png"
$URL_last = "https://signatures.companygroup.it/last-banner/$CompFolder.png"
}

if ($function -eq $null )
{
$function = "signature"
$URL_lower = "https://signatures.companygroup.it/lower-banner/$CompFolder.png"
$URL_upper = "https://signatures.companygroup.it/upper-banner/$CompFolder.png"
$URL_mid = "https://signatures.companygroup.it/mid-banner/$CompFolder.png"
$URL_last = "https://signatures.companygroup.it/last-banner/$CompFolder.png"
}
echo $URL_lower


####################################Set TERMS AND CONDITION FOR COM
if ($CompFolder -match "company_4" )
{
    $TeC = "https://www.company_1.it"
}

if ($CompFolder -match "company_1" )
{
    $TeC = "https://www.company_1.it/sites/default/files/2020-10/sales-conditions_company_12020_en.pdf"
}

if ($CompFolder -match "company_5" )
{
    $TeC = "https://www.company_1.it"
}

if ($CompFolder -match "company_6" )
{
    $TeC = "https://www.company_1.it"
}

if ($CompFolder -match "company_3" )
{
    $TeC = "https://www.company_3srl.it/sites/default/files/2021-06/sales-conditions-company_32020-en.pdf"
}

if ($CompFolder -match "company_7" )
{
    $TeC = "https://www.company_1.it"
}

if ($CompFolder -match "company_2" )
{
    $TeC = "https://www.company_2actuators.com/sites/default/files/2021-06/sales-conditions-company_2.pdf"
}

if ($CompFolder -match "company_8" )
{
    $TeC = "https://www.company_1.it"
}

if ($CompFolder -match "company_9" )
{
    $TeC = "https://www.company_9srl.it/sites/default/files/2019-11/company_9-general-conditions-of-sale.pdf"
}


####################################Set Quality Declaration Link
if ($CompFolder -match "company_4" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_1" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_5" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_6" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_3" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_7" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_2" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_8" )
{
    $QDL = "https://www.company_1.it/download"
}

if ($CompFolder -match "company_9" )
{
    $QDL = "https://www.company_1.it/download"
}


####################################USER PROP

#$CurrentUser=( ( Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username ) -split '\\' )[1]
$CurrentUser=$env:USERNAME
$ADTitle=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['title']
$ADCn=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['cn']
$ADMobile=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['mobile']
$ADTelephoneNumber=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['telephoneNumber']
$ADDepartment=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['Department']
$ADCompany=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['Company']
$ADOfficePhone=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['OfficePhone']
$ADMail=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['Mail']
$ADHomepage=([adsisearcher]"(&(objectClass=user)(samaccountname=$CurrentUser))").FindOne().Properties['Homepage']

#echo $ADTitle
#echo $ADCn
#echo $ADMobile
#echo $ADTelephoneNumber
#echo $ADDepartment
#echo $ADCompany
#echo $ADOfficePhone
#echo $ADMail
#echo $ADHomepage
#timeout 20
echo $QDL

####################################RUN SCRIPT DEFAULT

################FIRMA STANDARD PER TUTTI

$Company = $ADCompany

New-Item -Path "$PathSignature\$User" -ItemType Container –Force

foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignatureTemplates\$FileName.$Ext" "$PathSignatureUser\$FileName.$Ext"
}

if ($ADMobile -match ".{2}" )
{
    $MobileP = "M " + $ADMobile
}


if ($ADTelephoneNumber -match ".{4}" )
{
    $DirectP = "T " + $ADTelephoneNumber
}

foreach ($Ext in $FileExtension)
{
(Get-Content "$PathSignatureUser\$FileName.$Ext") | Foreach-Object {
$_`
-replace "@NAME", $ADCn`
-replace "@TITLE",$ADTitle `
-replace "@DEPARTMENT",$ADDepartment `
-replace "@COMPANY", $ADCompany `
-replace "@MOBILE", $MobileP `
-replace "@DIRECT", $DirectP `
-replace "@OFFICEPHONE", $ADOfficePhone `
-replace "@EMAIL", $ADMail `
-replace "@WEBSITE", $ADHomepage `
-replace "@LOGOIMG", $LOGO_IMG `
-replace "@DOWN", $URL_lower `
-replace "@MID", $URL_mid `
-replace "@UP", $URL_upper `
-replace "@TEC", $TeC `
-replace "@QD", $QDL `
-replace "@LAST", $URL_last `

} | Set-Content "$PathSignatureUser\$FileName.$Ext"
}


foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function.$Ext"
write-host "$PathSignatureUser\$FileName $function.$Ext"
write-host "$AppSignatures\$CompFolder $function.$Ext"

Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function $ROAMING.$Ext"
write-host "$PathSignatureUser\$FileName $function $ROAMING.$Ext"
write-host "$AppSignatures\$CompFolder $function $ROAMING.$Ext"

Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function $ROAMING_undercase.$Ext"
write-host "$PathSignatureUser\$FileName$function $ROAMING_undercase.$Ext"
write-host "$AppSignatures\$CompFolder$function $ROAMING_undercase.$Ext"
}


#############SE AREA COMMERICALE

if ($area -match "com" )
{

$PathSignatureTemplates = "\\corp.companygroup.it\svc\group\app_deploy_group\M365\Signatures\_TermsCond\$CompFolder"

####################################RUN SCRIPT TERMS AND CONFITION

$Company = $ADCompany
New-Item -Path "$PathSignature\$User" -ItemType Container –Force

foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignatureTemplates\$FileName.$Ext" "$PathSignatureUser\$FileName.$Ext"
}

if ($ADMobile -match ".{2}" )
{
    $MobileP = "M " + $ADMobile
}


if ($ADTelephoneNumber -match ".{4}" )
{
    $DirectP = "T " + $ADTelephoneNumber
}

foreach ($Ext in $FileExtension)
{
(Get-Content "$PathSignatureUser\$FileName.$Ext") | Foreach-Object {
$_`
-replace "@NAME", $ADCn`
-replace "@TITLE",$ADTitle `
-replace "@DEPARTMENT",$ADDepartment `
-replace "@COMPANY", $ADCompany `
-replace "@MOBILE", $MobileP `
-replace "@DIRECT", $DirectP `
-replace "@OFFICEPHONE", $ADOfficePhone `
-replace "@EMAIL", $ADMail `
-replace "@WEBSITE", $ADHomepage `
-replace "@LOGOIMG", $LOGO_IMG `
-replace "@DOWN", $URL_lower `
-replace "@MID", $URL_mid `
-replace "@UP", $URL_upper `
-replace "@TEC", $TeC `
-replace "@QD", $QDL `
-replace "@LAST", $URL_last `

} | Set-Content "$PathSignatureUser\$FileName.$Ext"
}

foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function TeC.$Ext"
write-host "$PathSignatureUser\$FileName $function TeC.$Ext"
write-host "$AppSignatures\$CompFolder $function TeC.$Ext"

Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function TeC $ROAMING.$Ext"
write-host "$PathSignatureUser\$FileName $function TeC $ROAMING.$Ext"
write-host "$AppSignatures\$CompFolder $function TeC $ROAMING.$Ext"

Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function TeC $ROAMING_undercase.$Ext"
write-host "$PathSignatureUser\$FileName$function TeC $ROAMING_undercase.$Ext"
write-host "$AppSignatures\$CompFolder$function TeC $ROAMING_undercase.$Ext"
}
############TEMP DELETE OLD SIGNATURES
Remove-Item -Path "$AppSignatures\$CompFolder holidays TeC.htm"
Remove-Item -Path "$AppSignatures\$CompFolder holidays TeC.txt"
Remove-Item -Path "$AppSignatures\$CompFolder holidays TeC.rtf"
Remove-Item -Path "$AppSignatures\$CompFolder closed TeC.htm"
Remove-Item -Path "$AppSignatures\$CompFolder closed TeC.txt"
Remove-Item -Path "$AppSignatures\$CompFolder closed TeC.rtf"

} 


#########RIMOZIONE VECCHIE FIRME
Remove-Item -Path "$AppSignatures\$CompFolder holidays.htm"
Remove-Item -Path "$AppSignatures\$CompFolder holidays.txt"
Remove-Item -Path "$AppSignatures\$CompFolder holidays.rtf"
Remove-Item -Path "$AppSignatures\$CompFolder closed.htm"
Remove-Item -Path "$AppSignatures\$CompFolder closed.txt"
Remove-Item -Path "$AppSignatures\$CompFolder closed.rtf"
Remove-Item -Path "$AppSignatures\$CompFolder Editable.htm"
Remove-Item -Path "$AppSignatures\$CompFolder Editable.txt"
Remove-Item -Path "$AppSignatures\$CompFolder Editable.rtf"



###########FIRMA EDITABILE DA UTENTE 
$Editable = "$AppSignatures\$CompFolder Editabile.htm"

##SE LA FIRMA EDITABILE NON EISTE LA DEVE CREARE

if (-not(Test-Path -Path $Editable -PathType Leaf)) {

$Company = $ADCompany
$function = "Z-Editabile"
$PathSignatureTemplates = "\\corp.companygroup.it\svc\group\app_deploy_group\M365\Signatures\$CompFolder"

$URL_lower = "https://signatures.companygroup.it/lower-banner/$CompFolder.png"
$URL_mid = "https://signatures.companygroup.it/fixed/pixel.png"
$URL_upper = "https://signatures.companygroup.it/fixed/pixel.png"
$URL_last = "https://signatures.companygroup.it/fixed/pixel.png"

New-Item -Path "$PathSignature\$User" -ItemType Container –Force

foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignatureTemplates\$FileName.$Ext" "$PathSignatureUser\$FileName.$Ext"
}

if ($ADMobile -match ".{2}" )
{
    $MobileP = "M " + $ADMobile
}


if ($ADTelephoneNumber -match ".{4}" )
{
    $DirectP = "T " + $ADTelephoneNumber
}

foreach ($Ext in $FileExtension)
{
(Get-Content "$PathSignatureUser\$FileName.$Ext") | Foreach-Object {
$_`
-replace "@NAME", $ADCn`
-replace "@TITLE",$ADTitle `
-replace "@DEPARTMENT",$ADDepartment `
-replace "@COMPANY", $ADCompany `
-replace "@MOBILE", $MobileP `
-replace "@DIRECT", $DirectP `
-replace "@OFFICEPHONE", $ADOfficePhone `
-replace "@EMAIL", $ADMail `
-replace "@WEBSITE", $ADHomepage `
-replace "@LOGOIMG", $LOGO_IMG `
-replace "@DOWN", $URL_lower `
-replace "@MID", $URL_mid `
-replace "@UP", $URL_upper `
-replace "@LAST", $URL_last `

} | Set-Content "$PathSignatureUser\$FileName.$Ext"
}


foreach ($Ext in $FileExtension)
{
Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function.$Ext"
write-host "$PathSignatureUser\$FileName $function.$Ext"
write-host "$AppSignatures\$CompFolder $function.$Ext"

Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function $ROAMING.$Ext"
write-host "$PathSignatureUser\$FileName $function $ROAMING.$Ext"
write-host "$AppSignatures\$CompFolder $function $ROAMING.$Ext"

Copy-Item -Force "$PathSignatureUser\$FileName.$Ext" "$AppSignatures\$CompFolder $function $ROAMING_undercase.$Ext"
write-host "$PathSignatureUser\$FileName$function $ROAMING_undercase.$Ext"
write-host "$AppSignatures\$CompFolder$function $ROAMING_undercase.$Ext"

}
}
 else {
    timeout 1
 }
```

