<#
 .Synopsis
  List all matching Certificates stored in the local certificate store with the Subject name entered. 
  Also allows user to Export the certificates either in CER or PFX format.
  
 .Description
  This module will search the local certificate store for all certificates with Subject name or DNS Name 
  matching with the Subject name provided as input. All certificates will be Exported to the User's Home directory.
  User will have to enter the PFX export password if they select -ExportPFX switch.
  Same password will be used for all Certificates.
  
 .Parameter subject
  subject, Subject name you want to search.
  
 .Example
  Find-LocalCerts -subject VpnServerRoot
  Find-LocalCerts -subject VpnServerRoot -ExportCER
  Find-LocalCerts -subject VpnServerRoot -ExportPFX
  
#>

#------------------------------------------------------------------------------
#
#
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
#
#------------------------------------------------------------------------------


function Find-LocalCerts () { 
Param (

    [Parameter(Mandatory=$true,
    ValueFromPipeline=$true)]
    [String]$subject,

    [Parameter(Mandatory=$false)]
    [Switch]$ExportCER,

    [Parameter(Mandatory=$false)]
    [switch]$ExportPFX 
)


[string]$sub1 = 'CN='+$subject

$loc = Get-ChildItem -path Cert:\ -Recurse   | Where-Object {$_.Subject -match $sub1 -or $_.DnsNameList -match $subject}

$certpath = @()
[array]$certpath = $loc | ForEach-Object {$_.PSParentPath.Split(":")[-1]}


# Custom object
$outtest = @()
for($i=0; $i -lt $loc.Count; $i++){

$outtest += New-Object PSObject -Property @{
Subject = $($loc[$i].Subject)
Thumbprint = $($loc[$i].Thumbprint)
Path = $certpath[$i]
Expiry = $($loc[$i].NotAfter)
Issuer = $($loc[$i].Issuer)
IsPFXExportable = $($loc[$i].PrivateKey.CspKeyContainerInfo.Exportable)
}}



# Export CER part
if($exportcer -eq $true){
Write-host "CER Certificates will be exported to $($HOME)" -ForegroundColor Green
$loc | ForEach-Object {  Export-Certificate -Cert $_ -FilePath $home"\$($_.Subject)$($_.Thumbprint).cer"}
}


# Export PFX Part
if($exportPFX -eq $true){
$locp = $loc | Where-Object {$_.PrivateKey.CspKeyContainerInfo.Exportable -eq $true}
if($locp -ne $null){
Write-host "PFX Certificates will be exported to $($HOME)" -ForegroundColor Green
$pass = Read-Host "Enter the password"
$mypwd = ConvertTo-SecureString -String $pass -Force -AsPlainText
$locp | ForEach-Object {Export-PfxCertificate -Cert $_ -FilePath $home"\$($_.Subject)$($_.Thumbprint).pfx" -Password $mypwd}}
Else{Write-Host "Cannot export non-exportable private key. Use -ExportCER Switch" -ForegroundColor Red}
}


$outtest

}



