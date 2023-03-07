function Write-Log
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)][System.String]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "$timestamp - $Message"
}

function GetOpperation{
    [CmdletBinding()]param ([System.Int32]$Selected)
    try{
        write-host ""
        Write-Host "Select Opperation to perform" -ForegroundColor Green
        Write-Host "-----------------------------------------------------------------"`n -ForegroundColor Green
        Write-Host "0) Exit" -ForegroundColor Green
        Write-Host "1) List available templates" -ForegroundColor Green
        Write-Host "2) Export a site as template"-ForegroundColor Green
        Write-Host "3) Import Template"-ForegroundColor Green
        Write-Host "4) Delete Design"-ForegroundColor Green
        Write-Host "5) Delete Script"-ForegroundColor Green
        Write-Host ""

        $Selected = Read-Host -Prompt "Please enter an opperation"
        if ($Selected -eq 0) {
            Write-Host " Exiting ..."
            exit
        }
        elseif($Selected -eq 1){        
            ListTemplate
            GetOpperation
            Disconnect-PnPOnline 
        }
        elseif ($Selected -eq 2) {
            ExportSiteAsTemplate
            GetOpperation
            Disconnect-PnPOnline 
        }
        elseif ($Selected -eq 3) {
            ImportTemplate
            GetOpperation
            Disconnect-PnPOnline 
        }
        elseif ($Selected -eq 4) {
            DeleteDesign
            GetOpperation
            Disconnect-PnPOnline 
        }
        elseif ($Selected -eq 5) {
            DeleteScript
            GetOpperation
            Disconnect-PnPOnline 
        }
    }
    catch{
        Write-Host "Error Invalid Entry "
        Write-Host $_.Exception.Message
    }
}

Function ConnectCert($site_url){
    try{
        Connect-PnPOnline -Tenant $tenant -ClientId $AppId -Thumbprint $Cert -Url $site_url 
        $connection = Get-PnPConnection
        if($connection){
         Write-Host "Connected"
        }
        else {
         Write-Host "Could not connect!"
        }
    }
    catch{
        Write-Host $_.Exception.Message
    }
    
}

#List all sit scripts and designs
Function ListTemplate{    
    try{
        $siteURL = "https://lxmk.sharepoint.com"
        ConnectCert($siteUrl) 
        Write-Host "Site Scripts"
        $script = Get-pnpSiteScript 
    
        foreach($s in $script){
            Write-Host "-" $s.title 
            write-host "-- Script ID: "$s.Id
            write-host "-- Version: " $s.version
            write-host ""
        }
        Write-Host ""
        write-Host "Designs"
        $design = Get-PnPSiteDesign 
    
        foreach($d in $design){
            write-host "- " $d.title 
            write-host "-- iDefault: " $d.isDefault 
            write-host "-- Design ID: " $d.Id
            write-host "-- Script ID: " $d.SiteScriptIds
            write-host "-- Version: " $d.version
            write-host "-- WebTemplate: " $d.WebTemplate
            write-host ""
        }

    }catch{
        Write-Host -ForegroundColor red "Could not list templates"
        Write-Host -ForegroundColor red $_.Exception.Message
    }
    
    
}

#Export a SharePoint site as a template. Option2 
Function ExportSiteAsTemplate{
    try{
        $siteURL = Read-Host -Prompt "Enter the URL to Export "
        $templateName = Read-Host -Prompt "Enter a name for the template "
        write-host $siteURL
        $outFilePath = "Templates/$templateName.json"
        Connect-PnPOnline -Tenant $tenant -ClientId $AppId -Thumbprint $Cert -Url $siteURL
        $template = Get-PnPSiteScriptFromWeb -Url $siteURL -IncludeAll
        $template | Out-File $outFilePath

    }catch{
        Write-Host -ForegroundColor red "Could not export"
        Write-Host -ForegroundColor red $_.Exception.Message
    }
    
    
}

Function ImportTemplate{
    try{
        $templateName = Read-Host -Prompt "Enter a template name "
        $templateDescription = Read-Host -Prompt "Enter a description "
        
        $isDefault = Read-Host -Prompt "Set as default (Y) "
        $siteURL = Read-Host -Prompt "Site URL "
        $templatePath = Read-Host -Prompt "Enter the template path "
        $imagURL = "https://lxmk.sharepoint.com/sites/SiteTemplates/SiteAssets/blueprint.png"
        $webTemplate = "64" #64 = Team Site, 68 = Communication Site, 1 = Groupless Team Site
        ConnectCert($siteURL)
        
        $siteScript = Add-PnPSiteScript -Title $templateName -Content (Get-Content $templatePath -Raw)
        Write-Host "Script ID: "$siteScript.Id
        if($isDefault -eq "Y" ){
            $siteDesign = Add-PnPSiteDesign -SiteScriptIds $siteScript.Id -Title $templateName -IsDefault -WebTemplate $webTemplate -Description $templateDescription -ThumbnailUrl $imagURL
            Write-host "isDefault" $siteDesign.isDefault
            
        }
        else{
            $siteDesign = Add-PnPSiteDesign -SiteScriptIds $siteScript.Id -Title $templateName -WebTemplate $webTemplate -Description $templateDescription -ThumbnailUrl $imagURL
            
            Write-host $siteScript.isDefault
        }
    }
    catch{
        Write-Host -ForegroundColor red "Could not import"
        Write-Host -ForegroundColor red $_.Exception.Message
    }
    
}

Function DeleteScript{
    try{
        ConnectCert("https://lxmk.sharepoint.com")
        $id = Read-Host -Prompt "Script Id "
        Remove-PnPSiteScript -Identity $id -Force
    }catch{
        Write-Host -ForegroundColor red "Could not delete"
        Write-Host -ForegroundColor red $_.Exception.Message
    }
}

Function DeleteDesign{
    try{
        ConnectCert("https://lxmk.sharepoint.com")
        $id = Read-Host -Prompt "Script Id "
        Remove-PnPSiteDesign -Identity $id -Force
    }catch{
        Write-Host -ForegroundColor red "Could not delete"
        Write-Host -ForegroundColor red $_.Exception.Message
    }
    

}



##################################################################################
##################################################################################

#Config file path
$Configfile = Join-Path $PSScriptRoot -ChildPath "\Config\config.json"


#Import variables from config file
$Config = Get-Content $Configfile |ConvertFrom-Json
$AppId = $Config.AppId
$Cert = $Config.CertificateThumbprint
$Tenant = $Config.Tenant
$URL = $Config.URL

Write-Host $tenant
Write-Host "-----------------------------------------------------------------"
Write-Host ""

Write-Log -Message 'Checking for required modules'
$RequiredModules = @('PnP.PowerShell')

 foreach($m in $RequiredModules){
   $Module = Get-Module $m -ListAvailable
   if ($Module){ 
        Write-Log -Message "$m Versionersion: $($Module.Version)"
   }
   else{
        Write-Log "[ERROR] not all required modules are installed, cancelling!"
        exit
    } 
} 
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
 { $ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition }
 else
 { $ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
     if (!$ScriptPath){ $ScriptPath = "." } 
}

#Display menu
GetOpperation
