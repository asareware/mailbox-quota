<#
Name: John Asare
Date: 2025-09-15
Description: Azure PowerShell configuration script modifying and creating resources in the asareware subscription.
#>

# check if Azure module is installed and if not, install it
if (-not (Get-Module -ListAvailable -Name Az)) {
    Install-Module -Name Az -Scope CurrentUser -Repository PSGallery -Force
} else {
    Write-Host "Azure module is already installed."
}

#import the Azure module
Import-Module Az

# Connect to Azure account
Connect-AzAccount

# Get properties of a fucntion app
Get-AzFunctionAppSetting -ResourceGroupName "rg-mailbox-quota" -Name "fa-mailbox-quota" 