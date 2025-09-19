<#
Name: John Asare
Date: 2025-09-15
Description: Azure PowerShell script to get properties of a app service  in the asareware subscription
#>

#import the Azure module
Import-Module Az

# Connect to Azure account
Connect-AzAccount

# Get properties of a app service
Get-AzWebApp -ResourceGroupName "rg-mailbox-quota" -Name "wa-mailbox-quota-windows" 