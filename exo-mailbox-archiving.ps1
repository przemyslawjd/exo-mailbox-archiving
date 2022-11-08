<#
---------------------------------------------------------------
Name:  Exchange Online - Mailbox in-place archiving setup
Description: A simple powershell script to enable and check status the in-place archiving of the exchange mailbox. 
Requirements: Exchange Online PowerShell module [Install-Module -Name ExchangeOnlineManagement]
Permission: You must be assigned the Mail Recipients role in Exchange Online to enable or disable archive mailboxes.  
Author: https://github.com/przemyslawjd
---------------------------------------------------------------
#>

#### Parameters and variables
Param(
    [Parameter(Mandatory = $false)]
    [string]$MailboxUserName,
    [string]$MailboxDomainName,
    [int]$Action
)

$Global:MailboxUserName = $MailboxUserName
$Global:MailboxDomainName = $MailboxDomainName

#### Function - Check exo module and connect to exchange
function ConnectToExchange() {
    
    Write-Host "Checking Exchange Online Menagement module availability ..." -ForegroundColor Cyan

    $ExchangetModuleList = get-module -ListAvailable -Name "ExchangeOnlineManagement"

    if ($ExchangetModuleList.count -eq 0) {
            Write-Host "No module installed, please install the module and try again" -ForegroundColor Red
            Exit
        }

    else {
            Write-Host "EXO Module is available" -ForegroundColor Green
            Write-Host "Authentication Exchange administrator..." -ForegroundColor Cyan
            Get-PSSession | Remove-PSSession
            
            try {
                Connect-ExchangeOnline -ErrorAction Stop
                Write-Host "Authentication to Exchange was successful" -ForegroundColor Green
                }

            catch {
                    Write-Host "Connection has not been established, please try again later" -ForegroundColor Red
                    Exit
                }
        }
}

#### Function - Get username and domain
function GetUserInfo {

    if ($global:MailboxDomainName -eq "") {
        Write-Host "Enter company domain name"  -ForegroundColor Yellow
        $global:MailboxDomainName = Read-Host -Prompt "Domain"
    }

    if ($global:MailboxUserName -eq "") {
        Write-Host "Enter mailbox username in domain <$global:MailboxDomainName> to manage archive settings"  -ForegroundColor Yellow
        $global:MailboxUserName = Read-Host -Prompt "Username"
    }

    $global:FullMailboxUserName = $global:MailboxUserName + '@' + $global:MailboxDomainName
}

#### Fuction [1]- Check and enable archive
function EnableArchive() {

    Write-Host ""        
    Write-Host "### Enable Archive on Mailbox ###" -ForegroundColor Magenta     
    Write-Host ""

    if (((Get-EXOMailbox -Identity $global:FullMailboxUserName).ArchiveStatus) -eq 'Active') {

        Write-Host "Archive on mailbox $global:FullMailboxUserName is curently enabled" -ForegroundColor Green
    }
    
    else {

        Write-Host "Archive on mailbox $global:FullMailboxUserName is curently disabled " -ForegroundColor Red
        Write-Host "Attempt to enable archiving on mailbox $global:FullMailboxUserName..." -ForegroundColor Cyan
    
        try {
                Get-EXOMailbox -Identity $global:FullMailboxUserName | Enable-Mailbox -Archive -ErrorAction stop
                Write-Host "Archiving of the mailbox $global:FullMailboxUserName has been started" -ForegroundColor Green
            }
            
        catch {
                Write-Host "Archive startup error on mailbox $global:FullMailboxUserName, try again later " -ForegroundColor Green
                exit
              } 
    }

    Write-Host -NoNewLine -ForegroundColor Yellow  'Press any key to return to the menu...' 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host ""
}

#### Function [2] - Assign archive policy

function AssignPolicy() {

    Write-Host ""        
    Write-Host "### Assign archive policy on mailbox ###" -ForegroundColor Magenta    
    Write-Host ""

    $PolicyRetention = Get-RetentionPolicy
    $PolicyList = @{}

    foreach ($Policy in $PolicyRetention) {

        $CurrentPolicyIndex = $PolicyRetention.IndexOf($Policy)+1
        $CurrentPolicyName = ($Policy.Name).ToString()

        Write-Host "    $CurrentPolicyIndex. $CurrentPolicyName"
        $PolicyList.Add($CurrentPolicyIndex, $CurrentPolicyName)
    }

    Write-Host ""
    Write-Host "Select policy number to apply on mailbox " -ForegroundColor Yellow 
    Write-Host ""
    [int]$PolicyIndexToSet = Read-Host -Prompt "Policy number"

    if ($PolicyIndexToSet -ge 1 -and $PolicyIndexToSet -le $PolicyRetention.count) {
   
        try {
            Set-Mailbox $global:FullMailboxUserName -RetentionPolicy $PolicyList[$PolicyIndexToSet]
            Write-Host "Successfull set policy" -ForegroundColor Green
        }

        catch {
            Write-Host "There was a problem with the policy assignment: $PolicyList[$PolicyIndexToSet]" -ForegroundColor Red
            exit
        }
    }

    else {
        
        Write-Host "Incorrect policy index, please try again" -ForegroundColor Red
        exit
    }

    Write-Host -NoNewLine -ForegroundColor Yellow  'Press any key to return to the menu...' 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host ""
}

#### Function [3]- Disable retention hold

function DisableRetentionHold() {

    Write-Host ""        
    Write-Host "### Disable retention hold on mailbox ###" -ForegroundColor Magenta     
    Write-Host ""

    Write-Host "Disabling retention hold..." -ForegroundColor cyan

    try {
        Set-Mailbox $global:FullMailboxUserName -RetentionHoldEnabled $false 
        Write-Host "Retention hold is successfully disabled" -ForegroundColor green
    }
    catch {
        Write-Host "Disabling retention hold ended with an error, try again" -ForegroundColor red
        exit
    }

    Write-Host -NoNewLine -ForegroundColor Yellow  'Press any key to return to the menu...' 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host ""
}

#### Function [4] - Enable Folder Assistant
function EnableFolderAssistant() {
    
    Write-Host ""        
    Write-Host "### Enable Folder Assistant on Mailbox ###" -ForegroundColor Magenta     
    Write-Host ""
    Write-Host "Enabling Managed Folder Assistant..." -ForegroundColor cyan

    try {
        (Get-EXOMailbox $global:FullMailboxUserName).identity | Start-ManagedFolderAssistant -ErrorAction Stop
        Write-Host "Folder Management Assistant has been successfully enabled" -ForegroundColor green
    }
    catch {
        Write-Host "Folder Management Assistant was not enabled due to an error" -ForegroundColor red
        exit
       }

       Write-Host -NoNewLine -ForegroundColor Yellow  'Press any key to return to the menu...' 
       $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
       Write-Host ""
}

#### Function [5] - Show mailbox information

function MailboxStatus() {
    Write-Host ""        
    Write-Host "### Mailbox archive status ###" -ForegroundColor Magenta    
    Write-Host ""

    Write-Host "Total mailbox size" -ForegroundColor Cyan
    Get-EXOMailboxStatistics -Identity ($global:FullMailboxUserName) | Select-Object TotalItemSize, ItemCount | Out-Host

    Write-Host "Current mailbox archive settings" -ForegroundColor Cyan
    Get-Mailbox -Identity $global:FullMailboxUserName | Select-Object -Property PrimarySmtpAddress, ArchiveStatus, ArchiveName, RetentionPolicy, RetentionHoldEnabled | Out-Host
    
    if ((Get-Mailbox -Identity $global:FullMailboxUserName).ArchiveStatus -eq 'Active') {
        Write-Host "Mailbox archive progress" -ForegroundColor Cyan
        Get-EXOMailboxStatistics -Identity ($global:FullMailboxUserName) -Archive | Select-Object DisplayName, TotalItemSize, ItemCount | Out-Host
    }

    Write-Host -NoNewLine -ForegroundColor Yellow  'Press any key to return to the menu...' 
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    Write-Host ""
}

#### Function [6] - Change mailbox
function ChangeUserInfo {

    Write-Host ""        
    Write-Host "### Change mailbox ###" -ForegroundColor Magenta    
    Write-Host ""

    Write-Host "Enter mailbox username in domain <$global:MailboxDomainName> to manage archive settings"  -ForegroundColor Yellow
    $global:MailboxUserName = Read-Host -Prompt "Username"
    $global:FullMailboxUserName = $global:MailboxUserName + '@' + $global:MailboxDomainName
    Write-Host ""
}
   
##### Main function
function Main() {
    do {
        if($Action -eq "") {               
            Write-Host ""        
            Write-Host "--------------------------------" -ForegroundColor Cyan
            Write-Host "Currently managed mailbox: $global:FullMailboxUserName" -ForegroundColor Cyan     
            Write-Host "--------------------------------" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "    1. Enable archive"
            Write-Host "    2. Assign policy"
            Write-Host "    3. Disable retention hold"
            Write-Host "    4. Enable Managed Folder Assistant "
            Write-Host "    5. Show mailbox information/archive status"
            Write-Host "    6. Change mailbox to menage"
            Write-Host "    0. Exit"
            Write-Host "" 

            $GetAction = Read-Host 'Please choose the action to continue' 
            Write-Host "" 
       }

       else {
        $GetAction=$Action
       }
       
       switch ($GetAction) {

            1 {EnableArchive}
            2 {AssignPolicy}
            3 {DisableRetentionHold}
            4 {EnableFolderAssistant}
            5 {MailboxStatus}
            6 {ChangeUserInfo}
            0 {exit}
       }

    }   While ($GetAction -ne 0) 
}

##### Start
Write-Host "### Management mailbox archiving on exchange online server ###" -ForegroundColor Magenta 

ConnectToExchange
GetUserInfo
Main


