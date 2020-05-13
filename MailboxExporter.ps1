#################################################
#                                               #
#    Exchange Mailbox Export Script             #
#      Tested on Exchange 2016	                #
#                                               #
#    v.1.0 - Sebastian Storholm 13.05.2020      #
#                                               #
#################################################

# The script exports the specified users mailbox as a PST to the specified path with the filename [alias]_YYYYMMDD.pst 

$UserCredential = Get-Credential
$SessionExchange = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.example.com/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $SessionExchange -DisableNameChecking

# Path to network share used to store the exported PST-files
$ServerPath = "\\SERVER1\PST Backups\Exported\"

Write-Host "Mailbox PST Exporter v.1.0"

$Alias = Read-Host -Prompt 'User Alias?';
$UserName = $Alias.replace( ".","_")
$Date = Get-Date -f yyyyMMdd
$Path = $ServerPath + $UserName + "_" + $Date + ".pst”
New-MailboxExportRequest -Mailbox $Alias -FilePath $Path

# Check progress using either of these:
# Get-MailboxExportRequestStatistics -Identity $Alias
# Get-MailboxExportRequest | Get-MailboxExportRequestStatistics
