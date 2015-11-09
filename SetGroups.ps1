# We need some credentials of course
$UserCredential = Get-Credential
 
# Create the session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://outlook.office365.com/powershell-liveid/ `
    -Credential $UserCredential `
    -Authentication Basic -AllowRedirection
 
Set-Location $PSScriptRoot
[xml]$xmlfile = Get-Content ".\XML\Groups.xml"

$Groups = $xmlfile.Groups.Group

# Import the session
Import-PSSession $Session
 
foreach ($Group in $Groups) {
    try {
        # Set Group
        Set-UnifiedGroup -Identity $Group.Identity -DisplayName $Group.DisplayName -MailTip $Group.MailTip -PrimarySmtpAddress $Group.PrimarySmtpAddress
        Write-Host "Group changed to:" $Group.Identity "/" $Group.DisplayName "/" $Group.MailTip "/" $Group.PrimarySmtpAddress -ForegroundColor Green
    }
    catch {
        Write-Host "Oops! Something went wrong with the following group:" $Group.DisplayName -ForegroundColor Red
    }
    finally {
        Write-Host "End of script"
    }
}
 
# Kill the session
Remove-PSSession $Session