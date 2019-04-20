# Exchange Calendar Access Manager for On Prem or 365 Exchange Online - A Welsh

# Constants
$strOnPremServer = "https://server.domain.tld"
$strURI = ($strOnPremServer + "/powershell-liveid/")

# Functions
Function Clean-and-Split($Dirty) {
	$Dirty = $Dirty.Replace(" ", "")
	$Dirty = $Dirty.Replace(";", ",")
	$Dirty = $Dirty.Split(",")
	Return $Dirty
}

Function Connect-365 {
	$strCredential = Get-Credential
	$365Esession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $strURI -Credential $strCredential -Authentication Basic -AllowRedirection
	Import-PSSession $365Esession
}

Function Connect-OnPrem {
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchERI -Authentication Kerberos
	Import-PSSession $Session
}

# Select Connection
Clear-Host
Write-Host "Exchange Online or On Premise Exchange?"
Write-Host "1 - Exchange Online"
Write-Host "2 - Exchange OnPrem"
$strExchange = Read-Host ">"

If ($strExchange -eq "1"){
	Connect-365
}
ElseIf ($strExchange -eq "2"){
	Connect-OnPrem
}
Clear-Host

<#-------ACTUAL EXCHANGE WORK-------#>

# Input action type
$strChoice = Read-Host "Enter 1 to grant rights or 2 to remove"

# Choice 1 - Grant Access
If ($strChoice -eq 1){
	Clear-Host

	# Input username(s) that need acccess
	Write-Host "Enter the LoginID(s) of the user(s) that needs calendar access."
	$strUsers = Read-Host "Separate multiple LoginIDs with commas"
	Clear-Host

	# Input username(s) of target(s)
	Write-Host "Enter the LoginID(s) of the user(s) who's calendar you"
	Write-Host "want to grant access to."
	$strSubjects = Read-Host "Separate multiple LoginIDs with commas"
	Clear-Host

	# Input access type
	$strUsers = Clean-and-Split($strUsers)
	$strSubjects = Clean-and-Split($strSubjects)
	Write-Host "Enter a number for the permission level you wish to set"
	Write-Host "To stop meetings appearing as BUSY you probably want Reviewer"
	Write-Host ""
	Write-Host "1 - Owner (Full Access)"
	Write-Host "2 - Editor (Read Write - Edit All)"
	Write-Host "3 - Author (Read Write - Edit Own)"
	Write-Host "4 - Reviewer (Read Only)"
	Write-Host "5 - Contributor (Write Only)"
	$strLevel = Read-Host ">"
	If ($strLevel -eq 1){
		$strLevelWord = "Owner"
	}
	ElseIf ($strLevel -eq 2){
		$strLevelWord = "Editor"
	}
	ElseIf ($strLevel -eq 3){
		$strLevelWord = "Author"
	}
	ElseIf ($strLevel -eq 4){
		$strLevelWord = "Reviewer"
	}
	ElseIf ($strLevel -eq 5){
		$strLevelWord = "Contributor"
	}
	else {
	Clear-Host
	Write-Host "Invalid choice - Press return to close"
	$Done = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	$host.SetShouldExit(0)
	exit
	}
	Clear-Host

	#Grant access
	Foreach ($strSubject in $strSubjects) {
		Foreach ($strUser in $strUsers) {
		$strSubjectget = Get-Mailbox -Identity $strSubject
		$strSubjectSMTP = $strSubjectget.PrimarySmtpAddress
		$strSubjectstring = ($strSubjectSMTP + ":\calendar")
		Add-MailboxFolderPermission -Identity $strSubjectstring -User $strUser -AccessRights $strLevelWord
		}
	}
	Clear-Host
	Write-Host "All permissions have been applied. Press return to close."
	$Done = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	$host.SetShouldExit(0)
	exit
}

# Choice 2 - Remove Access
ElseIf ($strChoice -eq 2){

	# Take username(s) of users to remove rights for
	Clear-Host
	Write-Host "Enter the LoginID(s) of the user(s) who's rights you wish to remove."
	$strUsers = Read-Host "Separate multiple LoginIDs with commas"
	Clear-Host
	
	# Take username(s) of target(s)
	Write-Host "Enter the LoginID(s) of the user(s) who's calendar you"
	Write-Host "want to remove access from."
	$strSubjects = Read-Host "Separate multiple LoginIDs with commas"
	Clear-Host
	
	#Strip rights
	$strUsers = Clean-and-Split($strUsers)
	$strSubjects = Clean-and-Split($strSubjects)
	Clear-Host
	Foreach ($strSubject in $strSubjects) {
		Foreach ($strUser in $strUsers) {
		$strSubjectget = Get-Mailbox -Identity $strSubject
		$strSubjectSMTP = $strSubjectget.PrimarySmtpAddress
		$strSubjectstring = ($strSubjectSMTP + ":\calendar")
		Remove-MailboxFolderPermission -Identity $strSubjectstring -User $strUser -Confirm:$false
		}
	}
	Clear-Host
	Write-Host "All permissions have been removed. Press return to close."
	$Done = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	$host.SetShouldExit(0)
	exit
}

#Failure routine. Invalid choices
else {
Clear-Host
	Write-Host "Invalid choice selected, aborting. Press return key to close"
	$Done = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	$host.SetShouldExit(0)
	exit
}
