<# Pre-reqs #>
Set-ExecutionPolicy Unrestricted

<# Variables are defined in this section #>
# Set variables
[int]$loginLimit = 3
[int]$loginRound = 1
[str]$findAgain = $null
[int]$loginRemaining = $null
<#---------------------------------------#>
<# Functions are defined in this section #>
# Loops function 'findUser'
function findUserAgain {
    Write-Host "---------------------------------------------"
    Write-Host "Choices:"
    Write-Host "Y = perform another lookup."
    Write-Host "N = quit this application."
    While ("y", "Y", "n", "N" -notcontains $findAgain){
        $findAgain = Read-Host -Prompt "Find another? (y/n)"
    }
    write-host "loop find is $findAgain"
    if ($findAgain -like "n") {
        exit
    }
}
# Find user alias
function findUser {
    try {
        $ProxyAddress = Read-Host -Prompt "Input the user's email address"
        Get-Recipient | where {$_.EmailAddresses -match "$ProxyAddress"} | fL Name, RecipientType,emailaddresses
    }
    catch {
        Write-Host "Try again"
    }
}
# Call O365 session handler
function logIn {
    try {
        Write-Host "Begin login attempt $loginRound."
        $UserCredential = Get-Credential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection -ErrorAction Stop
        Import-PSSession $Session -AllowClobber
    }
    catch [System.Exception] {
            While ($loginRound -lt 3) {
                $loginRemaining = $loginLimit - $loginRound
                Write-Warning -Message "You have provided invalid credentials, please try again. $loginRemaining remaining attempt(s)."
                $loginRound++
                Start-Sleep -Milliseconds 1000
                logIn
            }
            badLogIn -ErrorAction silentlyContinue
    }
}
# Quit due to invalid logins
function badLogIn {
        Write-Host "ERROR: Maximum number of login attempts reached. Exiting in 5 seconds." -BackgroundColor Black -ForegroundColor Red
        Start-Sleep -Milliseconds 5000
        Exit
}
<#---------------------------------------#>
<# Run main application #>
Write-Host "

================================
   _____          _____ _    _ 
  / ____|   /\   |_   _| |  | |
 | |       /  \    | | | |  | |
 | |      / /\ \   | | | |  | |
 | |____ / ____ \ _| |_| |__| |
  \_____/_/    \_\_____|\____/ 
                               
================================
          www.caiu.org
--------------------------------
                               "
Write-Host "This tool will check for user aliases in Office 365.
You will first be required to sign in using your a valid Microsoft account.
"
Read-Host "[Press enter to continue]"
Start-Sleep -Milliseconds 300
logIn
Write-Host "Login Successful" -ForegroundColor Green
findUser
Write-Host "End of results!" -ForegroundColor DarkYellow
findUserAgain
Do {
    findUserAgain
    } until ($findAgain -like "n")
write-host "Bye!"
Start-Sleep -Milliseconds 1000
exit
