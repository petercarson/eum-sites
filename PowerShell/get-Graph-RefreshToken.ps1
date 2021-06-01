#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Add System Web Assembly to encode ClientSecret
Add-Type -AssemblyName System.Web

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Function Get-AuthCode {
	Add-Type -AssemblyName System.Windows.Forms

	$form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
	$web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($url -f ($Scope -join "%20")) }

	$DocComp  = {
		$Global:uri = $web.Url.AbsoluteUri        
		if ($Global:uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
	}
	$web.ScriptErrorsSuppressed = $true
	$web.Add_DocumentCompleted($DocComp)
	$form.Controls.Add($web)
	$form.Add_Shown({$form.Activate()})
	$form.ShowDialog() | Out-Null

	$queryOutput = [System.Web.HttpUtility]::ParseQueryString($web.Url.Query)
	$output = @{}
	foreach($key in $queryOutput.Keys){
		$output["$key"] = $queryOutput[$key]
	}

	$output
}

Write-Host "Enter Service Principal Client ID and Secret for the Azure subscription"
$Credential = Get-Credential
$clientId = $Credential.Username
$clientSecret = $Credential.GetNetworkCredential().Password

$RedirectUrl = "https://localhost:8000"
$ResourceUrl = "https://graph.microsoft.com"

# UrlEncode the ClientID and ClientSecret and URL's for special characters 
$clientIDEncoded = [System.Web.HttpUtility]::UrlEncode($ClientId)
$clientSecretEncoded = [System.Web.HttpUtility]::UrlEncode($ClientSecret)
$redirectUrlEncoded =  [System.Web.HttpUtility]::UrlEncode($RedirectUrl)
$resourceUrlEncoded = [System.Web.HttpUtility]::UrlEncode($ResourceUrl)
$scopeEncoded = [System.Web.HttpUtility]::UrlEncode("https://outlook.office.com/user.readwrite.all")

# Get AuthCode
$url = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=$redirectUrlEncoded&client_id=$clientID&resource=$resourceUrlEncoded&scope=$scopeEncoded"

Get-AuthCode
	
# Extract Access token from the returned URI
$regex = '(?<=code=)(.*)(?=&)'
$authCode = ($uri | Select-string -pattern $regex).Matches[0].Value

# Get Authentication Token and Refresh Token
$body = "grant_type=authorization_code&redirect_uri=$redirectUrl&client_id=$clientId&client_secret=$clientSecretEncoded&code=$authCode&resource=$resourceUrl"
$Authorization = Invoke-RestMethod https://login.microsoftonline.com/common/oauth2/token `
	-Method Post -ContentType "application/x-www-form-urlencoded" `
	-Body $body `
	-UseBasicParsing


# Store refreshToken
Set-Content $PSScriptRoot"\refreshToken.txt" $Authorization.refresh_token