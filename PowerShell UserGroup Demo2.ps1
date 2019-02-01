# create a new team under an existing group
# sample by https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/team_put_teams

# 0. bevor wir beginnen
# um ein AuthToken zu bekommen, müssen wir zuerst eine App in AzureAD registrieren.
# da lesen wir die AppID ab und verwenden sie hier
$client_id     = '12345678-abcd-efgh-ihkl-mn1234512345'   
$client_secret = 'SJ3C1Uiva3zzYDHJ4vRvv0k9t0m99chbYi9s='
$redirectUrl   = 'https://localhost/'

# 1. wir müssen uns anmelden und ein AuthToken erhalten
# https://blogs.technet.microsoft.com/ronba/2016/05/09/using-powershell-and-the-office-365-rest-api-with-oauth/

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Web
$loginUrl = "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code&redirect_uri=" + 
            [System.Web.HttpUtility]::UrlEncode($redirectUrl) + 
            "&client_id=$client_id" + 
            "&prompt=login"

$form = New-Object -TypeName System.Windows.Forms.Form -Property @{Width=440;Height=640}
$web  = New-Object -TypeName System.Windows.Forms.WebBrowser -Property @{Width=420;Height=600;Url=($loginUrl ) }
$DocComp  = {
    $Global:uri = $web.Url.AbsoluteUri
    if ($Global:Uri -match "error=[^&]*|code=[^&]*") {$form.Close() }
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
    
$AuthorizationPostRequest = 
    "grant_type=authorization_code" + "&" +
    "redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) + "&" +
    "client_id=$client_id" + "&" +
    "client_secret=" + [System.Web.HttpUtility]::UrlEncode("$client_secret") + "&" +
    "code=" + $queryOutput["code"] + "&" +
    "resource=" + [System.Web.HttpUtility]::UrlEncode("https://graph.microsoft.com/")

$Authorization = Invoke-RestMethod   -Method Post `
                        -ContentType application/x-www-form-urlencoded `
                        -Uri https://login.microsoftonline.com/common/oauth2/token `
                        -Body $AuthorizationPostRequest

$token = $Authorization.access_token

# 2. wir erstellen zunächst die Office 365 group
# https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_post_groups

# POST /groups

$newOffice365groupName = "Name Der Gruppe 1"

$newOffice365GroupRequest = "{ 
    description: ""per Microsoft Graph erstellt"", 
    displayName: ""NameDerGruppe"",
    groupTypes: [ ""Unified"" ],
    mailEnabled: true, 
    mailNickname: ""AliasDerGruppe"", 
    securityEnabled: false    
}"

$header = @{Authorization = ("Bearer "+$token)} 

$newOffice365Group = Invoke-RestMethod -Uri https://graph.microsoft.com/beta/groups -Headers $header -Method POST -Body $newOffice365GroupRequest -ContentType "application/json"   -Verbose

# HTTP request

# just wait for debugging
$waitForInput = read-host -Prompt "press enter!"

# PUT /groups/{id}/team - ?PUT, weil wir ja ein bestehendes Element einfügen?
# content-type: application/jscn
# body <- ist noch ein json-String, den ich zerlegen muss - Berechtigungen etc

$groupid = $newOffice365Group.id

$newTeamProperties = "{  
    ""memberSettings"": {
      ""allowCreateUpdateChannels"": true,
      ""allowDeleteChannels"": true,
      ""allowAddRemoveApps"": true,
      ""allowCreateUpdateRemoveTabs"": true,
      ""allowCreateUpdateRemoveConnectors"": true    
    },
    ""guestSettings"": {
      ""allowCreateUpdateChannels"": true,
      ""allowDeleteChannels"": true 
    },
    ""messagingSettings"": {
      ""allowUserEditMessages"": true,
      ""allowUserDeleteMessages"": true,
      ""allowOwnerDeleteMessages"": true,
      ""allowTeamMentions"": true,
      ""allowChannelMentions"": true    
    },
    ""funSettings"": {
      ""allowGiphy"": true,
      ""giphyContentRating"": ""strict"",
      ""allowStickersAndMemes"": true,
      ""allowCustomMemes"": true
    }
  }"

$header = @{Authorization = ("Bearer "+$token)}
$newTeamCreated = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/{$groupid}/team" -Headers $header -Method PUT -ContentType "application/json" -Body $newTeamProperties  -Verbose

# just wait for debugging
$waitForInput = read-host -Prompt "press enter!"

# clean it up
$groupid = $newOffice365Group.id
$deleteGroupAtTheEnd = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/{$groupid}/" -Method DELETE -Headers $header
