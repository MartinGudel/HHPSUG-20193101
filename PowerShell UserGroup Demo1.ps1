# based on https://docs.microsoft.com/de-de/onedrive/developer/rest-api/?view=odsp-graph-online

# Add-Type -AssemblyName System.Windows.Forms
# Add-Type -AssemblyName System.Web

# # client_id und client_secret aus der AzureAD App Registration ablesen
$client_id     = '36f8e707-f066-44dd-bb26-41ea48517546'   
$client_secret = 'SJ3C1Uiva3zzYDHJ4vRv/fk17iHaIvLolN9VHpbYi9s='
$redirectUrl   = 'https://localhost/'

# wo finde ich die TenantID? https://docs.microsoft.com/en-us/onedrive/find-your-office-365-tenant-id
$tokenendpoint = 'https://login.microsoftonline.com/Hier-Bitte-Eure-TenantID/oauth2/token'
$username = "powershelldemouser@teamsplayer.de"
$password = 'ZuSimplesPassword'

$api = 'https://graph.microsoft.com/'

$AuthorizationPostRequest = "grant_type=password" + "&" +
                                "username=" + $username + "&" +
                                "password=" + $password + "&" +
                                "redirect_uri=" + [System.Web.HttpUtility]::UrlEncode($redirectUrl) + "&" +
                                "client_id=$client_id" + "&" +
                                "client_secret=" + [System.Web.HttpUtility]::UrlEncode("$client_secret") + "&" +
                                "resource=" + [System.Web.HttpUtility]::UrlEncode($api)

$Authorization = Invoke-RestMethod -Method Post `
                        -ContentType application/x-www-form-urlencoded `
                        -Uri  $tokenendpoint `
                        -Body $AuthorizationPostRequest

$Token = $Authorization.access_token

$idOrUserPrincipalName = "powershelldemouser@teamsplayer.de"
$resource = "https://graph.microsoft.com/v1.0/users/$idOrUserPrincipalName/drive"

$header = @{'Authorization' = ("Bearer $token")}
$body = $null
$result = Invoke-RestMethod -Uri $resource -Headers $header -Method GET -ContentType "application/json"  -Verbose

# get drive for $sourceUser - needs Sharepoint Libray permissions first
$sourceUser = "powershelldemouser@teamsplayer.de"
$sourceResource = "https://graph.microsoft.com/v1.0/users/$sourceUser/drive/root"
$sourceRoot = Invoke-RestMethod -Uri $sourceResource -Headers $header -Method GET -ContentType "application/json"  -Verbose

Connect-MicrosoftTeams -AadAccessToken $Token -AccountId "powershelldemouser@teamsplayer.de"   -Verbose
