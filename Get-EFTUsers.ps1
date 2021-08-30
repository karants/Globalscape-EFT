#Load SFTP COM object
$SFTPServer = New-Object -COM "SFTPCOMInterface.CIServer"

#Auth Parameters
$ComputerName = "localhost"
$ServicePort = 1000
$Username = ""
$securestring = (Get-Content "E:\Scripts\key.txt" | ConvertTo-SecureString)
$Passkey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securestring))

#Connection String
$SFTPServer.Connect($ComputerName, $ServicePort, $Username, '')

#Get List of Sites
$SFTPSites = $SFTPServer.Sites()

#Initialize Users Array
$UserList = @()

#Select the first site and Get Users
$site = $SFTPSites.Item(0)

$users = $site.GetUsers()

#For each user, get the associated attributes and store them in the $UserList array
    
foreach ($user in $users) {

            $UserSettings = $site.GetUserSettings($user)
            $Enabled = $null
            $userObj = "" | Select Username,IsEnabled,Email,HomeDirectory,LastConnectionTime,AccountCreationTime,LastModificationTime
            $userObj.Username = $user
            $userObj.Email = $UserSettings.Email
            $userObj.IsEnabled = $UserSettings.GetEnableAccount([ref]$Enabled)
            $userObj.HomeDirectory = $UserSettings.GetHomeDirString()
            $userObj.LastConnectionTime = $UserSettings.LastConnectionTime
            $userObj.AccountCreationTime = $UserSettings.AccountCreationTime
            $userObj.LastModificationTime = $UserSettings.LastModificationTime
            $UserList += $userObj #Add the current user details to the array
        }
    
#Export the array as a CSV (Temporary)
$UserList | Export-CSV -Path "\\..\EFT\users.csv" -NoTypeInformation

#Define the header and CSS to be used within the HTML Body of the Webpage
 $header = @"
<meta http-equiv="Content-type" content="text/html; charset=utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no">
<link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta1/dist/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.10.24/css/dataTables.bootstrap5.min.css">
<link rel="stylesheet" type="text/css" href=https://cdnjs.cloudflare.com/ajax/libs/semantic-ui/2.3.1/semantic.min.css">
<link rel="stylesheet" type="text/css" href=https://cdn.datatables.net/1.10.24/css/dataTables.semanticui.min.css">
<script type="text/javascript" language="javascript" src="https://code.jquery.com/jquery-3.5.1.js"></script>
<script type="text/javascript" language="javascript" src="https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js"></script>
<script type="text/javascript" language="javascript" src="https://cdn.datatables.net/1.10.24/js/dataTables.bootstrap5.min.js"></script>
<script type="text/javascript">
`$(document).ready(function() {
    `$('#userlist').DataTable();
} );
</script>
"@

#Import CSV and Export to HTML 

$html = Import-CSV "\\..\EFT\users.csv" | ConvertTo-Html -Head $header -Body "<center><h1>EFT Users Status</h1>`n<h5>Generated on $(Get-Date)</h5></center><br>" 

$htmlpage = $html -replace "<table>","<table id=""userlist"" class=""ui celled-table"" style=""width:100%"">" `
                  -replace "<colgroup>", "" `
                  -replace "<col/>","" `
                  -replace "</colgroup>", "<thead>" `
                  -replace "</th></tr>","</th></tr></thead><tbody>" `
                  -replace "</table>","</tbody></table>"

$htmlpage | Out-File "\\..\users.html"


#Remove-Item -Path "\\..\EFT\users.csv" -Force

$SFTPServer.Close()