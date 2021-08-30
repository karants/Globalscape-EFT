#Load SFTP COM object
$SFTPServer = New-Object -COM "SFTPCOMInterface.CIServer"

#Auth Parameters
$ComputerName = "localhost"
$ServicePort = 1000
$Username = ""

#Date
$Today = Get-Date

#Connection String
$SFTPServer.Connect($ComputerName, $ServicePort, $Username, '')

#Get List of Sites
$SFTPSites = $SFTPServer.Sites()

#Initialize Users Array
$UserList = @()

$securestring = (Get-Content "..\key.txt" | ConvertTo-SecureString)
$Passkey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securestring))


#Select the first site and Get Users
$site = $SFTPSites.Item(0)

$users = $site.GetSettingsLevelUsers('Default')

#For each user, compare the dates and delete if they have not connected in the past 180 days and have not been created in the past 14 days
Write-host "Removing Users That Have Not Connected in the past 180 days):\n\n"
Write-Host "UserName -- Last Connection Time"
foreach ($user in $users) {

            $UserSettings = $site.GetUserSettings($user)
         
			if($UserSettings.LastConnectionTime -lt $Today.AddDays(-180)){
				
				if($UserSettings.AccountCreationTime -lt $Today.AddDays(-14)){
			
					write-host $user " -- " $UserSettings.LastConnectionTime
					$site.removeuser($user)		
				}
			}
        }
    

$SFTPServer.Close()
Read-Host -Prompt “Press Enter to exit”