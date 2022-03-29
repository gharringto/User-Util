Clear
$ClientDomain = Read-Host "Are we on a client's domain? Y/N"
If ($ClientDomain -ne "Y") {$ClientDomain = $False} Else {$ClientDomain = $True}
$Client = New-Object -TypeName PSObject
$Client | Add-Member -NotePropertyMembers @{DisplayName=""; PC=""; MSG=@(); AD=""; ADG=@()} -TypeName Asset
$Admin = New-Object -TypeName PSObject
$Admin | Add-Member -NotePropertyMembers @{DisplayName=""; PC=""; MS=""; MSCredentials=""; AD=""} -TypeName Asset
$User = New-Object -TypeName PSObject
$User | Add-Member -NotePropertyMembers @{DisplayName=""; PC=""; MS=""; MSG=""; AD=""; ADG=@(); Licenses=""} -TypeName Asset
$ConnectedDomains = Get-MsolDomain -ErrorAction SilentlyContinue
$AutoLocate = "blank"
$AutoTicket = "blank"

Function Global:Print-CMDs {
Write-Host "Located the following commands for the utility:"
Write-Host "-----------------------------------------------"
Write-Host "  Print-CMDs"
Write-Host "  Write-Notes"
Write-Host "  Get-Started !"
Write-Host "! Check-PSModules"
Write-Host "! Get-Client-Info"
Write-Host "! Get-User-Info"
Write-Host "! Print-User"
Write-Host "  Offboard-User *"
Write-Host "* User-EXOMailbox-HideFromGAL"
Write-Host "* User-EXOMailbox-Shared"
Write-Host "* User-EXOMailbox-AddDelegate"
Write-Host "* User-MS-ClearGroups"
Write-Host "* User-MS-Disabled"
Write-Host "  User-MS-Enabled"
Write-Host "* User-MS-PWRandom"
Write-Host "* User-MS-ResetMFA"
Write-Host "* User-MS-DisableMFA"
Write-Host "* User-MS-ClearLicenses"
Write-Host "* User-AD-Description"
Write-Host "* User-AD-Disabled"
Write-Host "  User-AD-Enabled"
Write-Host "* User-AD-PWRandom"
Write-Host "* User-AD-ClearGroups"
Write-Host "* User-AD-HideFromGAL"
Write-Host "* User-AD-MvToRetntnOU"
Write-Host "* Client-AD-DirSync"
}

Function Global:Write-Notes {
<#
	.SYNOPSIS
		Takes pipeline or written input and appends to a file.
	.DESCRIPTION
        	This Function Global:will take mutiple or single pipleline inputs and or written input and appends to a folder/file path that you (optionally) can define.
        	-PipeLine Value: (optional) Takes one or more values stored from the Pipeline and stores it to a file.
        	-Message: (optional) Creates a message you define
            	You can write multiple messages at once with , (Get-Help #Write-Notes -Examples) for more information.
        	-FolderPath: (optional) Default folder path: $env:Userprofile\desktop
        	-FileName: (optional) Default file Name: Offboarding Notes.txt
	.EXAMPLE
		Pipeline Example: Get-WmiObject -Class Win32_Processor -ComputerName $env:COMPUTERName | #Write-Notes -FileName CPUInfo -FilePath $env:Userprofile\Desktop\
	Written  Example: #Write-Notes -Message "$env:COMPUTERName system information", "Home Drive: $env:HOMEDRIVE", "Logon Server: $env:LOGONSERVER" -FilePath $env:UserPROFILE\Desktop -FileName "$env:COMPUTERName Information"
	.NOTES	FunctionName : #Write-Notes	Created by : Zach Hudson	Date Coded : 03/11/2021		Modified by : Zach Hudson	Date Modified : 03/13/2021
#>
	[CmdletBinding()]
	param (
		[Parameter(ValueFromPipeline=$true)]
		$pipeValue,
		[Parameter()]
		$Message,
		[Parameter()]
		$FilePath = "$env:UserPROFILE\Desktop",
		[Parameter()]
		$FileName = "$User Offboarding Notes"
	)
    	begin {
		<# Process each value and get them ready as 1 unit for the process block.
	If no begin block, they'll be processed as single items. #>
	}
	process {
		<# Take each item in the pipeline and write it out to a file. #>
		ForEach ($pValue in $pipeValue) {
			$pValue | Out-File "$FilePath\$FileName.txt" -Append
		}
	}
	end {
		<# Take the -message parameter and append the message to the file. #>
		$Message | Out-File "$FilePath\$FileName.txt" -Append
	}
}

#Connect to Partner Center
Function Global:Get-Started {
        Try {
            Check-PSModules
            If (!(Get-PartnerContext)) {Connect-PartnerCenter ; Write-Host }
            Get-Client-Info
            Get-User-Info
        }
            Catch {
            $Error[0]
        }
}

Function Global:Check-PSModules {
	Write-Host "Checking for required modules..."
	ForEach ($Module in @("PowerShellGet","PartnerCenter","MSOnline","AzureAD","ExchangeOnlineManagement","Microsoft.Online.SharePoint.PowerShell","MicrosoftTeams")) {
		$PSModule = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue
		If ($null -eq $PSModule) {
			Try {
				Install-Module $Module  -Force -AllowClobber
				### Write-Notes  -message "Observed missing powershell module, installed" $Module "module..."
				Write-Host "Check-PSModules: Installed $Module."
			}
			Catch {
				$Error[0]
			}
		}
        Else {Update-Module $Module}
    	}
	Write-Host "Check-PSModules: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
}

#Connect to Microsoft Online
#Connect to Exchange Online
Function Global:Get-Client-Info {
<#
	.SYNOPSIS
		Grabs the relevant Client and admin info from AD if not at work, then from PC, then from PC Tenant.
	.NOTES	Created by : Gregory Harrington 	Date Coded : 2/24/2022	Date Updated : 3/5/2022
#>
    Try{
        #Wipe any occurances of Client variables.
        $Client.DisplayName = $Null
        $Client.PC = $Null
        $Client.MSG = $Null
        $Client.AD = $Null
        $Client.ADG = $Null
        $Admin.DisplayName = $Null
        $Admin.PC = $Null
        $Admin.MS = $Null
        $Admin.MSCredentials = $Null
        $Admin.AD = $Null
        #Active Directory Magic!
        If ($ClientDomain -eq $True) {
            #Setup some default values.
            $Client.AD = $env:UserDNSDOMAIN
            $Client.DisplayName = $env:UserDOMAIN
            #Grab the Active Directory Administrator.
            Do {
                $Set = $False
                If (!$Admin.AD) {$Admin.AD = Get-AdUser -Filter * | Where-Object {$_.UserPrincipalName -like "*admin*"}  | Sort-Object -Property UserPrincipalName }
                If ($Admin.AD.Length -gt 1) {Write-Host "Please make a selection for Client's AD Admin:" }
                If ($Admin.AD.Length -eq 1) {Write-Host "Located the following for Client's AD Admin:" }
                $Admin.AD | Format-List DisplayName, UserPrincipalName
                If ($Admin.AD.Length -gt 1) {
                    Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $Admin.AD.Count
                    Write-Host
                        If (($Sel -le $Admin.AD.Length) -and ($Sel -gt 0)) {
                            $Admin.AD = $Admin.AD[$Sel-1]
                            $Set = $true
                            Write-Output "Selected:" $Admin.AD.Name
			            }
                        Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
		            }
		            Until ($Set -eq $True)
	            }
            If (!$Admin.AD) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
            }
            #Loop this until we have a good one.
            Until ($Admin.AD)
        }
        #Grab the Client's Microsoft Tenant.
	    Do {
            $Set = $False
            #Use the Active Directory NetBios Name, if available.
            If ($Client.DisplayName.Length -gt 4) {
                $Search = $Client.DisplayName
                $Client.PC = Get-PartnerCustomer | Where-Object {$_.Name -like "*$Search*"} | Sort-Object -Property Name
            }
            #Use the Active Directory Administrator's email domain Name if, available and needed.
            If ((!$Client.PC) -and ($Admin.AD)) {
                $Search = $Admin.AD.UserPrincipalName.Split('@')[1].Split('.')[0]
       		    $Client.PC = Get-PartnerCustomer | Where-Object {$_.Name -like "*$Search*"} | Sort-Object -Property Name
            }
            #Use the host operator, if needed.
            If (!$Client.PC) {
                $Search = Read-Host "What Client are we working with?"
       		    $Client.PC = Get-PartnerCustomer | Where-Object {$_.Name -like "*$Search*"} | Sort-Object -Property Name
            }
            #Determine if we have multiple entries for Client's Microsoft Tenant.
		    If ($Client.PC.Length -gt 1) {Write-Host "Please make a selection for Client's M365 Tenant:" }
            #If we don't then use the proper sentance instead.
            If ($Client.PC.Length -eq 1) {Write-Host "Located the following for Client's M365 Tenant:" }
            #Display our option(s).
            $Client.PC | Format-List Name, Domain, CustomerID
            #Handle the selection if multiple entries.
		    If ($Client.PC.Length -gt 1) {
                Do {
				    [int]$Sel = Read-Host "Please select a record, 1 -" $Client.PC.Length
				    Write-Host
				    If (($Sel -le $Client.PC.Length) -and ($Sel -gt 0)) {
					    $Client.PC = $Client.PC[$Sel-1]
                        $Set = $true
                        Write-Host "Selected:" $Client.PC.Name
		    	    }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
			    } 
			    Until ($Set -eq $True)
		    }
            #Waterfall a variable.
            If ($Client.PC) {$Client.DisplayName = $Client.PC.Name}
            #Notify if any issues.
            If (!$Client.PC) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
		}
        #Loop this until we have a good one.
		Until ($Client.PC)
        #Grab the Client's Partner Center Tenant Administrator.
        Do {
            $Set = $False
            #Use our best guess, not any host operator.
            If (!$Admin.PC) {$Admin.PC = Get-PartnerCustomerUser $Client.PC.CustomerID | Where-Object {$_.UserPrincipalName -like "*admin*"}  | Sort-Object -Property UserPrincipalName }
            #Determine if we have multiple entries for Client's Microsoft Tenant Administrator.
            If ($Admin.PC.Length -gt 1) {Write-Host "Please make a selection for Client's M365 Global Admin:" }
            #If we don't then use the proper sentance instead.
            If ($Admin.PC.Length -eq 1) {Write-Host "Located the following for Client's M365 Global Admin:" }
            #Display our option(s).
            $Admin.PC | Format-List DisplayName, UserPrincipalName
            #Handle the selection if multiple entries.
            If ($Admin.PC.Length -gt 1) {
                Do {
                [int]$Sel = Read-Host "Please select a record, 1 -" $Admin.PC.Length
                Write-Host
                    If (($Sel -le $Admin.PC.Length) -and ($Sel -gt 0)) {
                        $Admin.PC = $Admin.PC[$Sel-1]
                        $Set = $True
                        Write-Host "Selected:" $Admin.PC.DisplayName
 	                }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
                } 
                Until ($Set -eq $True)
            }
            #Waterfall a variable.
            If ($Admin.PC) {$Admin.DisplayName = $Admin.PC.Name}
            #Notify if any issues.
            If (!$Admin.PC) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
            }
        #Loop this until we have a good one.
        Until ($Admin.PC)
        #Check if we're previously connected, reduces lag and moar.
        If ($ConnectedDomains) {
            $Set = $False
            #See if any available Microsoft Online domains belong to the Client's Microsoft Tenant Administrator.
            ForEach ($Domain in $ConnectedDomains) { If ($Domain.Name -eq $Admin.PC.UserPrincipalName.Split('@')[1]) { $Set = $True } }
            #Log us out if we're not connected to any domains that belong to the Client's Microsoft Tenant Administrator.
            If ($Set -eq $False ) {
                [Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()
                $ConnectedDomains = $Null
                Get-PSSession | Remove-PSSession
            }
        }
        #Check if we're not previously connected.
        If (!($ConnectedDomains)) {
            Write-Host "Please sign into Microsoft 365 with the Client's Gloal Admin..."
            Write-Host "HINT: You can paste the selected admin email address right now."
            $Admin.PC.UserPrincipalName | Set-Clipboard
            Connect-MsolService
            Write-Host
            Write-Host "Please sign into Exchange Online with the Client's Gloal Admin..."
            Connect-ExchangeOnline -UserPrincipalName $Admin.PC.UserPrincipalName
        }
        
        #Grab the Client's Partner Center Tenant Administrator's Microsoft account.
        Do {
            $Set = $False
            #Use the Client's Partner Center Tenant Administrator.
            $Admin.MS = Get-MsolUser -UserPrincipalName $Admin.PC.UserPrincipalName | Sort-Object -Property Name
            #Handle the selection if multiple entries (which should actually never be possible).
            If ($Admin.MS.Length -gt 1) {Write-Host "Please make a selection for admin's M365 profile:"
    		    $Admin.MS | Format-List DisplayName, UserPrincipalName
                Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $Admin.MS.Length
                    Write-Host
                    If (($Sel -le $Admin.MS.Length) -and ($Sel -gt 0)) {
                        $Admin.MS = $Admin.MS[$Sel-1]
                        $Set = $true
                        Write-Host "Selected" $Admin.MS.DisplayName
    	            }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	            } 
	            Until ($Set -eq $True)
		    }
            #Notify if any issues.
            If (!$Admin.MS) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    }
        #Loop this until we have a good one.
	    Until ($Admin.MS)
        #Grab more information from Partner Center Tenant 
        Write-Host "Getting M365 Groups..."
        $Client.MSG = Get-MSOLGroup -All | Where {$_.LastDirSyncTime -eq $Null}
        ### Write-Notes -message "Located admin acount: $Admin.AD.Name"
        Write-Host "Get-Client-Info: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host 
    }
    Catch {
        Write-Output $Error[0]
    }
}

Function Global:Get-User-Info {
<#
	.SYNOPSIS
		Grabs the relevant User info from AD if not at work, then from PC Tenant.
	.NOTES	Created by : Gregory Harrington 	Date Coded : 2/24/2022	Date Updated : 3/5/2022
#>
    Try {
        #Wipe any occurances of Client variables.
        $User.DisplayName = $Null
        $User.PC = $Null
        $User.MS = $Null
        $User.MSG = @()
        $User.AD = $Null
        $User.ADG = @()
        $User.Licenses = $Null
        #Grab the Client's Partner Center Tenant User.
        Do {
            $Set = $False
            #No assumptions.$auto
            If ($AutoLocate) {$Search = $AutoLocate ; $AutoLocate = $Null}
            Else {$Search = Read-Host "What User are we working with?"}
            If ($Search -eq "") {Write-Host ; Return}
            $User.PC = Get-PartnerCustomerUser $Client.PC.CustomerID | Where-Object {$_.DisplayName -like "*$Search*"}  | Sort-Object -Property Name
            #Handle the selection if multiple entries.
            If ($User.PC.Length -gt 1) {
                Write-Host "Please make a selection for User's Partner Center profile:"
                $User.PC | Format-List DisplayName, UserPrincipalName
                Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $User.PC.Length
                    Write-Host
                    If (($Sel -le $User.PC.Length) -and ($Sel -gt 0)) {
                        $User.PC = $User.PC[$Sel-1]
                        $Set = $true
                        Write-Host "Selected" $User.PC.DisplayName
    	            }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	            } 
	            Until ($Set -eq $True)
		    }
            #Be sure we aren't selecting our Administrator.
            If ($User.PC.UserPrincipalName -eq $Admin.PC.UserPrincipalName) {
                Write-Host "User cannot be Admin!" -ForegroundColor Red -BackgroundColor White
                $User.PC = $Null
            }
            #Waterfall a variable.
            If ($User.PC) {$User.DisplayName = $User.PC.Name}
            #Notify if any issues.
            If (!$User.PC) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    }
        #Loop this until we have a good one.
	    Until ($User.PC)
        #Grab the Client's Partner Center Tenant User's Microsoft account.
        Do {
            $Set = $False
            #Use the Client's Partner Center Tenant User.
            $User.MS = Get-MsolUser -UserPrincipalName $User.PC.UserPrincipalName | Sort-Object -Property Name
            #Handle the selection if multiple entries (which should actually never be possible).
            If ($User.MS.Length -gt 1) {Write-Host "Please make a selection for User's M365 profile:"
    		    $User.MS | Format-List DisplayName, UserPrincipalName
                Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $User.MS.Length
                    Write-Host
                    If (($Sel -le $User.MS.Length) -and ($Sel -gt 0)) {
                        $User.MS = $User.MS[$Sel-1]
                        $Set = $true
                        Write-Host "Selected" $User.MS.DisplayName
    	            }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	            } 
	            Until ($Set -eq $True)
		    }
            #Be sure we aren't selecting our Administrator (which should actually never be possible at this point).
            If ($User.MS.UserPrincipalName -eq $Admin.MS.UserPrincipalName) {
                Write-Host "User cannot be Admin!" -ForegroundColor Red -BackgroundColor White
                $User.MS = $Null
            }
            #Notify if any issues.
            If (!$User.MS) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    }
        #Loop this until we have a good one.
	    Until ($User.MS)
        #Active Directory Magic!
        If ($ClientDomain -eq $True) {
       	    Do {
                $Set = $False
                #Use the Client's Partner Center Tenant User's Microsoft account.
                $User.AD = Get-ADUser -Filter * | Where-Object {$_.UserPrincipalName -like $User.MS.UserPrincipalName} | Sort-Object -Property Name
                #Use a fuzzy search of Client's Partner Center Tenant User's Microsoft account, if needed.
                If (!$User.AD) {
		            $Search = $User.MS.DisplayName
		            $User.AD = Get-ADUser -Filter * | Where-Object {$_.Name -like "*$Search*"} | Sort-Object -Property Name
                }
                #Use the host operator, if needed.
                If (!$User.AD) {
		            $Search = Read-Host "What User are we working with?"
		            $User.AD = Get-ADUser -Filter * | Where-Object {$_.Name -like "*$Search*"} | Sort-Object -Property Name
                }
                #Determine if we have multiple entries for Client's Microsoft Tenant.
                If ($User.AD.Length -gt 1) {Write-Host "Please make a selection for Users's AD profile:" }
                #If we don't then use the proper sentance instead.
                If ($User.AD.Length -eq 1) {Write-Host "Located the following for User's AD profile:" }
                #Display our option(s).
    		    $User.AD | Format-List Name,UserPrincipalName
                #Handle the selection if multiple entries.
	    	    If ($User.AD.Length -gt 1) {
		    	    Do {
   			    		[int]$Sel = Read-Host "Please select a record, 1 -" $User.AD.Length
    			    	Write-Host
	    			    If (($Sel -le $User.AD.Length) -and ($Sel -gt 0)) {
		       			    $User.AD = $User.AD[$Sel-1]
    				    }
                        Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    		    }
                    Until ($Set -eq $True)
		        }
                #Notify if any issues.
                If (!$User.AD) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
    	    }
            #Loop this until we have a good one.
	        Until ($User.AD)
            #Write-Notes -Message "Saved copy of Active Directory Groups $env:Userprofile\desktop\$User ADGroups.txt"
        }
        #Require verification to proceed, wipe items if not verified.
        Write-Host "Please VERIFY you want to select:" $User.PC.DisplayName -ForegroundColor Red -BackgroundColor White
        #Act on input, clear if not verified.
        If ((Read-Host "Type YES") -ne "YES" ) {
            $User.DisplayName = $Null
            $User.PC = $Null
            $User.MS = $Null
            $User.Licenses = $Null
            Write-Host "Get-User-Info: Cancelled" -BackgroundColor White -ForegroundColor Red ; Write-Host
            Throw "The operator did not verify the User object as needed, no selection was made."
        }
        #At this point we must be verified (which should actually always be true at this point).
        Else {
            Write-Host "VERIFIED!" -BackgroundColor White -ForegroundColor Blue 
            #Set Texas time on password time stamp.
            $User.MS.LastPasswordChangeTimestamp = ($User.MS.LastPasswordChangeTimestamp) - (New-TimeSpan -Hours 6)
            #Grab more information from Partner Center Tenant
            $Count = $Client.MSG.Length 
            If ($Count -lt 500) {
                Write-Host "Checking $Count Microsoft Groups..."
                ForEach ($Group in $Client.MSG) {If (Get-MsolGroupMember -All -GroupObjectId $Group.ObjectId | Where {$_.Emailaddress -eq $User.MS.UserPrincipalName } ) {$User.MSG += $Group.ObjectId} }
                If ($User.MSG) {$User.MSG = $User.MSG | Sort-Object }
            }
            Else {Write-Host "Too many groups ($Count)! You must remove groups manually from M365!..." ; Pause}
            #Active Directory Magic!
            If ($ClientDomain -eq $True) {
                Write-Host "Checking AD Groups..."
                #Grab more information from Active Directory. 
                $User.ADG = Get-ADPrincipalGroupMembership -Identity $User.AD.SamAccountName | Select-Object -ExpandProperty Name
                If ($User.ADG) {$User.ADG = $User.ADG | Select -Unique | Sort-Object }

            }
            Write-Host
            Print-User
            Write-Host "Get-User-Info: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
        }
    }
    Catch {
        $Error[0]
    }
}

Function Global:Print-User {
$PWD = ($User.MS.LastPasswordChangeTimestamp) - (New-TimeSpan -Hours 6)
$Now = Get-Date
If ($User.MS.StrongAuthenticationRequirements) {$MFA = $True}
Else {$MFA = $False}
Write-Host "Located the following information for the User:"
Write-Host "-----------------------------------------------"
Write-Host "    Name:" $User.PC.DisplayName
Write-Host "---MS365:--------------------------------------"
Write-Host "   Email:" $User.PC.UserPrincipalName
#Write-Host "  UserID:" $User.PC.UserID
#Write-Host " Created:" $User.MS.WhenCreated
Write-Host " PW Date:" $PWD
Write-Host " Blocked:" $User.MS.BlockCredential
Write-Host " MFA Req:" $MFA
    If ($User.MS.Licenses) { ForEach ($License in $User.MS.Licenses) { Write-Host " License:" $License.AccountSkuId.Split(':')[1] } }
    Else { Write-Host " License: Unlicensed" }
    If ($User.MSG) {
        ForEach ($Entry in $User.MSG) {
$Entry = ($Client.MSG | Where {$_.ObjectID -eq $Entry}).DisplayName
Write-Host "   Group:" $Entry
        }
    }
    Else {
Write-Host "MS Group: (N/A)"
    }
    If ($User.AD) {
Write-Host "AciveDir:--------------------------------------"
        ForEach ($Group in $User.ADG) {
Write-Host "   Group:" $Group
        }
    $Location = ($User.AD | Select-Object -ExpandProperty DistinguishedName).Split(',') | Select-Object -Skip 1 #-expandproperty DistinguishedName
Write-Host "AD Local:"
Write-Host $Location
    }
Write-Host
Write-Host "   As of:" $Now
}

Function Global:Offboard-User {
        Try {
            User-EXOMailbox-HideFromGAL
            User-EXOMailbox-Shared
            User-EXOMailbox-AddDelegate
            User-MS-ClearGroups
            User-MS-Disabled
            User-MS-PWRandom
            User-MS-ResetMFA
            User-MS-DisableMFA
            User-MS-ClearLicenses
            If ($ClientDomain -eq $True) {
                User-AD-Description
                User-AD-Disabled
                User-AD-PWRandom
                User-AD-ClearGroups
                User-AD-HideFromGAL
                User-AD-MvToRetntnOU
                Client-AD-DirSync
            }
            Print-User
        }
        Catch {
        $Error[0]
        }
}

Function Global:User-EXOMailbox-HideFromGAL {
    Try {
        Set-Mailbox -Identity $User.MS.UserPrincipalName -HiddenFromAddressListsEnabled $True
        Write-Host "User-EXOMailbox-HideFromGAL: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-EXOMailbox-Shared {
    Try {
        Set-Mailbox -Identity $User.MS.UserPrincipalName -Type Shared
        Write-Host "User-EXOMailbox-Shared: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-EXOMailbox-Forward {
    Try {
            Do {
            $Delegate = $Null
            Do {
            If (!$Delegate) {
                $Search = Read-Host "Who needs the forwarded emails?"
                If ($Search -eq "") {Write-Host ; Return}
                $Delegate = Get-MsolUser -All | Where-Object {$_.DisplayName -like "*$Search*"} | Sort-Object -Property Name
            }
            If ($Delegate.Length -gt 1) {Write-Host "Please make a selection for target mailbox delegate:" }
            If ($Delegate.Length -eq 1) {Write-Host "Located the following for target mailbox delegate:" }
		    $Delegate | Format-List DisplayName, UserPrincipalName
		    If ($Delegate.Length -gt 1) {
                Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $Delegate.Length
                    Write-Host
                    If (($Sel -le $Delegate.Length) -and ($Sel -gt 0)) {
                        $Delegate = $Delegate[$Sel-1]
                        $Set = $true
                        Write-Host "Selected" $Delegate.DisplayName
    	            }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	            } 
	            Until ($Set -eq $True)
		    }
            If (!$Delegate) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    }
	    Until ($Delegate)
        Set-Mailbox -Identity $User.MS.UserPrincipalName -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $Delegate.UserPrincipalName
        Write-Host "User-EXOMailbox-Forward: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
        }
        Until ($Done)
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-EXOMailbox-AddDelegate {
    Try {
            Do {
            $Delegate = $Null
            Do {
            If (!$Delegate) {
                $Search = Read-Host "Who needs delegate access?"
                If ($Search -eq "") {Write-Host ; Return}
                $Delegate = Get-MsolUser -All | Where-Object {$_.DisplayName -like "*$Search*"} | Sort-Object -Property Name
            }
            If ($Delegate.Length -gt 1) {Write-Host "Please make a selection for target mailbox delegate:" }
            If ($Delegate.Length -eq 1) {Write-Host "Located the following for target mailbox delegate:" }
		    $Delegate | Format-List DisplayName, UserPrincipalName
		    If ($Delegate.Length -gt 1) {
                Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $Delegate.Length
                    Write-Host
                    If (($Sel -le $Delegate.Length) -and ($Sel -gt 0)) {
                        $Delegate = $Delegate[$Sel-1]
                        $Set = $true
                        Write-Host "Selected" $Delegate.DisplayName
    	            }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	            } 
	            Until ($Set -eq $True)
		    }
            If (!$Delegate) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    }
	    Until ($Delegate)
        Add-MailboxPermission -Identity $User.MS.UserPrincipalName -User $Delegate.UserPrincipalName -AccessRights FullAccess -InheritanceType All
        #Add-MailboxPermission -Identity $User.MS.UserPrincipalName -User $Delegate.UserPrincipalName -AccessRights SendAs -InheritanceType All
        Write-Host "User-EXOMailbox-AddDelegate: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
        }
        Until ($Done)
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-ClearGroups {
    Try {
        ForEach ($Group in $User.MSG) {
            $Group = $Client.MSG | Where-Object {$_.ObjectID -eq $Group}
            $Member = (Get-MsolGroupMember -All -GroupObjectId $Group.ObjectId.ToString() | Where {$_.Emailaddress -eq $User.MS.UserPrincipalName } )
            Remove-MsoLGroupMember -GroupObjectId $Group.ObjectId.ToString() -GroupMemberType User -GroupmemberObjectId $Member.ObjectId -ErrorAction SilentlyContinue
            Remove-UnifiedGroupLinks -Identity $Group.DisplayName -LinkType Members -Links $User.MS.UserPrincipalName -Confirm:$False -ErrorAction SilentlyContinue
            $User.MSG = $User.MSG | Where-Object {$_ -ne $Group.ObjectID}
        }
        ### Write-Notes -Message "Reset User's Microsoft 365 groups."
        Write-Host "User-MS-ClearGroups: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
   } 
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-Disabled {
    Try {
        Set-MsolUser -UserPrincipalName  $User.MS.UserPrincipalName -BlockCredential $True
        ### Write-Notes -Message "Disabled User's Microsoft 365 sign-in."
        $User.MS.BlockCredential = Get-MsolUser -UserPrincipalName $User.PC.UserPrincipalName | Select-Object -ExpandProperty BlockCredential
        Write-Host "User-MS-Disabled: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-Enabled {
    Try {
        Set-MsolUser -UserPrincipalName  $User.MS.UserPrincipalName -BlockCredential $False
        ### Write-Notes -Message "Un-blocked User's Microsoft 365 sign-in."
        $User.MS.BlockCredential = Get-MsolUser -UserPrincipalName $User.PC.UserPrincipalName | Select-Object -ExpandProperty BlockCredential
        Write-Host "User-MS-Enabled: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-PWRandom {
    Try {
        Write-Host "User MS passsword reset to:"
        $Password = (Invoke-RestMethod http://www.dinopass.com/password/strong)
        Set-MsolUserPassword -UserPrincipalName $User.MS.UserPrincipalName -NewPassword $Password
        ### Write-Notes -Message "Reset User's Microsoft 365 password."
        $User.MS.LastPasswordChangeTimestamp = Get-MsolUser -UserPrincipalName $User.PC.UserPrincipalName | Select-Object -ExpandProperty LastPasswordChangeTimestamp
        Write-Host "User-MS-PWRandom: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-ResetMFA {
    Try {
        Reset-MsolStrongAuthenticationMethodByUpn -UserPrincipalName $User.MS.UserPrincipalName
        ### Write-Notes -Message "Reset User's Microsoft 365 Multi-Factor Authentication."
        Write-Host "User-MS-ResetMFA: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
   } 
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-DisableMFA {
    Try {
        Set-MsolUser -UserPrincipalName $User.MS.UserPrincipalName -StrongAuthenticationRequirements @()
        ### Write-Notes -Message "Reset User's Microsoft 365 Multi-Factor Authentication."
        $User.MS.StrongAuthenticationRequirements = Get-MsolUser -UserPrincipalName $User.PC.UserPrincipalName | Select-Object -ExpandProperty StrongAuthenticationRequirements
        Write-Host "User-MS-DisableMFA: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
   } 
    Catch {
        $Error[0]
    }
}

Function Global:User-MS-ClearLicenses {
    Try {

        ForEach ($License in $User.MS.Licenses) {
            Set-MsolUserLicense -UserPrincipalName $User.MS.UserPrincipalName -RemoveLicenses $License.AccountSkuID
            ### Write-Notes -Message "Removed User's $Licenses.AccountSkuID  Microsoft license."
        }
        $User.MS.Licenses = Get-MsolUser -UserPrincipalName $User.PC.UserPrincipalName | Select-Object -ExpandProperty Licenses
        Write-Host "User-MS-ClearLicenses: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
   } 
    Catch {
        $Error[0]
    }
}

Function Global:User-AD-Description {
    Try {
        If ($AutoTicket) {$Ticket = $AutoTicket}
        Else {$Ticket = Read-Host "Please enter Service Ticket #"}
        Set-ADUser -Identity $User.AD.SamAccountName -Replace @{description="Disabled Employee - Service Ticket $Ticket"} -ErrorAction SilentlyContinue
        ### Write-Notes -Message "Hide User's Active Directory profile in Global Address List."
        Write-Host "User-AD-Description: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-AD-Disabled {
	Try {
		Disable-ADAccount -Identity $User.AD.SamAccountName
        ### Write-Notes -Message "Disabled User's Active Directory sign-in."
		Write-Host "User-AD-Disabled: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
	}
	Catch {
	    $Error[0]
	}
}

Function Global:User-AD-Enabled {
	Try {
		Enable-ADAccount -Identity $User.AD.SamAccountName
        ### Write-Notes -Message "Enabled User's Active Directory sign-in."
		Write-Host "User-AD-Enabled: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
	}
	Catch {
	    $Error[0]
	}
}

Function Global:User-AD-PWRandom {
    Try {
        $Password = (Invoke-RestMethod http://www.dinopass.com/password/strong)
        Write-Host "User AD passsword reset to:"
        Write-Host $Password
        Set-ADAccountPassword -Identity $User.AD.SamAccountName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force)
        ### Write-Notes -Message "Reset User's Active Directory password."
        Write-Host "User-AD-PWRandom: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-AD-ClearGroups {
    Try {
        ForEach ($Group in $User.ADG) {
            If ($Group -ne "Domain Users") {Remove-ADPrincipalGroupMembership -Identity $User.AD.SamAccountName -MemberOf "$Group" -Confirm:$false -ErrorAction Continue }
        $User.ADG = $User.ADG | Where-Object {$_ -ne $Group}    
        }
        ### Write-Notes -Message "Reset User's Active Directory groups."
        Write-Host "User-AD-ClearGroups: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
   } 
    Catch {
        $Error[0]
    }
}

Function Global:User-AD-HideFromGAL {
    Try {
        Set-ADUser -Identity $User.AD.SamAccountName -Replace @{msExchHideFromAddressLists="TRUE"} -ErrorAction SilentlyContinue
        ### Write-Notes -Message "Hide User's Active Directory profile in Global Address List."
        Write-Host "User-AD-HideFromGAL: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Function Global:User-AD-MvToRetntnOU {
    Try {
        $CurrentOU = $User.AD.DistinguishedName
        $DisabledOU = Get-ADOrganizationalUnit -Filter 'Name -like "*disabled*"'
        $Set = $False
        Do {
            If ($DisabledOU.Length -gt 1) {Write-Host "Please make a selection for target OU:" }
            If ($DisabledOU.Length -eq 1) {Write-Host "Located the following for target OU:" }
		    $DisabledOU | Format-List DistinguishedName
		    If ($DisabledOU.Length -gt 1) {
  
                Do {
                    [int]$Sel = Read-Host "Please select a record, 1 -" $DisabledOU.Length
                    Write-Host
                    If (($Sel -le $DisabledOU.Length) -and ($Sel -gt 0)) {
                        $DisabledOU = $DisabledOU[$Sel-1]
                        $Set = $true
                        Write-Host "Selected" $DisabledOU.Name
    	            }
                    Else {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	            } 
	            Until ($Set -eq $True)
		    }
            If (!$DisabledOU) {Write-Host "Invalid selection!" -ForegroundColor Red -BackgroundColor White}
	    }
        Until ($DisabledOU)

        Write-Host "Please VERIFY you want to select:" $DisabledOU.DistinguishedName -ForegroundColor Red -BackgroundColor White
        Write-Host "SELECTING THE WRONG OU CAN RESULT IN MAILBOX DELETION" -ForegroundColor Red -BackgroundColor White
        #Act on input, clear if not verified.
        If ((Read-Host "Type YES") -ne "YES" ) {
            $CurrentOU = $Null
            $DisabledOU = $Null
            Write-Host "User-AD-MvToRetntnOU: Cancelled" -BackgroundColor White -ForegroundColor Red ; Write-Host
            Throw "The operator did not verify the User object as needed, no selection was made."
        }
        #At this point we must be verified (which should actually always be true at this point).
        Else { Write-Host "VERIFIED!" -BackgroundColor White -ForegroundColor Blue 
        Move-ADObject -Identity $CurrentOU -TargetPath $DisabledOU
        ### Write-Notes -Message "Moved User's Active Directory profile into $DisabledOU."
        Write-Host "User-AD-MvToRetntnOU: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host}
    }
    Catch {
        $Error[0]
    }
}

Function Global:Client-AD-DirSync {
    Try {
        $DirsyncService = get-service -Name ADSync | Select-Object -ExpandProperty Status -ErrorAction SilentlyContinue
        If ($DirsyncService -eq "Running") {
            Start-ADSyncSyncCycle -Policytype Delta
            ### Write-Notes -message "Ran manual AzureAD directory sync"
        }
        Write-Host "Client-AD-DirSync: Successful" -BackgroundColor White -ForegroundColor Blue ; Write-Host
    }
    Catch {
        $Error[0]
    }
}

Print-CMDs

### Danger, there be dragons (and unfinished, but contained code) ahead! ###
