#Script Use: Microsoft Teams Voice Migration, Configuration, & Management Utility
#Created by Eric Marsi (www.UCIT.Blog)
#Version 2108.1_BETA
#Date: March 10, 2021
#Updated: August 19, 2021

Clear-Host

Write-Host "**************************************************************************************" -ForegroundColor Green
Write-Host "*        Microsoft Teams Voice Migration, Configuration, & Management Utility        *" -ForegroundColor Green
Write-Host "*                         Created by Eric Marsi (www.UCIT.Blog)                      *" -ForegroundColor Green
Write-Host "*                                 Version 2108.1_BETA                                *" -ForegroundColor Green
Write-Host "*                                 Date: March 10, 2021                               *" -ForegroundColor Green
Write-Host "*                               Updated: August 19, 2021                             *" -ForegroundColor Green
Write-Host "**************************************************************************************`n" -ForegroundColor Green

#Script Tests 
#Verify that the Script is executing as an Administrator
Write-Host "Verifying that the script is executing as an Administrator"
function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    }
    if (!(Test-IsAdmin)){
        throw "Please run this script as an Administrator!"
    }
    else {
        Write-Host "Pass: The script is executing as an Administrator `n" -ForegroundColor Green       
    }

#Verify that at least PowerShell 5.1 is Installed
Write-Host "Verifying that at least PowerShell 5.1 is Installed"
    if([Version]'5.1.00000.000' -GT $PSVersionTable.PSVersion)
    {
        Write-Error "The host must be upgraded to at least PowerShell 5.1! Please Refer to: https://www.ucit.blog/post/setting-up-the-microsoft-teams-powershell-module" -ErrorAction Stop
    }else {
        Write-Host "Pass: The host has at least PowerShell 5.1 installed`n" -ForegroundColor Green
    }

#Verify that the script is executing in the PowerShell Console and not the ISE
Write-Host "Verifying that the script is executing in the PowerShell Console and not the ISE"
    if((Get-Host).Name -eq "ConsoleHost")
    {
        Write-Host "Pass: The script is executing in the PowerShell Console`n" -ForegroundColor Green
    }else {
        Write-Error "The script is not executing in the PowerShell Console!" -ErrorAction Stop
    }

#Verify that the Teams PowerShell Module is installed. If not installed, attempt to install
Write-Host "Verifying that the Teams PowerShell Module is Installed"
    try {
            if(Get-Module -ListAvailable MicrosoftTeams)
            {
                Write-Host "Pass: Microsoft Teams Module is Installed`n" -ForegroundColor Green
            }else {
                Write-Host "Attempting to Install the Microsoft Teams Module"
                Install-Module MicrosoftTeams
                Write-Host ""
            }   
        }
    catch {
        Write-Host "An Unexpected Error occured! The exception caught was $_ " -ForegroundColor Red
        Write-Error "Script Terminating due to error during the Teams PS Module Test! " -ErrorAction Stop
        }

#Connect to Teams Online PowerShell
Write-Host "Connecting to Teams Online PowerShell"
    try{
        Import-Module MicrosoftTeams
        $null = Connect-MicrosoftTeams
        Write-Host "Pass: Connected to Microsoft Teams PowerShell`n" -ForegroundColor Green
    }catch{
        Write-Output  "An Unexpected Error occured! The exception caught was $_ "
        Write-Error "Script Terminating due to error when connecting to Microsoft Teams Online! " -ErrorAction Stop
    }

do{
#Mode Selection
    Write-Host "Mode Selection---------------------------------------------------------------------" -ForegroundColor Yellow
    Write-Host "Skype for Business Server Modes----------------------------------------------------"
    Write-Host "Option 1 - Export a Dial Plan from Skype for Business Server" -ForegroundColor Green
    Write-Host "Option 2 - Export Trunk Translation Rules from Skype for Business Server`n" -ForegroundColor Green
    Write-Host "Microsoft Teams Modes--------------------------------------------------------------"
    Write-Host "Option 3 - Create a New Teams Dial Plan (1000 Dial Plan Tenant Limit)" -ForegroundColor Green
    Write-Host "Option 4 - Remove an Existing Teams Dial Plan" -ForegroundColor Green
    Write-Host "Option 5 - Import Normalization Rules into an existing Teams Dial Plan (No Limit)" -ForegroundColor Green
    Write-Host "Option 6 - Import Trunk Translation Rules into the Teams Tenant (400 Rule Tenant Limit)" -ForegroundColor Green
    Write-Host "Option 7 - Assign Trunk Translation Rules to a Specified Gateway (100 Rule Limit per Direction)" -ForegroundColor Green
    Write-Host "Option 8 - Import Voice Routes into the Teams Tenant (No Limit)`n" -ForegroundColor Green
    Write-Host "Option 99 - Disconnect from Teams Online PowerShell & Terminate this Script`n" -ForegroundColor Green
    Write-Host "-----------------------------------------------------------------------------------`n"
    $Mode = Read-Host "Of the above options, what mode would you like to run this script in? (Enter the Option Number)"
    Write-Host ""
    Write-Host "-----------------------------------------------------------------------------------"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 1 - Export a Dial Plan from Skype for Business Server Mode
if ($Mode -eq "1")
    {
        Write-Host "Export a Dial Plan from Skype for Business Server Mode Selected`n" -ForegroundColor Green

        #Verify that the Skype for Business Server PowerShell Module is installed and that this is a SFB Server.
        Write-Host "Verifying that the Skype for Business PowerShell Module is Installed"
        try {
                $null = Get-Module -ListAvailable SkypeforBusiness -ErrorAction Stop
                $null = Get-CsServerPatchVersion -ErrorAction Stop
                Write-Host "Pass: The Skype for Business Server Module is Installed and this is a SFB Server`n" -ForegroundColor Green
            }
        catch 
            {
                Write-Host "An Unexpected Error Occured! The exception caught was $_ " -ForegroundColor Red
                Write-Error "Script Terminating due to error during the SFB Server PS Module Test! " -ErrorAction Stop
            }

        #Prompt for what Dial Plan to export normalization rules from
        Write-Host "Obtaining all Dial Plans from Skype for Business Server. Please Standby...`n"
        try{
            (SkypeforBusiness\Get-CsDialPlan).Identity | ForEach-Object { $_ -replace 'Tag:',''}
        }catch{
            Write-Host "An Unexpected Error occured while obtaining all dial plans from Skype for Business Server! The exception caught was: $_ " -ForegroundColor Red
            Write-Error "Script Terminating due to error while obtaining all dial plans from Skype for Business Server! " -ErrorAction Stop
        }
        Write-Host 
        $SDP = Read-Host "Which Dial Plan do you want to export normalization rules from? (Type the full name above)"
        Write-Host ""

        #Prompt for where dial plan rules should be saved to
        try
            {
                Write-Host "Please select the folder where you would like the normalization rules from the $($SDP) dial plan to be saved"
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                [void]$FileBrowser.ShowDialog()
                $SPath = $FileBrowser.SelectedPath 
                Write-Host "$($SPath) selected as the save directory`n" -ForegroundColor Green
            }
        catch
            {
                Write-Error "No File Selected!" -ErrorAction Stop
            }

        #Save the dial plan to file
        Write-Host "Exporting the $($SDP) dial plan from Skype for Business Server to file - This action can take some time so please be patient..."
        $DTS = Get-Date -Format "MM-dd-yyyy-HHmmssfff"
        try 
            {
                SkypeforBusiness\Get-CsVoiceNormalizationRule -ErrorAction Stop | Select-Object Identity,Priority,Description,Pattern,Translation,Name,IsInternalExtension | Where-Object Identity -like "*$($SDP)*" | Export-Csv -Path "$($SPath)\SFB_DP_$($SDP)_$($DTS).csv"
                Write-Host "Successfully saved the $($SDP) dial plan from Skype for Business Server to file at $($SPath)\SFB_DP_$($SDP)_$($DTS).csv`n" -ForegroundColor Green
            }
        catch 
            {
                Write-Host "An Unexpected Error occured while saving the $($SDP) dial plan from Skype for Business Server! The exception caught was: $_ " -ForegroundColor Red
                Write-Error "Script Terminating due to error while saving the $($SDP) dial plan from Skype for Business Server! " -ErrorAction Stop
            }
    }   

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 2 - Export Trunk Translation Rules from Skype for Business Server Mode
elseif ($Mode -eq "2")
    {
        Write-Host "Export Trunk Translation Rules from Skype for Business Server Mode Selected`n" -ForegroundColor Green

        #Verify that the Skype for Business Server PowerShell Module is installed and that this is a SFB Server.
        Write-Host "Verifying that the Skype for Business PowerShell Module is Installed"
        try {
                $null = Get-Module -ListAvailable SkypeforBusiness -ErrorAction Stop
                $null = Get-CsServerPatchVersion -ErrorAction Stop
                Write-Host "Pass: The Skype for Business Server Module is Installed and this is a SFB Server`n" -ForegroundColor Green
            }
        catch 
            {
                Write-Host "An Unexpected Error Occured! The exception caught was $_ " -ForegroundColor Red
                Write-Error "Script Terminating due to error during the SFB Server PS Module Test! " -ErrorAction Stop
            }

        #Prompt for what trunk rules should be exported from
        Write-Host "Obtaining all Trunks from Skype for Business Server. Please Standby...`n"
        try{
            (SkypeforBusiness\Get-CsTrunkConfiguration).Identity
        }catch{
            Write-Host "An Unexpected Error occured while obtaining all trunks from Skype for Business Server! The exception caught was: $_ " -ForegroundColor Red
            Write-Error "Script Terminating due to error while obtaining all trunks from Skype for Business Server! " -ErrorAction Stop
        }
        Write-Host 
        $STrunk = Read-Host "Which trunk do you want to export translation rules from? (Type the full name above)"
        Write-Host ""

        #Prompt for where translation rules should be saved to
        try
            {
                Write-Host "Please select the folder where you would like the translation rules from the $($STrunk) trunk to be saved"
                Add-Type -AssemblyName System.Windows.Forms
                $FileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
                [void]$FileBrowser.ShowDialog()
                $SPath = $FileBrowser.SelectedPath 
                Write-Host "$($SPath) selected as the save directory`n" -ForegroundColor Green
            }
        catch
            {
                Write-Error "No File Selected!" -ErrorAction Stop
            }
        
        $DTS = Get-Date -Format "MM-dd-yyyy-HHmmssfff"
        $STrunkFN = $STrunk | ForEach-Object { $_ -replace ':','_'}

        #Save the Trunk Outbound Calling Number Translation Rules to file
        Write-Host "Exporting the Trunk Outbound Calling Number Translation Rules for the $($STrunk) Trunk from Skype for Business Server to file - This action can take some time so please be patient..."
        $TOCNTRTest = (SkypeforBusiness\Get-CsTrunkConfiguration -Identity $STrunk).OutboundCallingNumberTranslationRulesList
        if ($TOCNTRTest -ne "") 
            {
                try 
                    {
                        (SkypeforBusiness\Get-CsTrunkConfiguration -Identity $STrunk).OutboundCallingNumberTranslationRulesList | Select-Object Name,Description,Pattern,Translation | Export-Csv -Path "$($SPath)\SFB_Trunk_$($STrunkFN)_OBCallingNumTran_$($DTS).csv"
                        Write-Host "Successfully saved the Trunk Outbound Calling Number Translation Rules from Skype for Business Server to file at $($SPath)\SFB_Trunk_$($STrunkFN)_OBCallingNumTran_$($DTS).csv`n" -ForegroundColor Green
                    }
                catch 
                    {
                        Write-Host "An Unexpected Error occured while saving the Trunk Outbound Calling Number Translation Rules from Skype for Business Server! The exception caught was: $_ " -ForegroundColor Red
                        Write-Error "Script Terminating due to error while saving the Trunk Outbound Calling Number Translation Rules from Skype for Business Server! " -ErrorAction Stop
                    }
                }
        else
            {
                Write-Host "The $($STrunk) Trunk does not contain any Outbound Calling Number Translation Rules`n" -ForegroundColor Yellow
            }
        $TOCNTRTest = $null
        
        #Save the Trunk Outbound Called Number Translation Rules to file
        Write-Host "Exporting the Trunk Outbound Called Number Translation Rules for the $($STrunk) Trunk from Skype for Business Server to file - This action can take some time so please be patient..."
        $TOCNTRTest = (SkypeforBusiness\Get-CsTrunkConfiguration -Identity $STrunk).OutboundTranslationRulesList
        if ($TOCNTRTest -ne "") 
            {
                try 
                    {
                        (SkypeforBusiness\Get-CsTrunkConfiguration -Identity $STrunk).OutboundTranslationRulesList | Select-Object Name,Description,Pattern,Translation | Export-Csv -Path "$($SPath)\SFB_Trunk_$($STrunkFN)_OBCalledNumTran_$($DTS).csv"
                        Write-Host "Successfully saved the Trunk Outbound Called Number Translation Rules Trunk from Skype for Business Server to file at $($SPath)\SFB_Trunk_$($STrunkFN)_OBCalledNumTran_$($DTS).csv`n" -ForegroundColor Green
                    }
                catch 
                    {
                        Write-Host "An Unexpected Error occured while saving the Trunk Outbound Called Number Translation Rules from Skype for Business Server! The exception caught was: $_ " -ForegroundColor Red
                        Write-Error "Script Terminating due to error while saving the Trunk Outbound Called Number Translation Rules from Skype for Business Server! " -ErrorAction Stop
                    } 
                }
        else
            {
                Write-Host "The $($STrunk) Trunk does not contain any Outbound Called Number Translation Rules`n" -ForegroundColor Yellow
            }
        $TOCNTRTest = $null
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 3 - Create a New Teams Dial Plan Mode
elseif ($Mode -eq "3")
    {
        Write-Host "Create a New Teams Dial Plan Mode Selected`n" -ForegroundColor Green
        Write-Host "Verifying if more Dial Plans can be created in this tenant"
        $TCountDP= (Get-CsTenantDialPlan).Count
        $TRemainDP = (999 - (Get-CsTenantDialPlan).Count)
            if ($TCountDP -gt 9999)
                {
                    Write-Host "The Tenant has hit the maximum limit of 1000 Dial Plans that can be created." -ForegroundColor Red
                    Write-Error "Script Terminating due to error when creating a new Dial Plan! " -ErrorAction Stop
                }else {
                    Write-Host "Pass: $($TRemainDP) more Dial Plan(s) can be created in this tenant`n" -ForegroundColor Green
                    $DP = Read-Host "What would you like the new Dial Plan to be called?`n"
                    Write-Host ""
                    Write-Host "Creating the Dial Plan named $($DP), Please Standby`n"
                }

        try{
            $null = New-CsTenantDialPlan -Identity $DP -ErrorAction Stop
            Write-Host "The Dial Plan named $($DP) was created successfully" -ForegroundColor Green
        }catch{
            Write-Host "An Unexpected Error occured while creating the Dial Plan named $($DP)! The exception caught was: $_ " -ForegroundColor Red
        }
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 4 - Remove an Existing Teams Dial Plan Mode
elseif ($Mode -eq "4")
    {
        Write-Host "Remove an Existing Teams Dial Plan Mode Selected`n" -ForegroundColor Green

        #Prompt for what Dial Plan to import normalization rules into
        Write-Host "Obtaining all Dial Plans in the Tenant. Please Standby...`n"
        try{
            (Get-CsTenantDialPlan).Identity | ForEach-Object { $_ -replace 'Tag:',''}
        }catch{
            Write-Host "An Unexpected Error occured while obtaining all Dial Plans from the tenant! The exception caught was: $_ " -ForegroundColor Red
            Write-Error "Script Terminating due to error while obtaining all Dial Plans from the tenant! " -ErrorAction Stop
        }
        Write-Host 
        $DP = Read-Host "Which Dial Plan would you like to Delete? (Type the full name above)`n"
        Write-Host ""
        Write-Host "Deleting the Dial Plan named $($DP), Please Standby`n"

        try{
            $null = Remove-CsTenantDialPlan -Identity $DP -ErrorAction Stop
            Write-Host "The Dial Plan named $($DP) was deleted successfully" -ForegroundColor Green
        }catch{
            Write-Host "An Unexpected Error occured while deleting the Dial Plan named $($DP)! The exception caught was: $_ " -ForegroundColor Red
        }
    }


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 5 - Import Normalization Rules into an existing Teams Dial Plan Mode
elseif ($Mode -eq "5")
    {
        Write-Host "Import Normalization Rules into an existing Teams Dial Plan Mode Selected`n" -ForegroundColor Green

        #Prompt for CSV Import of a Given Dial Plan
        try{
            Write-Host "Please Select the CSV containing the normalization rules you wish to import"
            Write-Host ""
            Add-Type -AssemblyName System.Windows.Forms
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
            $FileBrowser.filter = "Csv (*.csv)| *.csv"
            [void]$FileBrowser.ShowDialog()
            $NormRules = Import-Csv -Path $FileBrowser.FileName
        }
        catch{
            Write-Error "No File Selected!" -ErrorAction Stop
        }

        #Prompt for what Dial Plan to import normalization rules into
        Write-Host "Obtaining all Dial Plans in the Tenant. Please Standby...`n"
        try{
            (Get-CsTenantDialPlan).Identity | ForEach-Object { $_ -replace 'Tag:',''}
        }catch{
            Write-Host "An Unexpected Error occured while obtaining all Dial Plans from the tenant! The exception caught was: $_ " -ForegroundColor Red
            Write-Error "Script Terminating due to error while obtaining all Dial Plans from the tenant! " -ErrorAction Stop
        }
        Write-Host 
        $DP = Read-Host "Which Dial Plan do you want to import these normalization Rules into? (Type the full name above)`n"

        #Import Dial Plan Rules
        $S1 = 1
        $Count = $NormRules.Count
        $Remain = $NormRules.Count
        Write-Host ""
        Write-Host "Importing $($Count) Normalization Rules into the Dial Plan named $($DP) - This action can take some time so please be patient"
        
        $Import = @()        
        foreach ($NormRule in $NormRules)
            {
                try
                    {
                        [bool]$IntExt = "$" + $NormRule.IsInternalExtension
                        $NR = New-CsVoiceNormalizationRule -Parent Global -Name $NormRule.Name -Priority $NormRule.Priority -Pattern $NormRule.Pattern -Translation $NormRule.Translation -IsInternalExtension $IntExt -InMemory -Verbose -ErrorAction Stop
                        $Import += $NR
                        $Remain = ($($Remain) - 1)
                        Write-Host "Rule $($S1) | $($NormRule.Name) was created and is awaiting import to the tenant | $($Remain) Rules Remaining..." -ForegroundColor Green
                        $S1 ++
                    }
                catch
                    {
                        $Remain = ($($Remain) - 1)
                        Write-Host "Rule $($S1) | $($NormRule.Name) was NOT created. The exception caught was: $_ | $($Remain) Rules Remaining..." -ForegroundColor Red
                    }
            }

        try{
            Write-Host ""
            Write-Host "Importing $($Count) Normalization Rules into the $($DP) Dial Plan. Please Standby..."
            Set-CsTenantDialPlan -Identity $DP -NormalizationRules @{add=$Import} -Force -Verbose -ErrorAction Stop
            Write-Host "The import of $($Count) Normalization Rules into the $($DP) Dial Plan Completed Successfully!" -ForegroundColor Green
        }catch{
            Write-Host ""
            Write-Host "The import of $($Count) Normalization Rules into the $($DP) Dial Plan failed. The exception caught was: $_" -ForegroundColor Red
        }
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 6 - Import Trunk Translation Rules into the Teams Tenant Mode
elseif ($Mode -eq "6")
    {
    Write-Host "Import Trunk Translation Rules into the Teams Tenant Mode Selected`n" -ForegroundColor Green

    #Prompt for CSV Import of a Given Dial Plan
    try{
        Write-Host "Please Select the CSV containing the translation rules you wish to import`n"
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.filter = "Csv (*.csv)| *.csv"
        [void]$FileBrowser.ShowDialog()
        $NormRules = Import-Csv -Path $FileBrowser.FileName
    }
    catch{
        Write-Error "No File Selected!" -ErrorAction Stop
    }

    #Trunk Translation Rules Limit Test
    $TRTenant = (Get-CsTeamsTranslationRule).Count
    $Count = $NormRules.Count
    $TRAdd = $Count + $TRTenant
    $TRSub = 400 - $Count

    
    #Verify that the tenant does not already have 400 or more existing translation rules
    if ($TRTenant -ge 400)
    {
        Write-Host "The Tenant has hit the maximum limit of 400 Trunk Translation Rules that can be created." -ForegroundColor Red
        Write-Error "Script Terminating due to a limit error while importing Trunk Translation Rules! " -ErrorAction Stop
    
    #Verify that the import will not fail when the rules in the import CSV are added to the existing translation rules in the tenant
    }elseif ($TRAdd -gt 400){
        Write-Host "There are $($TRAdd) Rules between existing and new rules which is above the tenant limit of 400. If you continue, $($TRSub) rule(s) can be imported" -ForegroundColor Yellow
        $GTRTerm = Read-Host "Would you like to continue to import $($TRSub) Rules which will fail after Rule #$($TRSub)? (Y/N)`n"
            if ($GTRTerm -eq 'Y')
                {
                    Write-Host "Continuing to Import $($TRSub) Rules into the Tenant`n"
                }else {
                    Write-Error "Script Terminating due to a limit error while importing Trunk Translation Rules! " -ErrorAction Stop
                }

    }else {
        Write-Host "$($Count) Translation Rules will be imported into the tenant`n"
    }

    #Import Tranlation Rules into the Tenant
    Write-Host "Importing Trunk Translation Rules into the tenant - This action can take some time so please be patient...`n"
    $S1 = 1
    $Remain = $NormRules.Count
    foreach ($NormRule in $NormRules)
       {6
            try{
                New-CsTeamsTranslationRule -Identity $NormRule.Name -Description $NormRule.Description -Pattern $NormRule.Pattern -Translation $NormRule.Translation -ErrorAction Stop | Out-Null
                $Remain = ($($Remain) - 1)
                Write-Host "Rule $($S1) | $($NormRule.Name) was created | $($Remain) Rules Remaining..." -ForegroundColor Green
            }catch {
                $Remain = ($($Remain) - 1)
                Write-Host "Rule $($S1) | Creating $($NormRule.Name) failed because: $_ | $($Remain) Rules Remaining..." -ForegroundColor Red
            }
            $S1 ++
        }
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 7 - Assign Trunk Translation Rules to a Specified Gateway Mode
elseif ($Mode -eq "7")
    {
    Write-Host "Assign Trunk Translation Rules to a Specified Gateway Mode Selected`n" -ForegroundColor Green
    
    #Prompt for CSV Import of a Given Dial Plan
    try{
        Write-Host "Please Select the CSV containing the translation rules you wish to import`n"
        Add-Type -AssemblyName System.Windows.Forms
        $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $FileBrowser.filter = "Csv (*.csv)| *.csv"
        [void]$FileBrowser.ShowDialog()
        $NormRules = Import-Csv -Path $FileBrowser.FileName
    }
    catch{
        Write-Error "No File Selected!" -ErrorAction Stop
    }

    #Prompt for what DR Gateway to import translation rules into
    Write-Host "Obtaining all Direct Routing SBCs, Please Standby`n"
    (Get-CsOnlinePSTNGateway).Identity
    Write-Host ""
    $GW = Read-Host "Which Direct Routing SBC do you want to import these translation rules into? (Type the full name above)"
    Write-Host ""

    Write-Host "Import Destination-----------------------------------------------------------------`n"
    Write-Host "Option 1 - InboundTeamsNumberTranslationRules" -ForegroundColor Green
    Write-Host "Option 2 - InboundPSTNNumberTranslationRules" -ForegroundColor Green
    Write-Host "Option 3 - OutboundTeamsNumberTranslationRules" -ForegroundColor Green
    Write-Host "Option 4 - OutboundPSTNNumberTranslationRules`n" -ForegroundColor Green
    Write-Host "-----------------------------------------------------------------------------------`n"
    $Dest = Read-Host "Of the above options, where should the rules be imported? (Enter the Option Number)"
    Write-Host ""
    Write-Host "-----------------------------------------------------------------------------------`n"

    $Count = $NormRules.Count
    $Rules = $NormRules.Name
    #Create PS Array List with Rule Names
    [array]$Rules2 = $null
        $null = foreach ($Rule in $Rules)
            {
                if ($Null -ne $Rules2)
                    {
                        $Rules2.Add("'" + $Rule + "'")
                    }
                else 
                    {
                        $Rules2 = "'" + $Rule + "'"
                        [System.Collections.ArrayList]$Rules2 = $Rules2
                    }
            }

    #Test to see if each rule has already created, if not, remove from the array list
    foreach ($Rule in $Rules2)
        {
            Try
                {
                    Get-CsTeamsTranslationRule -Identity $($Rule) -ErrorAction Stop
                    Write-Host "The $($Rule) rule already exists and does not need to be created. Removing from the Array List" -ForegroundColor Green
                    $Rules2.Remove("'" + $Rule + "'")
                }
            catch
                {
                    Write-Host "$($Rule) Needs to be Created" -ForegroundColor Green
                }
                
        }
        
    #Test for more than 100 Rules being imported       
    if ($Count -gt 100)
        {
            $CountO = $Count - 100
            $CRes = Read-Host "There are $($CountO) to many Translation Rules to Import. Do you want to assign the first 100 of the $($Count) Translation Rules to the $($GW) SBC? (Y/N)`n"
            if ($Cres -eq 'Y')
                    {
                        while ($Rules2.Count -gt 100)
                            {
                                $Rules2.Remove($Rules2[100])
                            }
                            $Count ='100'
                    } else {
                        Write-Host "Terminating Function as the user canceled because of Import Count being incorrect!`n" -ForegroundColor Yellow
                        $OP2 = 'N'
                    }
        } else {
            Write-Host "There are $($Count) Translation Rules which is less than the maximum limit of 100, Continuing`n"
        }


    do{
        Write-Host "Assigning Translation Rules to the $($GW) SBC - This action can take some time so please be patient...`n"
        $Import = $Rules2 -join ","
      
        if($Dest -eq '1')
            {
                try{
                    Invoke-Expression "Set-CsOnlinePSTNGateway -Identity `$GW -InboundTeamsNumberTranslationRules @{add=$Import}" -ErrorAction Stop
                    Write-Host "$($Count) Translation Rules imported to the $($GW) SBC as InboundTeamsNumberTranslationRules" -ForegroundColor Green
                }catch{
                    Write-Host "$($Count) Translation Rules cannot be imported to the $($GW) SBC as InboundTeamsNumberTranslationRules because: $_" -ForegroundColor Red
                }
            }elseif ($Dest -eq '2')
            {
                try{
                    Invoke-Expression "Set-CsOnlinePSTNGateway -Identity `$GW -InboundPSTNNumberTranslationRules @{add=$Import}" -ErrorAction Stop
                    Write-Host "$($Count) TranslationRules imported to the $($GW) SBC as InboundPSTNNumberTranslationRules" -ForegroundColor Green
                }catch{
                    Write-Host "$($Count) Translation Rules cannot be imported to the $($GW) SBC as InboundPSTNNumberTranslationRules because: $_" -ForegroundColor Red
                }
            }elseif ($Dest -eq '3')
            {
                try{
                    Invoke-Expression "Set-CsOnlinePSTNGateway -Identity `$GW -OutboundTeamsNumberTranslationRules @{add=$Import}" -ErrorAction Stop
                    Write-Host "$($Count) Translation Rules imported to the $($GW) SBC as OutboundTeamsNumberTranslationRules" -ForegroundColor Green
                }catch{
                    Write-Host "$($Count) Translation Rules cannot be imported to the $($GW) SBC as OutboundTeamsNumberTranslationRules because: $_" -ForegroundColor Red
                }
            }elseif ($Dest -eq '4')
            {
                try{
                    Invoke-Expression "Set-CsOnlinePSTNGateway -Identity `$GW -OutboundPSTNNumberTranslationRules @{add=$Import}" -ErrorAction Stop
                    Write-Host "$($Count) Translation Rules imported to the $($GW) SBC as OutboundPSTNNumberTranslationRules" -ForegroundColor Green
                }catch{
                    Write-Host "$($Count) Translation Rules cannot be imported to the $($GW) SBC as OutboundPSTNNumberTranslationRules because: $_" -ForegroundColor Red
                }
            }else {
                Disconnect-MicrosoftTeams
                Write-Error "Invalid or no script option mode selected!" -ErrorAction Stop
            }
        }
        while ($OP2 -eq "Y")     
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 8 - Import Voice Routes into the Teams Tenant Mode
elseif ($Mode -eq "8")
    {
        Write-Host "Import Voice Routes into the Teams Tenant Mode Selected`n" -ForegroundColor Green

        #Prompt for CSV Import of voice routes
        try{
            Write-Host "Please Select the CSV containing the voice routes you wish to import"
            Write-Host ""
            Add-Type -AssemblyName System.Windows.Forms
            $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
            $FileBrowser.filter = "Csv (*.csv)| *.csv"
            [void]$FileBrowser.ShowDialog()
            $VoiceRoutes = Import-Csv -Path $FileBrowser.FileName
        }
        catch{
            Write-Error "No File Selected!" -ErrorAction Stop
        }
        
        #Prompt for which SBC the voice routes should be assigned
        Write-Host "Obtaining all SBCs in the Tenant. Please Standby...`n"
        try{
            (Get-CsOnlinePSTNGateway).Identity
        }catch{
            Write-Host "An Unexpected Error occured while obtaining all SBCs from the tenant! The exception caught was: $_ " -ForegroundColor Red
            Write-Error "Script Terminating due to error while obtaining all SBCs from the tenant! " -ErrorAction Stop
        }
        Write-Host 
        $SBC = Read-Host "Which SBC do you want to assign these voice routes to? (Type the full name above)`n"

        #Prompt for what PSTN Usage to import normalization rules into
        Write-Host ""
        Write-Host "Obtaining all PSTN Usages in the Tenant. Please Standby...`n"
        try{
            (Get-CsOnlinePstnUsage).Usage
        }catch{
            Write-Host "An Unexpected Error occured while obtaining all PSTN Usages from the tenant! The exception caught was: $_ " -ForegroundColor Red
            Write-Error "Script Terminating due to error while obtaining all PSTN Usages from the tenant! " -ErrorAction Stop
        }
        Write-Host 
        $PSTNUsage = Read-Host "Which PSTN Usage do you want assign these voice routes to? (Type the full name above)`n"

        #Import Voice Routes
        $Count = $VoiceRoutes.Count
        Write-Host ""
        Write-Host "Importing $($Count) Voice Routes into the tenant - This action can take some time so please be patient"
        $Remain = $VoiceRoutes.Count
        $Priority = 0
        $RuleCount = 1
        foreach ($VoiceRoute in $VoiceRoutes)
            {
                try{
                    $null = New-CsOnlineVoiceRoute -Identity $($VoiceRoute.Name) -OnlinePstnUsages @{add="$($PSTNUsage)"} -OnlinePstnGatewayList @{add="$($SBC)"} -Priority $($Priority) -NumberPattern $($VoiceRoute.Pattern) -ErrorAction Stop
                    $Remain = ($($Remain) - 1)
                    $Priority =($($Priority) + 1)
                    Write-Host "Rule $($RuleCount) | $($VoiceRoute.Name) was created successfully | $($Remain) Rules Remaining..." -ForegroundColor Green
                    $RuleCount =($($RuleCount) + 1)

                }catch{
                    $Remain = ($($Remain) - 1)
                    Write-Host "Rule $($RuleCount) | $($VoiceRoute.Name) was NOT created. The exception caught was: $_ | $($Remain) Rules Remaining..." -ForegroundColor Red
                }  
            }                
    }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Option 99 - Script Termination
elseif ($Mode -eq "99")
    {
        $Terminate = $True
    }

else
    {
        Write-Host "No Mode Selected or invalid response!" -ForegroundColor Red
    }    
Write-Host ""
if ($Terminate -eq $True)
    {
        $SResponse = "N"
    }
    else
    {
        $SResponse = Read-Host "Would you like to return to the mode selection window? (Y/N)"
        Write-Host ""
        Write-Host "-----------------------------------------------------------------------------------`n"
    }
}
while ($SResponse -eq "Y") 
Write-Host "Disconnecting from Teams Online PowerShell"
#Disconnect-MicrosoftTeams
Write-Host "Disconnected from Teams Online PowerShell" -ForegroundColor Green
