    function Main ($param) {

        if (!$param) { 
            Write-Output "Invalid Parameters";
            EXIT 1; 
        }
        
        $verified;
        $param = [system.string]::join(" ", $param);
        
        switch ($param)
        {
			## Windows Activation
            "checkAct" {
			   $lic = (cscript /NoLogo 'c:\\Windows\\System32\\slmgr.vbs' /dli)
			   $info = @();		## Collect Output From slmgr.vbs
				
			   foreach ($line in $lic) {
				   $filter = @("Name:", "Description");
				   $filter | ForEach-Object { if ($line -match $_) { continue; } }
				   if ([System.String]::IsNullOrWhiteSpace($line)) { continue; }				   
				   $info += $line;
			   }
			   
			   $verified = ($info -match "AB1XY").Length -GT 0;		## Find bulk license product key

			   $info[1];
			   $info[0];
            }				

			## Driver Errors
            "checkPnP" {
				$drivers = (Get-PnPDevice -Status 'Error' -ErrorAction SilentlyContinue).Caption;
				if ($drivers.Length -GT 0) { 
					$verified = $false;
					Write-Output "Driver Issue(s) Found";
					$drivers | Foreach-Object { Write-Output " - $_ "; };
				}
				
				else {
					$verified = $true;
					Write-Output "No device driver issues found";
				}
            }	

			## Domain Joined
            "checkDC" {
				$verified = (Get-CimInstance	-ClassName Win32_ComputerSystem).Domain -eq "contoso.local";
				if ($verified) { Write-Output "Domain is joined"; }
            }
			
			"checkUser" {
				## Get All AD Users In Local Groups
				## Not using $verified to keep button enabled
				$localUsers = Get-LocalGroup Administrators | Get-LocalGroupMember | Where-Object { $_.PrincipalSource -EQ "ActiveDirectory" -AND $_.ObjectClass -EQ "User" }
				$users = $localUsers.Count -GT 0
				if ($users) { Write-Output "User Added " $localUsers.Name; }
			}
			
			## BEST 
            "checkBD" {
				$verified = ([System.IO.DirectoryInfo]$env:ProgramFiles).GetDirectories("AntivirusProgram").Exists;
				if ($verified) { Write-Output "AntiVirusProgram is installed"; }
            }
			
			## Take Control Agent 
            "checkSW" {
				$verified = ([System.IO.DirectoryInfo]${env:ProgramFiles(x86)}).GetDirectories("RemoteAssistanceTool").Exists;              
				if ($verified) { Write-Output "RemoteAssistanceTool is installed"; }

            }

            default {
                Write-Output "Invalid Parameters.`r`n";
				Write-Output "Invalid verify attribute: $param";
                EXIT 1;
            }
        }

        if ($verified) { Write-Output "[DONE]" }
		else { Write-Output "[Check Fail]"; }
        EXIT 0
    }

    Main($args);