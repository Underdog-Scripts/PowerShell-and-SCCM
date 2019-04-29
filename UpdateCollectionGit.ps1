<#
.Synopsis
   This script allows for changes to SCCM collections
.DESCRIPTION
   This script allows for changes to SCCM collections by adding and removing from user and device collections
.PARAMETER List of device or user names
   Reads comma separated file containing a list of devices or user names to be added or removed from specific collection
.INPUTS
   Path to comma separated file to be read
.OUTPUTS
   Log file of each actions result
.NOTES
   Version:        1.0
   Author:         DaAnswerIs
   Creation Date:  12.11.18
   Purpose/Change: Initial script development
.EXAMPLE
   ./UpdateCollection.ps1
.EXAMPLE
   ./UpdateCollection.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ConfigMgr 3 character site code and site server
$SiteCode = "P01" 
$SiteServer = "YOUR SCCM SERVER FQN"

# Default input file and log file paths
$DefaultFilePref = "\\YOUR PATH TO INPUT AND LOG FILES\UpdateCollection"

# Default input file and log file paths from menu choices
$AddUserFilePref = "\\YOUR PATH TO INPUT AND LOG FILES\Add-Users"
$RemUserFilePref = "\\YOUR PATH TO INPUT AND LOG FILES\Remove-Users"
$AddDevFilePref = "\\YOUR PATH TO INPUT AND LOG FILES\Add-Devices"
$RemDevFilePref = "\\YOUR PATH TO INPUT AND LOG FILES\Remove-Devices"
$PH1FilePref = "\\YOUR PATH TO INPUT AND LOG FILES\PH1"
$PH2FilePref = "\\YOUR PATH TO INPUT AND LOG FILES\PH2"

# Default input file and log file extension
$UserAndComputers = Import-Csv -LiteralPath $DefaultFilePref".csv"
$option = [System.StringSplitOptions]::RemoveEmptyEntries

#Logfile information
$logfile = $DefaultFilePref+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

#######################################################################################################################

Function Log {
   param([Switch]$fout, [String]$text)
   AC $logfile $text
   if($showConsoleOutput) {
      if($fout) {     
      } else {
            
      }
      
   }

}

#######################################################################################################################

# AddUserToCollection function - (from MassAdd-UsersandComputersToCollections.ps1)
Function AddUserToCollection {
   Param ([string]$UserA, [string]$ColID)
   Trap { Return "Input File Value Error" }
   Try {
      $CMUser = Get-CMUser -Name $UserA
      If ($CMUser.ResourceID) {
      $CMCollection = Get-CMUserCollection -Id $ColID
         If ($CMCollection.Name) {
            Try {
               Add-CMUserCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID
               CD C:
               $LogText =  "User $UserA was added to Collection $ColID"
               log -text $LogText
               CD $SiteCode":"
            } Catch {
                
               CD C:
               $LogText =  "User $UserA is already in Collection $ColID"
               log -text $LogText
               CD $SiteCode":"
            }

         } Else {
               
            CD C:
            $LogText =  "User Collection was not found: $ColID"
            log -text $LogText              
            CD $SiteCode":"
         }

      } Else {
         
         CD C:
         $LogText =  "User not found: $UserA"
         log -text $LogText
         CD $SiteCode":"
      }
          
   } Catch {
      
      CD C:
      $LogText = "ERROR: User not found: $UserA"
      log -text $LogText
      CD $SiteCode":"
               
   }

}

#######################################################################################################################
             
Function RemoveUserFromCollection {
   Param ([string]$UserA, [string]$ColID)
   Trap { Return "error" }
   Try {
      $CMUser = Get-CMUser -Name $UserA
      If ($CMUser.ResourceID) {
         $CMCollection = Get-CMUserCollection -Id $ColID
         If ($CMCollection.Name) {
                   
            Remove-CMUserCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMUser.ResourceID -Force
            CD C:
            $LogText =  "User $UserA was removed from Collection $ColID"
            log -text $LogText
            CD $SiteCode":"
         } Else {
               
            CD C:
            $LogText = "User not found: $UserA"
            log -text $LogText
            CD $SiteCode":"
         }

      } Else {

         CD C:
         $LogText = "User not found: $UserA"
         log -text $LogText
         CD $SiteCode":"
      }

   } Catch {
        
      CD C:
      $LogText = "ERROR: User not found: $UserA"
      Log -text $LogText
      CD $SiteCode":"
               
   }

}

#######################################################################################################################

Function AddDeviceToCollection {
   Param ([string]$DeviceA, [string]$ColID)
   Trap { Return "error" }
   Try {
      $CMDevice = Get-CMDevice -Name $DeviceA
      If ($CMDevice.ResourceID) {
         $CMCollection = Get-CMDeviceCollection -Id $ColID
         If ($CMCollection.Name) {
            Try {
               Add-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMDevice.ResourceID
               CD C:
               $LogText = "Device $DeviceA was added to Collection $ColID"
               log -text $LogText
               CD $SiteCode":"
            } Catch {
                     
               CD C:
               $LogText = "Device $DeviceA not in Collection $ColID"
               Log -text $LogText
               CD $SiteCode":"
            }

         } Else {
                  
            CD C:
            $LogText = "Device Collection was not found: $ColID"
            log -text $LogText              
            CD $SiteCode":"
         }

      } Else {
               
         CD C:
         $LogText = "Device not found: $DeviceA"
         log -text $LogText
         CD $SiteCode":"
      }

    } Catch {
        
      CD C:
      $LogText = "ERROR: Device not found: $DeviceA"
      Log -text $LogText
      CD $SiteCode":"
   }

}

#######################################################################################################################

Function RemoveDeviceFromCollection {
   Param ([string]$DeviceA, [string]$ColID)
   Trap { Return "error" }
   Try {
      $CMDevice = Get-CMDevice -Name $DeviceA
      If ($CMDevice.ResourceID) {
         $CMCollection = Get-CMDeviceCollection -Id $ColID
         If ($CMCollection.Name) {
            Remove-CMDeviceCollectionDirectMembershipRule -CollectionId $ColID -ResourceId $CMDevice.ResourceID -Force
            CD C:
            $LogText = "Device $DeviceA was removed from Collection $ColID"
            log -text $LogText
            CD $SiteCode":"
         } Else {
                  
            CD C:
            $LogText = "Device Collection was not found: $ColID"
            log -text $LogText              
            CD $SiteCode":"
         }

      } Else {
               
         CD C:
         $LogText = "Device not found: $DeviceA"
         log -text $LogText
         CD $SiteCode":"
      }

   } Catch {
        
      CD C:
      $LogText = "ERROR: Device not found: $DeviceA"
      Log -text $LogText
      CD $SiteCode":"
                
   }

}  

#######################################################################################################################

# A function to create the log file path
Function DisplayLogPath {
   Add-Type -AssemblyName System.Windows.Forms
   Add-Type -AssemblyName System.Drawing
   $form0 = New-Object System.Windows.Forms.Form
   $Font = New-Object System.Drawing.Font("Ariel",10,[System.Drawing.FontStyle]::Regular)
   $Form0.Font = $Font
   $Label0 = New-Object System.Windows.Forms.Label
   $Label0.Text = 'Script Complete... Please review log file below'
   $Label0.Size = New-Object System.Drawing.Size(600,20)
   $Label0.Location = New-Object System.Drawing.Point(10,8)
   $form0.Text = 'Log File Path'
   $form0.Size = New-Object System.Drawing.Size(1000,106)
   $form0.StartPosition = 'CenterScreen'
   $textBox9 = New-Object System.Windows.Forms.TextBox
   $textBox9.Location = New-Object System.Drawing.Point(10,30)
   $textBox9.Size = New-Object System.Drawing.Size(960,20)
   $textBox9.Text = $logfile
   $form0.Controls.Add($textBox9)
   $form0.Controls.Add($Label0)
   $form0.Topmost = $true
   $form0.Add_Shown({$textBox9.Select()})
   $result = $form0.ShowDialog()
}

#######################################################################################################################

# A function to create the form 
function Update_Collection{
   [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
   [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    
   # Set the size of the form
   $Form = New-Object System.Windows.Forms.Form
   $Form.width = 900
   $Form.height = 460
   $Form.Text = 'Please select a task'

   # Create label to display SCCM site code
   $label0 = New-Object System.Windows.Forms.Label
   $label0.Location = New-Object System.Drawing.Point(140,20)
   $label0.Size = New-Object System.Drawing.Size(120,14)
   $label0.Text = 'Site Code = P01'
   $form.Controls.Add($label0)

   # Create label to display SCCM site server
   $label1 = New-Object System.Windows.Forms.Label
   $label1.Location = New-Object System.Drawing.Point(460,20)
   $label1.Size = New-Object System.Drawing.Size(240,14)
   $label1.Text = 'Site Server = YOUR SCCM SERVER FQN'
   $form.Controls.Add($label1)
 
   # Create text for fist button group
   $label3 = New-Object System.Windows.Forms.Label
   $label3.Location = New-Object System.Drawing.Point(40,68)
   $label3.Size = New-Object System.Drawing.Size(360,14)
   $label3.Text = 'Using (Name, Type, Action, CollectionID) CVS file format... Example ->'
   $form.Controls.Add($label3)

   # Add picture of Excel file content
   $img = [System.Drawing.Image]::Fromfile('YOUR IMAGE OF EXCEL FILE SAMPLE\UpdateCollection.jpg')
   $pictureBox = new-object Windows.Forms.PictureBox
   $picturebox.Location = New-Object System.Drawing.Point(400,46)
   $pictureBox.Width = $img.Size.Width
   $pictureBox.Height = $img.Size.Height
   $pictureBox.Image = $img
   $form.controls.add($pictureBox)
         
   # Create underline for group name   
   $label3a = New-Object System.Windows.Forms.Label
   $label3a.Location = New-Object System.Drawing.Point(40,71)
   $label3a.Size = New-Object System.Drawing.Size(368,20)
   $label3a.Text = '___________________________________________________________________'
   $form.Controls.Add($label3a)
   
   # Create label for text file path
   $label4 = New-Object System.Windows.Forms.Label
   $label4.Location = New-Object System.Drawing.Point(40,120)
   $label4.Size = New-Object System.Drawing.Size(280,20)
   $label4.Text = 'Please enter/verify the path to the CSV file below:'
   $form.Controls.Add($label4)

   # Add default file - Radio Button
   $RadioButton0 = New-Object System.Windows.Forms.RadioButton
   $RadioButton0.Location = New-Object System.Drawing.Point(40,140)
   $RadioButton0.Size = New-Object System.Drawing.Size(20,20)
   $RadioButton0.Checked = $false 
   $RadioButton0.Text = ''
   
   # Create input box for default file path
   $textBox0 = New-Object System.Windows.Forms.TextBox
   $textBox0.Location = New-Object System.Drawing.Point(60,140)
   $textBox0.Size = New-Object System.Drawing.Size(640,20)
   $textBox0.Text = $DefaultFilePref+".csv"
   $form.Controls.Add($textBox0)
   
   # Create text for second button group
   $label2 = New-Object System.Windows.Forms.Label
   $label2.Location = New-Object System.Drawing.Point(40,197)
   $label2.Size = New-Object System.Drawing.Size(296,14)
   $label2.Text = 'Using (Name, CollectionID) CVS file format... Example ->'
   $form.Controls.Add($label2)

   # Add picture of Excel file content
   $img = [System.Drawing.Image]::Fromfile('YOUR IMAGE OF EXCEL FILE SAMPLE\Add-Devices.jpg')
   $pictureBox0 = new-object Windows.Forms.PictureBox
   $picturebox0.Location = New-Object System.Drawing.Point(338,176)
   $pictureBox0.Width = $img.Size.Width
   $pictureBox0.Height = $img.Size.Height
   $pictureBox0.Image = $img
   $form.controls.add($pictureBox0)

   # Create underline for group name
   $label2a = New-Object System.Windows.Forms.Label
   $label2a.Location = New-Object System.Drawing.Point(40,200)
   $label2a.Size = New-Object System.Drawing.Size(300,20)
   $label2a.Text = '________________________________________________'
   $form.Controls.Add($label2a)
      
   # Add Users to a Collection - Radio Button
   $RadioButton1 = New-Object System.Windows.Forms.RadioButton
   $RadioButton1.Location = New-Object System.Drawing.Point(40,250)
   $RadioButton1.size = New-Object System.Drawing.Size(190,20)
   $RadioButton1.Checked = $false 
   $RadioButton1.Text = 'Add Users to a Collection'
         
   # Create input box for add Users to a Collection file path
   $textbox1 = New-Object System.Windows.Forms.TextBox
   $textbox1.Location = New-Object System.Drawing.Point(240,250)
   $textbox1.Size = New-Object System.Drawing.Size(610,20)
   $textbox1.Text = $AddUserFilePref+".csv"
   $form.Controls.Add($textbox1)
   
   # Remove Users from a Collection - Radio Button
   $RadioButton2 = New-Object System.Windows.Forms.RadioButton
   $RadioButton2.Location = New-Object System.Drawing.Point(40,270)
   $RadioButton2.size = New-Object System.Drawing.Size(190,20)
   $RadioButton2.Checked = $false 
   $RadioButton2.Text = 'Remove Users from a Collection'

   # Create input box for remove Users from a Collection file path
   $textbox2 = New-Object System.Windows.Forms.TextBox
   $textbox2.Location = New-Object System.Drawing.Point(240,270)
   $textbox2.Size = New-Object System.Drawing.Size(610,20)
   $textbox2.Text = $RemUserFilePref+".csv"
   $form.Controls.Add($textbox2)

   # Add Devices to a Collection - Radio Button
   $RadioButton3 = New-Object System.Windows.Forms.RadioButton
   $RadioButton3.Location = New-Object System.Drawing.Point(40,290)
   $RadioButton3.size = New-Object System.Drawing.Size(190,20)
   $RadioButton3.Checked = $false
   $RadioButton3.Text = 'Add Devices to a Collection'

   # Create input box for add Devices to a Collection file path
   $textbox3 = New-Object System.Windows.Forms.TextBox
   $textbox3.Location = New-Object System.Drawing.Point(240,290)
   $textbox3.Size = New-Object System.Drawing.Size(610,20)
   $textbox3.Text = $AddDevFilePref+".csv"
   $form.Controls.Add($textbox3)

   # Remove Device from a Collection - Radio Button
   $RadioButton4 = New-Object System.Windows.Forms.RadioButton
   $RadioButton4.Location = New-Object System.Drawing.Point(40,310)
   $RadioButton4.size = New-Object System.Drawing.Size(190,20)
   $RadioButton4.Checked = $false
   $RadioButton4.Text = 'Remove Device from a Collection'

   # Create input box for remove Device from a Collection file path
   $textbox4 = New-Object System.Windows.Forms.TextBox
   $textbox4.Location = New-Object System.Drawing.Point(240,310)
   $textbox4.Size = New-Object System.Drawing.Size(610,20)
   $textbox4.Text = $RemDevFilePref+".csv"
   $form.Controls.Add($textbox4)

   # Placeholder 1 - Radio Button
   $RadioButton5 = New-Object System.Windows.Forms.RadioButton
   $RadioButton5.Location = New-Object System.Drawing.Point(40,330)
   $RadioButton5.size = New-Object System.Drawing.Size(190,20)
   $RadioButton5.Checked = $false
   $RadioButton5.Text = 'Placeholder 1'

   # Create input box for placeholder 1 file path
   $textbox5 = New-Object System.Windows.Forms.TextBox
   $textbox5.Location = New-Object System.Drawing.Point(240,330)
   $textbox5.Size = New-Object System.Drawing.Size(610,20)
   $textbox5.Text = $PH1FilePref+".csv"
   $form.Controls.Add($textbox5)

   # Placeholder 2 - Radio Button
   $RadioButton6 = New-Object System.Windows.Forms.RadioButton
   $RadioButton6.Location = New-Object System.Drawing.Point(40,350)
   $RadioButton6.size = New-Object System.Drawing.Size(190,20)
   $RadioButton6.Checked = $false
   $RadioButton6.Text = 'Placeholder 2'

   # Create input box for placeholder 2 file path
   $textbox6 = New-Object System.Windows.Forms.TextBox
   $textbox6.Location = New-Object System.Drawing.Point(240,350)
   $textbox6.Size = New-Object System.Drawing.Size(610,20)
   $textbox6.Text = $PH2FilePref+".csv"
   $form.Controls.Add($textbox6)

   # Add an OK button
   $OKButton = new-object System.Windows.Forms.Button
   $OKButton.Location = '335,384'
   $OKButton.Size = '75,23'
   $OKButton.Text = 'OK'
   $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
   # Add a Cancel button
   $CancelButton = new-object System.Windows.Forms.Button
   $CancelButton.Location = '470,384'
   $CancelButton.Size = '75,23'
   $CancelButton.Text = 'Cancel'
   $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
   # Add all the Form controls 
   $form.Controls.AddRange(@($OKButton,$CancelButton))

   # Add all the GroupBox controls
   $form.Controls.AddRange(@($Radiobutton0,$Radiobutton1,$RadioButton2,$RadioButton3,$RadioButton4,$RadioButton5,$RadioButton6))
       
   # Assign the Accept and Cancel options in the form to the corresponding buttons
   $form.AcceptButton = $OKButton
   $form.CancelButton = $CancelButton
 
   # Activate the form
   $form.Add_Shown({$form.Activate()})    
    
   # Get the results from the button click
   $dialogResult = $form.ShowDialog()

#######################################################################################################################

      # If the OK button is selected 
      if ($dialogResult -eq "OK") {
         if ($RadioButton0.Checked) { 
            # If the button 0 is selected 00000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
            # Customizations
            $initParams = @{}
            $ProviderMachineName = "YOUR SCCM SERVER FQN"
               if((Get-Module ConfigurationManager) -eq $null) {
                  # Import the ConfigurationManager.psd1 module
                  Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
               }

               if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
                  # Connect to the site's drive if it is not already present
                  New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
               }
            
            #Inputfile information
            $UserAndComputers = Import-Csv -LiteralPath $Textbox0.text
            $option = [System.StringSplitOptions]::RemoveEmptyEntries
            #Logfile information
            $logfile = $Textbox0.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"
            # Set the current location to be the site code.
            Set-Location "$($SiteCode):\" @initParams
               Try {
                  # Set the current location to be the site code.
                  Set-Location "$($SiteCode):\" @initParams
                  CD C:
                  $LogText = "Connected to the site: $SiteCode"
                  log -text $LogText
               } Catch {
            
                  CD C:
                  $LogText = "ERROR connecting to the site: $SiteCode"
                  log -text $LogText
               }

            CD $SiteCode":"
            Foreach ($UCObject in $UserAndComputers) {
               Switch ($UCObject.TYPE) {
                  "USER" {
                  Switch ($UCObject.Action) {
                     "ADD" {
                     Try {
                        AddUserToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
                     } Catch {

                        CD C:
                        $LogText =  "ERROR Finding User: $UCObject.Name"
                        log -text $LogText
                        CD $SiteCode":"
                     }
                   
                  }

                     "REMOVE" {
                     Try {
                        RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
                     } Catch {
      
                        CD C:
                        $LogText =  "ERROR Finding User: $UCObject.Name"
                        log -text $LogText
                        CD $SiteCode":"
                     }

                  }
                   
               }

            }
               
                  "DEVICE" {
                  Switch ($UCObject.Action) {
                     "ADD" {
                     Try {
                        AddDeviceToCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID
                     } Catch {

                        CD C:
                        $LogText =  "ERROR Finding Device: $UCObject.Name"
                        log -text $LogText
                        CD $SiteCode":"
                     }
                  }
                    
                     "REMOVE" {
                     Try {
                        RemoveDeviceFromCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID
                     } Catch {
                           
                        CD C:
                        $LogText =  "ERROR Finding Device: $UCObject.Name"
                        log -text $LogText
                        CD $SiteCode":"
                     }
             
                  } 

               }

            }

         }

      }


      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished processing AD and updating the SCCM Schedule update"
      log -text $LogText
      $script:logfile = $Textbox0.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

      # 000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000

      } elseif ($RadioButton1.Checked) { 
         # If the button 1 is selected 11111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "YOUR SCCM SERVER FQN"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
         #Inputfile information
         $UserAndComputers = Import-Csv -LiteralPath $Textbox1.text
         $option = [System.StringSplitOptions]::RemoveEmptyEntries
         #Logfile information
         $logfile = $Textbox1.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"
         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams

            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               AddUserToCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID
            } Catch {

               CD C:
               $LogText = "ERROR Finding User: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }    

      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding users to AD and updating the SCCM Schedule update"
      log -text $LogText
      $script:logfile = $Textbox1.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

      # 11111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111111

      } elseif ($RadioButton2.Checked) { # If the button 2 is selected 22222222222222222222222222222222222222222222222222222222222222222222222222222222
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "YOUR SCCM SERVER FQN"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $UserAndComputers = Import-Csv -LiteralPath $Textbox2.text
            $option = [System.StringSplitOptions]::RemoveEmptyEntries
            
            #Logfile information
            $logfile = $Textbox2.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
         Try{
            # Set the current location to be the site code.
            Set-Location "$($SiteCode):\" @initParams
            CD C:
            $LogText = "Connected to the site: $SiteCode"
            log -text $LogText
         } Catch {
            
            CD C:
            $LogText = "ERROR connecting to the site: $SiteCode"
            log -text $LogText
         }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

            } Catch {
               CD C:
               $LogText = "ERROR Finding User: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished removing users from AD and updating the SCCM Schedule update"
      log -text $LogText
      $script:logfile = $Textbox2.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"
            
      # 22222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222222      

      } elseif ($RadioButton3.Checked) { # If the button 3 is selected 3333333333333333333333333333333333333333333333333333333333333333333333333333
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "YOUR SCCM SERVER FQN"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $UserAndComputers = Import-CSV -LiteralPath $Textbox3.text
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $Textbox3.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               AddDeviceToCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID
            } Catch {

               CD C:
               $LogText = "ERROR Finding device: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished adding devices to AD and updating the SCCM Schedule update"
      log -text $LogText
      $script:logfile = $Textbox3.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"   
               
      # 33333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333333

      } elseif ($RadioButton4.Checked) { # If the button 4 is selected 44444444444444444444444444444444444444444444444444444444444444444444444444444444444
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "YOUR SCCM SERVER FQN"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $UserAndComputers = Import-CSV -LiteralPath $Textbox4.text
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $Textbox4.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveDeviceFromCollection -DeviceA $UCObject.Name -ColID $UCObject.CollectionID

            } Catch {
               CD C:
               $LogText = "ERROR Finding device: $UCObject.Name"
               log -text $LogText
               CD $SiteCode":"
            }

         }
         
      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished removing devices from AD and updating the SCCM Schedule update"
      log -text $LogText
      $script:logfile = $Textbox4.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"
  
      # 44444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444444  
               
      } elseif ($RadioButton5.Checked) { # If the button 5 is selected 55555555555555555555555555555555555555555555555555555555555555555555555555555555555
         # Customizations
         $initParams = @{}
         $ProviderMachineName = "YOUR SCCM SERVER FQN"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
            }
            
            #Inputfile information
            $UserAndComputers = Import-CSV -LiteralPath $textbox5.Text
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $textbox5.Text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

             } Catch {
                CD C:
                $LogText = "ERROR Finding Device: $UCObject.Name"
                log -text $LogText
                CD $SiteCode":"
             }

         }
         
      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished ??????? AD and updating the SCCM Schedule update"
      log -text $LogText
      $Script:logfile = $Textbox5.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"
      
      # 55555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555555
          
      } elseif ($RadioButton6.Checked) { # If the button 6 is selected 66666666666666666666666666666666666666666666666666666666666666666666666666666666666
      # Customizations
         $initParams = @{}
         $ProviderMachineName = "YOUR SCCM SERVER FQN"
            if((Get-Module ConfigurationManager) -eq $null) {
               # Import the ConfigurationManager.psd1 module
               Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams -ErrorAction SilentlyContinue
            }

            if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
               # Connect to the site's drive if it is not already present
               New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams 
            }
            
            #Inputfile information
            $UserAndComputers = Import-CSV -LiteralPath $textbox6.Text
            $option = [System.StringSplitOptions]::RemoveEmptyEntries

            #Logfile information
            $logfile = $textbox6.Text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"

         # Set the current location to be the site code.
         Set-Location "$($SiteCode):\" @initParams
            Try{
               # Set the current location to be the site code.
               Set-Location "$($SiteCode):\" @initParams
               CD C:
               $LogText = "Connected to the site: $SiteCode"
               log -text $LogText
            } Catch {
            
               CD C:
               $LogText = "ERROR connecting to the site: $SiteCode"
               log -text $LogText
            }

         CD $SiteCode":"
         Foreach ($UCObject in $UserAndComputers) {
            Try {
               RemoveUserFromCollection -UserA $UCObject.Name -ColID $UCObject.CollectionID

             } Catch {
                CD C:
                $LogText =  "ERROR Finding Device: $UCObject.Name"
                log -text $LogText
                CD $SiteCode":"
             }

         }
         
      CD C:
      $LogText =  "$(Get-Date) Running as - $($env:USERNAME) on $($env:COMPUTERNAME) Finished ??????? AD and updating the SCCM Schedule update"
      log -text $LogText
      $script:logfile = $Textbox6.text+"_"+(Get-Date -Format MM-dd-yyyy_hh.mm.ss)+".log"
      
      # 66666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666666         
           
      } else {
      
         [System.Windows.Forms.MessageBox]::Show("No Radio Button Selected")
      
      }
      
   }

}

#######################################################################################################################

Function Reset-Log { 
   #function checks to see if file in question is larger than the paramater specified if it is it will roll a log and delete the oldes log if there are more than x logs. 
   param([string]$fileName, [int64]$filesize = 1mb , [int] $logcount = 5) 
   $logRollStatus = $true 
   if(test-path $filename) { 
      $file = Get-ChildItem $filename 
      if((($file).length) -ige $filesize) { #this starts the log roll  
         $fileDir = $file.Directory 
         $fn = $file.name #this gets the name of the file we started with 
         $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
         $filefullname = $file.fullname #this gets the fullname of the file we started with 
         #$logcount +=1 #add one to the count as the base file is one more than the count 
         for ($i = ($files.count); $i -gt 0; $i--) {  
            #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of the file we are on 
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            $operatingFile = $files | ?{($_.name).trim($fn) -eq $i} 
            if ($operatingfile) {
               $operatingFilenumber = ($files | ?{($_.name).trim($fn) -eq $i}).name.trim($fn)
            } else {
               
               $operatingFilenumber = $null
            } 
 
               if(($operatingFilenumber -eq $null) -and ($i -ne 1) -and ($i -lt $logcount)) { 
                  $operatingFilenumber = $i 
                  $newfilename = "$filefullname.$operatingFilenumber" 
                  $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                  move-item ($operatingFile.FullName) -Destination $newfilename -Force 
               } elseif($i -ge $logcount) {
                 
                  if($operatingFilenumber -eq $null) {  
                     $operatingFilenumber = $i - 1 
                     $operatingFile = $files | ?{($_.name).trim($fn) -eq $operatingFilenumber
                  } 
                        
               } 
                  remove-item ($operatingFile.FullName) -Force 
            } elseif($i -eq 1) {
                 
               $operatingFilenumber = 1 
               $newfilename = "$filefullname.$operatingFilenumber" 
               move-item $filefullname -Destination $newfilename -Force 
            } else {
                 
               $operatingFilenumber = $i +1  
               $newfilename = "$filefullname.$operatingFilenumber" 
               $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
               move-item ($operatingFile.FullName) -Destination $newfilename -Force    
            } 
                     
         } 
 
                     
      } else {
          
         $logRollStatus = $false 
      }
         
   } else {
     
      $logrollStatus = $false 
   }

      $LogRollStatus
} 

#######################################################################################################################

# Run the function Update_Collection above
Update_Collection

# Run the function to display complete screen
DisplayLogPath
