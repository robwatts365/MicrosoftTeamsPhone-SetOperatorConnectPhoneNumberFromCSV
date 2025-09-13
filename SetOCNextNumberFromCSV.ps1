<# 
Set Operator Connect Next Available Number from CSV
    Version: v1.0
    Date: 12/11/2024
    Author: Rob Watts https://github.com/robwatts365

#>

# Pop out disclaimer
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

[Windows.Forms.MessageBox]::Show("
THIS CODE IS SAMPLE CODE. 

THESE SAMPLES ARE PROVIDED 'AS IS' WITHOUT WARRANTY OF ANY KIND.

MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. 

THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU.

IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.

BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.", "***DISCLAIMER***", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)

################### START OF GLOBAL VARIABLES ###################
$global:CityName = ""
################### END OF GLOBAL VARIABLES ###################

# Import Teams Module
Import-Module MicrosoftTeams

# Connects to Microsoft Teams
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams

# Enable File Picker
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

# File Picker  (Set File Path - Open File Browser)
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $FilePath = $OpenFileDialog.filename
   
# Store the data from NewUsersFinal.csv in the $ADUsers variable
    Write-Host "Importing CSV..."
    $Users = Import-Csv $FilePath

# Set Resource Account Name
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$CityNameMsg = 'City Name'
$CityNameTitle   = '"Please provide the city name for assigning numbers.'
$global:CityName = [Microsoft.VisualBasic.Interaction]::InputBox($CityNameTitle, $CityNameMsg)

# Loop through each row containing user details in the CSV file
foreach ($User in $Users) {

    # Read user data from each field in each row and assign the data to a variable as below
    $UPN = $User.UPN
    $TelephoneNumber = Get-CsPhoneNumberAssignment | Where-Object {$_.PstnAssignmentStatus -eq "Unassigned"} | Where-Object {$_.NumberType -eq "OperatorConnect"} | Where-Object {$_.City -like "*$global:CityName*"} | Select-Object -First 1
        
    Set-CSPhoneNumberAssignment -Identity $UPN -PhoneNumber $TelephoneNumber.TelephoneNumber -PhoneNumberType OperatorConnect

    }
