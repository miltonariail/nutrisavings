

<# 
   Working As Is ; Need to Dial in the Date Code to make dynamic
    ...
#>

Function Download-TrackingFile() {
$time = (Get-Date).AddDays(-3)
$Path = "C:\Native Data\NuSure\tracking_files\"
$DateStamp = ((Get-Date).ToUniversalTime()).ToString("MMddyyy")
$TestFile1 = "NS_NativeSun_tracking_" +  $DateStamp + "_010221.txt"
$TestFile2 = "NS_NativeSun_tracking_" +  $DateStamp + "_010222.txt"
$TextFile = "NS_NativeSun_tracking_" + $DateStamp + "_WORKING.txt"
$ExcelFile = "New_NuSure_Members_WORKING.xlsx"

$SFTPusername = "sftpnativesun"
$encrypted = Get-Content "C:\\Scripts\Data\ScriptsEncrypted_Password_1.txt" | ConvertTo-SecureString
$Credentials = New-Object System.Management.Automation.PsCredential($SFTPusername, $encrypted)
$Session = New-SFTPSession -Computername sftp.nutrisavings.com -Credential $Credentials


if(Test-SFTPPath -SFTPSession $Session -Path "/Out/$TestFile1")
{Get-SFTPfile -SFTPSession $Session -RemoteFile "/Out/$TestFile1" -LocalPath $Path -Overwrite
$TrackingFile = $TestFile1
Write-Host /Out/$TrackingFile}



if(Test-SFTPPath -SFTPSession $Session -Path "/Out/$TestFile2")
{Get-SFTPfile -SFTPSession $Session -RemoteFile "/Out/$TestFile2" -LocalPath $Path -Overwrite
$TrackingFile = $TestFile2
Write-Host /Out/$TrackingFile}

$Session.Disconnect()

Write-Host $Path$TrackingFile

Rename-Item -Path $Path$TrackingFile -NewName $Path$TextFile 
Export-Excel -Path $Path$ExcelFile -WorksheetName "New Members"
}

Function Update-NuSure_Membership() {
    # Open Excel file
    $excel = new-object -comobject excel.application
    $filePath = "C:\Native Data\Control Center\NuSure\Nutrisavings_Membership_Data.xlsm"
    $workbook = $excel.Workbooks.Open($FilePath)
    $excel.Visible = $true
    $worksheet = $workbook.worksheets.item(1)
    $excel.Run("NuSu_MembershipMaster")
    #$workbook.save()
    $workbook.close()
    $excel.quit()
    Write-Host "NuSure Membership Status Email Sent"
}





Download-TrackingFile
Update-NuSure_Membership

















