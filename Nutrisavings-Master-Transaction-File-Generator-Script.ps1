

<# 
   Working As Is ; Need to Dial in the Date Code to make dynamic
    ...
#>

Function Download-NutrisavingsTracking() {
$time = (Get-Date).AddDays(-3)
$Path = "C:\Native Data\Nutrisavings\tracking_files\NS_NativeSun_Tracking_*.txt"
$DateStamp = ((Get-Date).ToUniversalTime()).ToString("MMddyyy")
$TestFile = "NS_NativeSun_tracking_" +  $DateStamp + "_010221.txt"

$SFTPusername = "sftpnativesun"
$encrypted = Get-Content "C:\\Scripts\Data\ScriptsEncrypted_Password_1.txt" | ConvertTo-SecureString
$Credentials = New-Object System.Management.Automation.PsCredential($SFTPusername, $encrypted)
$Session = New-SFTPSession -Computername sftp.nutrisavings.com -Credential $Credentials

Get-SFTPfile -SFTPSession $Session -RemoteFile "/Out/$TestFile" -LocalPath "C:\Native Data\Nutrisavings\tracking_files\" -Overwrite

$Session.Disconnect()
}

Function Update-NutrisavingsMembers() {
    # Open Excel file
    $excel = new-object -comobject excel.application
    $filePath = "C:\Native Data\Control Center\Nutrisavings\Instacart Automation Control File.xlsm"
    $workbook = $excel.Workbooks.Open($FilePath)
    $excel.Visible = $true
    $worksheet = $workbook.worksheets.item(1)
    $excel.Run("InstacartTax_FileFormatMaster")
    #$workbook.save()
    $workbook.close()
    $excel.quit()
    Write-Host "Tax File Processed And Converted to XLSX"
}


Function Process-NutrisavingsTrasactions() {
    # Open Excel file
    $excel = new-object -comobject excel.application
    $filePath = "C:\Native Data\Control Center\Nutrisavings\Nutrisavings Automation Control File.xlsm"
    $workbook = $excel.Workbooks.Open($FilePath)
    $excel.Visible = $true
    $worksheet = $workbook.worksheets.item(1)
    $excel.Run("NUTS_MasterFileGenerator")
    #$workbook.save()
    $workbook.close()
    $excel.quit()
    Write-Host "This Week's Nutrisavings Transactions Are Ready"
}


# Download-TrackingFile

Process-NutrisavingsTrasactions







