

<# 
   Add Notes
    ...
#>


Function Process-NuSureTransactions() {
    # Open Excel file
    $excel = new-object -comobject excel.application
    $filePath = "C:\Native Data\Control Center\NuSure\NuSure Automation Control File.xlsm"
    $workbook = $excel.Workbooks.Open($FilePath)
    $excel.Visible = $true
    $worksheet = $workbook.worksheets.item(1)
    $excel.Run("NUTS_MasterFileGenerator")
    #$workbook.save()
    $workbook.close()
    $excel.quit()
    Write-Host "This Week's NuSure Transactions Are Ready"
}

Function Upload-NutrisavingsTransactions {
$DateStamp = ((Get-Date).ToUniversalTime()).ToString("MMddyyy")
$SFTPusername = "sftpnativesun"
$encrypted = Get-Content "C:\\Scripts\Data\ScriptsEncrypted_Password_1.txt" | ConvertTo-SecureString
$Credentials = New-Object System.Management.Automation.PsCredential($SFTPusername, $encrypted)
$Session = New-SFTPSession -Computername sftp.nutrisavings.com -Credential $Credentials
$TransactionFileName = "NativeSun_NS_transaction_" +  $DateStamp + "_101523.txt"
$FilePath = "C:\Native Data\NuSure\prepped_transaction_files\" + $TransactionFileName




Set-SFTPfile -SFTPSession $Session  -RemotePath "/Out/" -LocalFile $FilePath -Overwrite
$Session.Disconnect()

Write-Host  The file was successfully uploaded to the Nutrisavings SFTP Server

}

# Set and encrypt credentials to file using default method #"



Process-NuSureTransactions
Upload-NutrisavingsTransactions















