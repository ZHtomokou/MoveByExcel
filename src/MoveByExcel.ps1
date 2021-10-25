# Env
$Initialint = 2
$Endint = 5
# Env

# Open excel
$myDialog = New-Object System.Windows.Forms.OpenFileDialog
$myDialog.Title = "Please select a excel file"
$myDialog.InitialDirectory = Split-Path $MyInvocation.MyCommand.Path
$myDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
$result = $myDialog.ShowDialog()
If($result -eq "OK") {
    $theExcelDir = $myDialog.FileName
} else {
    Write-Output "Cancelled by user"
    Exit
}
# Open excel

# Move
$ExcelFile = Import-Excel -Path $theExcelDir
for ($num = $Initialint-2; $num -le $Endint-2; $num++) {
    $myIndex = $num
    $SourceDir = $ExcelFile[$myIndex].Source
    $DestDir = $ExcelFile[$myIndex].Dest

    if (Test-Path -Path $DestDir) {
        Write-Output "Moving $($SourceDir) to $($DestDir) ..."
        if (Test-Path -Path $SourceDir) {
            Move-Item -Path $SourceDir -Destination $DestDir
            Write-Output "Success."
        } else {
            Write-Output "$($SourceDir) not exist"
        }
    } else {
        Write-Output "Destination Directory $($DestDir) not exist!"
    }
}
# Move

