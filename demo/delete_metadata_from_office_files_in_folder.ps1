# By Igors with little modifications
# Source: https://journeymblog.wordpress.com/2017/01/10/powershell-script-for-removing-personal-info-from-word-excel-and-powerpoint-docs/
# Dependencies: Office primary interop assemblies. Installed with the Office or with Visual Studio
# dll's are located in "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15". After instalation of Visiual Studio, this folder can be copied and Visual Studio can be deleted.
$successlines=@()
$global:errorlinesexcel=@()
$global:errorlinesword=@()
$errorlinespowerpoint=@()
#------------------------------------------FUNCTIONS---------------------------------------------#
Function Select-FolderDialog
{
    param([string]$Description="Select Folder",[string]$RootFolder="Desktop")
 
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
     Out-Null     
 
   $objForm = New-Object System.Windows.Forms.FolderBrowserDialog
        $objForm.Rootfolder = $RootFolder
        $objForm.Description = $Description
        $Show = $objForm.ShowDialog()
        If ($Show -eq "OK")
        {
            Return $objForm.SelectedPath
        }
        Else
        {
            Write-Error "Operation cancelled by user."
        }
    }
 
function CheckforPasswordProtection ($obj){
$Binary = [System.IO.File]::ReadAllBytes($obj.FullName)
$Start = [System.Text.Encoding]::Default.GetString($Binary[0000..2000])
Switch ($obj.Extension){
    ".xls" {
        if($Start -match "E.n.c.r.y.p.t.e.d.P.a.c.k.a.g.e") {
            return $true
        }
        #if ($Binary[0x208]-eq 0xFE){
        #    return $true
        #    }
        if ($Binary[0x214]-eq 0x2F){
            return $true
            }
    }
    ".xlsx"{
        if($Start -match "E.n.c.r.y.p.t.e.d.P.a.c.k.a.g.e") {
            return $true
        }
        #if ($Binary[0x208]-eq 0xFE){
        #    return $true
        #    }
        if ($Binary[0x214]-eq 0x2F){
            return $true
            }
    }
    ".doc" {
        if($Start -match "E.n.c.r.y.p.t.e.d.P.a.c.k.a.g.e") {
            return $true
        }
        if ($Binary[0x20B]-eq 0x13){
                return $true
            }
    }
    ".docx"{
        if($Start -match "E.n.c.r.y.p.t.e.d.P.a.c.k.a.g.e") {
            return $true
        }
        if ($Binary[0x20B]-eq 0x13){
                return $true
        }
    }
}
}
 
function TestFileLock ($FilePath ){
    $FileLocked = $false
    $FileInfo = New-Object System.IO.FileInfo $FilePath
    trap {Set-Variable -name Filelocked -Value $true -scope 1; continue}
    $FileStream = $FileInfo.Open( [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None )
    if ($FileStream) {$FileStream.Close()}
    $FileLocked
}
 
function RemoveModificationProtectionsandPersonalInfoWord{
    $temp=$env:TEMP+"\"+$documents.Name
    try{
    $documents.SaveAs([ref] $temp, [ref] $null,[ref] $false ,[ref]'', [ref]$null,[ref]'',[ref] $false)
    }
    catch{
    Write-Host "File '$($obj.fullname)' saving failed, error: '$($_.Exception.Message)'" -ForegroundColor Red
    $global:errorlinesword+="'$($obj.fullname)' saving failed, error: '$($_.Exception.Message)'"
    }
    try{
    $documents.RemoveDocumentInformation($WdRemoveDocType::wdRDIAll)
    $documents.Save()
    }
    catch{
    Write-Host "File '$($obj.fullname)' clearing metadata failed" -ForegroundColor Red
    $global:errorlinesword+="'$($obj.fullname)' clearing metadata failed"
    }
    $documents.close()
    if (!$Error) {
    Move-Item -Path $temp -Destination $obj.FullName -Force
    } else {Remove-Item -Path $temp -Force}
    $temp=$null
}
function RemoveModificationProtectionsandPersonalInfoExcel{
    $temp=$env:TEMP+"\"+$documents.Name
    $documents.CheckCompatibility=$false
    try{
    $documents.SaveAs($temp,$documents.FileFormat,'','',$false)
    }
    catch{
    try{$objexcel.workbooks.close()}catch{}
    Write-Host "File '$($obj.fullname)' saving failed, error: '$($_.Exception.Message)'" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' saving failed, error: '$($_.Exception.Message)'"
    $temp=$null
    return
    }    
    try {
    $documents.RemoveDocumentInformation($XlRemoveDocType::xlRDIAll)
    $documents.Save()
    }
    catch{
    if ($_.Exception.Message -match "Cannot remove PII from this document because the document is signed, protected, shared, or marked as read-only" ){
    Write-Host "File '$($obj.fullname)' has protected sheets" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' has protected sheets"
    } else {
    try{$objexcel.workbooks.close()}catch{}
    Write-Host "File '$($obj.fullname)' saving failed, error: '$($_.Exception.Message)'" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' saving failed, error: '$($_.Exception.Message)'"
    }
    }
    try{$objexcel.workbooks.close()}catch{}
    if (!$Error) {
    Move-Item -Path $temp -Destination $obj.FullName -Force
    } else {try{Remove-Item -Path $temp -Force -ErrorAction Stop} catch {}}
    $temp=$null
}
#------------------------------------------SETTINGS---------------------------------------------#
# $path = Select-FolderDialog # choose folder dialor
#$path='C:\Temp'
#$path=(Get-Item -Path ".\" -Verbose).FullName
$path = $MyInvocation.MyCommand.Path
$path = Split-Path $path -Parent
$errorlog = $('{0}\Errors_{1}.txt' -f $path, $('{0:yyyy-MM-dd_HH-mm-ss}' -f (Get-Date)))
$successlog = $('{0}\Success_{1}.txt' -f $path, $('{0:yyyy-MM-dd_HH-mm-ss}' -f (Get-Date)))
if ($null -eq $path) {exit}
$pathToCurrentDirectory = $PWD.Path
Set-Location -Path "C:\Program Files (x86)\Microsoft Visual Studio\Shared\Visual Studio Tools for Office\PIA\Office15"
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
Add-Type -AssemblyName Microsoft.Office.Interop.Powerpoint
Set-Location -Path $pathToCurrentDirectory
$WdRemoveDocType = "Microsoft.Office.Interop.Word.WdRemoveDocInfoType" -as [type]
$XlRemoveDocType = "Microsoft.Office.Interop.Excel.XlRemoveDocInfoType" -as [type]
$PpRemoveDocType = "Microsoft.Office.Interop.PowerPoint.PpRemoveDocInfoType" -as [type]
$wordFiles = Get-ChildItem -Path $path -include *.doc, *.docx -Recurse 
$excelFiles = Get-ChildItem -Path $path -include *.xls, *.xlsx, *.xlsb, *.xlsm -Recurse
$powerpointfiles = Get-ChildItem -Path $path -include *.ppt, *.pptx -Recurse
#------------------------------------------WORD FILES---------------------------------------------#
if ($null -ne $wordFiles) {
$objword = New-Object -ComObject word.application
$objword.visible = $false 
Write-Host "Processing Word files"
foreach($obj in $wordFiles)
{
    if (TestFileLock $obj.FullName) {
    Write-Host "File '$($obj.fullname)' is locked" -ForegroundColor Red
    $global:errorlinesword+="'$($obj.fullname)' is locked"
    } else {
    if (CheckforPasswordProtection($obj) -eq $true){
    Write-Host "File '$($obj.fullname)' is password protected" -ForegroundColor Red
    $global:errorlinesword+="'$($obj.fullname)' is password protected"
    } else {
    $Error.Clear()
    try {
    $documents = $objword.Documents.Open($obj.fullname,$null,$true,$null,"11","",$false,"","",'wdOpenFormatAuto',$null,$false)
    }
    catch{
    if ($_.Exception.Message -match "The password is incorrect" ){
    Write-Host "File '$($obj.fullname)' has password for opening" -ForegroundColor Red
    $global:errorlinesword+="'$($obj.fullname)' has password for opening"
    }
    else {
    Write-Host "File '$($obj.fullname)' open failed, error '$($_.Exception.Message)'" -ForegroundColor Red
    $global:errorlinesword+="'$($obj.fullname)' open failed, error '$($_.Exception.Message)'"
    }
    }    
    if (!$Error) {RemoveModificationProtectionsandPersonalInfoWord}
    if (!$Error) {
    Write-Host "File '$($obj.fullname)' removed modification restrictions and cleared metadata" -ForegroundColor Green
    $successlines+="'$($obj.fullname)' removed modification restrictions and cleared metadata"
    }
    }
    }
}
$objword.Quit()
$documents=$null
}
 
#------------------------------------------EXCEL FILES---------------------------------------------#
if ($null -ne $excelFiles) {
$objexcel = New-Object -ComObject excel.application
$objexcel.visible = $false  
$culturebackup=Get-Culture
Set-Culture en-US
Write-Host "Processing Excel files"
foreach($obj in $excelFiles)
{
    if (TestFileLock $obj.FullName) {
    Write-Host "File '$($obj.fullname)' is locked" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' is locked"
    } else {
    if (CheckforPasswordProtection($obj) -eq $true){
    Write-Host "File '$($obj.fullname)' is password protected" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' is password protected"
    } else {
    $Error.Clear()
    try {
    $documents = $objexcel.Workbooks.Open($obj.fullname,$false,$true,5,'11','',$true)
    }
    catch{
    if ($_.Exception.Message -match "The password is incorrect" ){
    Write-Host "File '$($obj.fullname)' has password for opening" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' has password for opening"
    }
    else {
    Write-Host "File '$($obj.fullname)' open failed, error '$($_.Exception.Message)'" -ForegroundColor Red
    $global:errorlinesexcel+="'$($obj.fullname)' open failed, error '$($_.Exception.Message)'"
    }
    }
    if (!$Error) {RemoveModificationProtectionsandPersonalInfoExcel}
    if (!$Error) {
    Write-Host "File '$($obj.fullname)' removed modification restrictions and cleared metadata" -ForegroundColor Green
    $successlines+="'$($obj.fullname)' removed modification restrictions and cleared metadata"
    }
    }
    }
}
$objexcel.Quit()
Set-Culture $culturebackup
$documents=$null
}
#------------------------------------------POWERPOINT FILES---------------------------------------------#
if ($null -ne $powerpointfiles) {
$objpowerpoint = New-Object -ComObject Powerpoint.Application
Write-Host "Processing Powerpoint files"
foreach($obj in $powerpointfiles)
{
    if (TestFileLock $obj.FullName) {
    Write-Host "File '$($obj.fullname)' is locked" -ForegroundColor Red
    $errorlinespowerpoint+="'$($obj.fullname)' is locked"
    } else {
    $documents=$objpowerpoint.Presentations.Open($obj.FullName, $false, $null, $false)
#    Start-Sleep -s 2
    $documents.RemoveDocumentInformation($PpRemoveDocType::ppRDIAll) 
    $documents.Save() 
    $documents.Close()
    Write-Host "File '$($obj.fullname)' metadata cleared" -ForegroundColor Green
    $successlines+="'$($obj.fullname)' metadata cleared"
    }
}
$objpowerpoint.Quit()
$documents=$null
}
try {
$objpowerpoint.Quit()
$objexcel.Quit()
$objword.Quit()
} catch {write-host “Killing remaining active applications”}
#------------------------------------------LOG FILES---------------------------------------------#
# $global:errorlinesexcel | Out-File -FilePath $errorlog
# $global:errorlinesword | Out-File -FilePath $errorlog -Append
# $errorlinespowerpoint | Out-File -FilePath $errorlog -Append
# $successlines | Out-File -FilePath $successlog
#------------------------------------------SUMMARY---------------------------------------------#
Write-Host "Summary:" -ForegroundColor Red
Write-Host "Processed $($wordFiles.Count+$excelFiles.Count+$powerpointfiles.Count) files"
Write-Host "Excel files $($excelFiles.Count), Errors: $($global:errorlinesexcel.count)"
Write-Host "Word files $($wordFiles.Count), Errors: $($global:errorlinesword.count)"
Write-Host "Powerpoint files $($powerpointfiles.Count), Errors: $($errorlinespowerpoint.count)"
Write-Host "Success log is located: $($successlog)" -ForegroundColor Red
Write-Host "Error log is located: $($errorlog)" -ForegroundColor Red
$HOST.UI.RawUI.ReadKey(“NoEcho,IncludeKeyDown”) | OUT-NULL