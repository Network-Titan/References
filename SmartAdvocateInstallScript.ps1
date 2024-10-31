New-Item -Path HKCU:\Software -Name SmartAdvocate

New-Item -Path HKCU:\Software\SmartAdvocate -Name OutlookAddin
New-Item -Path HKCU:\Software\SmartAdvocate -Name WordTemplateEditorAddin
New-Item -Path HKCU:\Software\SmartAdvocate -Name WordTemplateMergeAddIn

New-ItemProperty -Path HKCU:\Software\SmartAdvocate\OutlookAddin -Name "UseWebService" -Type DWORD -Value 00000001
New-ItemProperty -Path HKCU:\Software\SmartAdvocate\OutlookAddin -Name "WebServiceURL" -Type ExpandString -Value "https://app.smartadvocate.com/SASVC/SAPluginService.svc"

New-ItemProperty -Path HKCU:\Software\SmartAdvocate\WordTemplateEditorAddin -Name "UseWebService" -Type DWORD -Value 00000001
New-ItemProperty -Path HKCU:\Software\SmartAdvocate\WordTemplateEditorAddin -Name "WebServiceURL" -Type ExpandString -Value "https://app.smartadvocate.com/SASVC/SAPluginService.svc"

New-ItemProperty -Path HKCU:\Software\SmartAdvocate\WordTemplateMergeAddIn -Name "UseWebService" -Type DWORD -Value 00000001
New-ItemProperty -Path HKCU:\Software\SmartAdvocate\WordTemplateMergeAddIn -Name "WebServiceURL" -Type ExpandString -Value "https://app.smartadvocate.com/SASVC/SAPluginService.svc"


# Get the current script directory
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Define the file paths
$file1 = Join-Path -Path $scriptDir -ChildPath "MergeAddin.exe"
$file2 = Join-Path -Path $scriptDir -ChildPath "TemplateEditor.exe"
$file3 = Join-Path -Path $scriptDir -ChildPath "OutlookPlugin.exe"

# Launch each file with a prompt
Write-Host "Press any key to install MergeAddin.exe..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Start-Process -FilePath $file1

Write-Host "Press any key to install TemplateEditor.exe..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Start-Process -FilePath $file2

Write-Host "Press any key to install OutlookPlugin.exe..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
Start-Process -FilePath $file3

Write-Host "Installation complete. Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
exit 0
