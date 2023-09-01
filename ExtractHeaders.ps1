Get-ChildItem -Path .\ -Filter *.msg -Recurse | ForEach-Object {
    $FullPath = $_.FullName
    Write-Output "Headers from: $_`n`r"
    $outlook = New-Object -ComObject Outlook.Application
    $msg = $outlook.CreateItemFromTemplate($FullPath)
    $headers = $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
    Write-Output $headers 
} | Out-File -FilePath .\HeaderOutput.txt

