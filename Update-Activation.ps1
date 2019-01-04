function Update-OfficeActivation {
    param (
        [Parameter(Mandatory = $true)]
        [string]$KMSAddress
    )
    $isConnectable = Test-Connection $KMSAddress -Quiet
    if (!$isConnectable){
        throw New-Object System.Net.WebException
    }
    $officeActivationPaths = New-Object System.Collections.ArrayList
    for ($i = 4; $i -le 9; $i++) {
        $officeActivationScript = [System.IO.Path]::Combine($env:ProgramFiles, "Microsoft Office", "Office1" + $i, "ospp.vbs")
        if (Test-Path $officeActivationScript){
            $officeActivationPaths.Add($officeActivationScript)
        }
    }
    if ($officeActivationPaths.Count -eq 0){
        throw New-Object System.IO.FileNotFoundException
    }
    $officeActivationPaths | %{
        "Activating via " + $_
        cscript.exe $_ //B /osppsvcrestart
        cscript.exe $_ //B /sethst:$KMSAddress
        cscript.exe $_ //B /sestprt:1688
        cscript.exe $_ //NoLogo /act
    }
}
function Update-WindowsActivation{
    param (
        [Parameter(Mandatory = $true)]
        [string]$KMSAddress
    )
    $isConnectable = Test-Connection $KMSAddress -Quiet
    $fullAddress = $KMSAddress + ":1688"
    if (!$isConnectable){
        throw New-Object System.Net.WebException
    }
    $path = [System.IO]::Combine($env:SystemRoot, "system32", "slmgr.vbs")
    if (Test-Path $path)
    {
        Start-Process -FilePath cscript.exe -ArgumentList $path, "-skms", $fullAddress
        Start-Process -FilePath cscript.exe -ArgumentList $path, "-ato"
    }
    throw New-Object System.IO.FileNotFoundException
}