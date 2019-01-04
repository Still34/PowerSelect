# Disables Windows scripting engine
function Enable-WScript{
    Update-WScript $true
}
function Disable-WScript {
    Update-WScript $false
}
function Update-WScript {
    param(
        [bool]$Enable = $false
    )
    $registryPath = "HKLM:\Software\Microsoft\Windows Script Host\Settings"
    if ($flag) {
        $valuePath = Join-Path -Path $registryPath -ChildPath "Enabled"
        if (Test-Path $valuePath){
            Remove-ItemProperty -Path $valuePath
        }
    }
    else {
        if (!(Test-Path $registryPath)) {
            New-Item -Path $registryPath -Force
        }
        New-ItemProperty -Path $registryPath -Name "Enabled" -PropertyType DWORD -Value 0 -Force
    }
}
