# example from https://github.com/PandaWood/Simple-MAPI.NET#how-to-use

Import-Module $(Join-Path $PSScriptRoot "simple-mapi.net\lib\net40\SimpleMapi.NET.dll")
$SimpleMapi = New-Object Win32Mapi.SimpleMapi
$SimpleMapi.AddRecipient("bob@gmail.com", $null, $false)
$SimpleMapi.Attach($(Join-Path $PSScriptRoot "..\README.md"));
$result = $SimpleMapi.Send("a subject", "a body text")

if ($false -eq $result)
{
    $SimpleMapi.Error()
}
