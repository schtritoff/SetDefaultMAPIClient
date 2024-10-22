<#
.SYNOPSIS
    Change default MAPI client via GUI

    References:
        https://stackoverflow.com/questions/54346779/is-there-any-ui-in-windows-10-to-change-the-current-mapi-provider
        https://docs.microsoft.com/en-us/powershell/scripting/samples/selecting-items-from-a-list-box?view=powershell-5.1
#>


# get default mapi for current user
# src: https://stackoverflow.com/a/43912590/1155121
$RegKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Clients\Mail"
$RegValue = "(Default)"
$mapi_current = (Get-ItemProperty -Path $RegKey -Name $RegValue -ErrorAction SilentlyContinue).$RegValue

# if undefined for current user, try for local machine
if ([string]::IsNullOrEmpty($mapi_current)) {
    # create registry path for current user since it doesn't exist
    New-Item -Path $RegKey -ErrorAction Continue

    # see what is set for local machine (but seems that this setting doesn't work in win11 - you NEED to have MAPI client set for current user)
    $RegKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Clients\Mail"
    $mapi_current = (Get-ItemProperty -Path $RegKey -Name $RegValue).$RegValue
}

# gui
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'MAPI Clients'
$form.Size = New-Object System.Drawing.Size(330,200)
$form.StartPosition = 'CenterScreen'

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(75,120)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(150,120)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = "Current MAPI client is '$mapi_current', set another:"
$form.Controls.Add($label)

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,40)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height = 80

# loop registry keys whose names are MAPI clients
(Get-ChildItem Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Clients\Mail) | ForEach-Object {
    $mapi_client = $_.Name.Substring($_.Name.LastIndexOf('\')+1)
    [void] $listBox.Items.Add($mapi_client)
}
# make current mapi client as selected
$listBox.SelectedItem = $mapi_current


$form.Controls.Add($listBox)
$form.Topmost = $true
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    #$mapi_outlook = "Microsoft Outlook"
    $mapi_current = $listBox.SelectedItem
    
    # set for current user
    Set-ItemProperty -Path Registry::\HKCU\SOFTWARE\Clients\Mail -Name "(Default)" -Value $mapi_current

    # set also for local machine (will work only if run as admin)
    Set-ItemProperty -Path Registry::\HKEY_LOCAL_MACHINE\SOFTWARE\Clients\Mail -Name "(Default)" -Value $mapi_current -ErrorAction Continue
}
