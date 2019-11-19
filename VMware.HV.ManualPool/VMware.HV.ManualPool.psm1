<#
    PowerShell Script Module:       VMware.HV.ManualPool
    Version:                        1.0.0
    Date:                           November 13th, 2019
    Author:                         James Wood
    Author email:                   woodj@vmware.com

    
    Permission is hereby granted, free of charge, to any person obtaining a copy of
    this software and associated documentation files (the "Software"), to deal in
    the Software without restriction, including without limitation the rights to
    use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
    of the Software, and to permit persons to whom the Software is furnished to do
    so, subject to the following conditions:
    
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>

function Show-UserChoices {
    <#
    .synopsis
    Displays a text menu prompting user for a choice.
    .description
    Displays a text based menu showing options for the user to choose to continue the script.
    .parameter InputObject
    Object containing the array of items to create the list out of.
    .parameter LabelProperty
    Name of the property of the -InputObject to use as the label for the list items.
    The property needs to be entered with a comma ',' in place of the dot '.' and omitting the first dot.
    See -Examples for an example
    .parameter Caption
    The caption (top line) for the user prompt
    .parameter Message
    The message (bottom line) for the user prompt
    .parameter DefaultChoice
    Index of the choice to use as the default choice.  Pass -1 for no default.
    .outputs
    [Int32] Index of the select option.
    .example
    Show-UserChoices -InputObject $object -LabelProperty foo,bar
    This will construct the expression $object.foo.bar
    #>
    param (
        [Parameter(Mandatory=$true)]
        $InputObject,
        [Parameter(Mandatory=$true)]
        [string[]]$LabelProperty,
        [string]$Caption = "Please make a selection",
        [string]$Message = "Type or copy/paste the text of your choice",
        [int]$DefaultChoice = -1
    )
    $choices = @()
    for($i=0;$i -lt $InputObject.Count;$i++){
        $exp = '$InputObject[$i]'
        foreach ($j in $LabelProperty){
            $exp += '.' + $j
        }
        $choices += [System.Management.Automation.Host.ChoiceDescription]("$(Invoke-Expression $exp)")
    }
    $userChoice = $host.UI.PromptForChoice($Caption,$Message,$choices,$DefaultChoice)
    return $userChoice
}
function Disconnect-Sessions {
    <#
    .synopsis
    ** Private Function intended for use from within the Update-DesktopPool cmdlet **
    Disconnects active sessions to vCenter and the Horizon View Connection Server
    .description
    ** Private Function intended for use from within the Update-DesktopPool cmdlet **
    Disconnects active sessions to vCenter and the Horizon View Connection Server
    .parameter err
    Array of errors encountered.
    .parameter vc
    vCenter object
    .parameter hv
    Horizon View Connection Server object
    #>

    param (
        $err,
        $vc,
        $hv
    )
    if ($vc){
        Disconnect-VIServer $vc -Confirm:$false
        Write-Output "Disconnected from vCenter Server."
    }
    if ($hv){
        Disconnect-HVServer $hv -Confirm:$false
        Write-Output "Disconnected from Horizon View Connection Server"
    }
    if ($err){
        Write-Warning "Errors have been encountered."
        ForEach-Object -InputObject $err {
            Write-Error $_.Exception.Message
        }
        $scriptPath = Split-Path -Parent $PSCommandPath
        $err | ConvertTo-Csv | Out-File -FilePath $scriptPath\errors.csv
        Write-Warning "The error details have been written to $scriptPath\errors.csv."
    }
    #exit
}
function New-RandomVMName {
    <#
    .synopsis
    ** Private Function intended for use from within the Update-DesktopPool cmdlet **
    Creates a random name for a new virtual desktop
    .description
    ** Private Function intended for use from within the Update-DesktopPool cmdlet **
    Creates a random name for a new virtual desktop
    #>
    #DESKTOP-3JLJH6S
    $randomName = "DESKTOP-" + -join ((48..57) + (65..90) | Get-Random -count 7 | ForEach-Object {[char]$_})
    return $randomName  
}
function Update-DesktopPool {
    <#
    .Synopsis
    Replaces all desktops in a manual, full-clone pool
    
    .Description
    Deletes all existing desktops and replaces with new, full-clone desktops based on input parameters

    .Parameter vCenterServer
    A string value for the FQDN or IP address of the vCenter server

    .Parameter HvServer
    A string value for the FQDN or IP address of the Horizon Connection Server

    .Parameter credential
    A System.Management.Automation.PSCredential object representing the user credentials.  If this is not provided the user will be prompted for it.

    #>

    param(
        [string]$vCenterServer = (Read-Host -Prompt "Enter vCenter IP or FQDN"),
        [string]$HvServer = (Read-Host -Prompt "Enter Horizon Connection Server IP or FQDN"),
        [System.Management.Automation.PSCredential]$credential = (Get-Credential -Message "Enter domain\username and Password.")
    )

#Region Connect to vCenter
Write-Output "Connecting to vCenter - $vCenterServer... "
do{
    try{
        $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential -ErrorAction Stop
    }
    Catch [VMware.VimAutomation.Sdk.V1.ErrorHandling.VimException.ViServerConnectionException]{
        Write-Warning "*** Error connecting to $vCenterServer. ***"
        Write-Output "Please re-enter the vCenter IP or FQDN."
        $vCenterServer = Read-Host -Prompt "Enter vCenter IP or FQDN"
    }
    Catch [VMware.VimAutomation.ViCore.Types.V1.ErrorHandling.InvalidLogin]{
        Write-Warning "Invalid login credentials."
        Write-Output "Please re-enter your domain\username and password."
        $credential = Get-Credential -Message "Enter your domain\username and password."
    }
    Catch{
        Write-Warning "An unexpected error has occurred. Please see the error log (errors.csv) for details."
        Write-Output "The script will now exit."
        Disconnect-Sessions -err $_ -vc $objvcenter
    }
}
until ($null -ne $objvcenter)
Write-Output "... Connected!"
Write-Output ""
#EndRegion Connect to vCenter

#Region Connect to Horizon Connection Server
Write-Output "Connecting to Horizon Connection Server - $HvServer..."
do{
    try{
        $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential -ErrorAction Stop
    }
    catch{
        Write-Warning "*** Unable to connect to $HvServer ***"
        Write-Output "Please re-enter the IP or FQDN."
        $HvServer = Read-Host -Prompt "Enter Horizon Connection Server IP or FQDN"
    }
}
until ($null -ne $objHvServer)
Write-Output "... Connected!"
Write-Output ""
#EndRegion Connect to Horizon Connection Server

#Region List desktop pools and select the pool to be updated
Write-Output "Getting a list of available desktop pools."
if (!($objHvServer.IsConnected)){
    $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential
}
$objPools = Get-HVPool -HvServer $objHvServer

if ($objPools.Count -gt 1){
    #Display menu
    $poolChoice = Show-UserChoices -InputObject $objPools -LabelProperty Base,DisplayName -Caption "The following desktop pools were found."
    $poolName = $objPools[$poolChoice].Base.Name
}
elseif ($objPools.Count -eq 1){
    #Only one pool exists
    Write-Output "Only found one desktop pool: $($objPools.Base.DisplayName)"
    $poolContinue = Read-Host "Do you want to continue and update this pool? (Y/N)"
    if ($poolContinue -eq "Y"){
        $poolName = $objPools.Base.Name
    }
    else {
        Write-Warning "The script will now exit."
        Disconnect-Sessions -vc $objvcenter -hv $objHvServer
    }
}
else {
    #No desktop pools found
    Write-Warning "No desktop pools were found on $HvServer."
    Write-Warning "The script will now exit."
    Disconnect-Sessions -vc $objvcenter -hv $objHvServer
}
#EndRegion List desktop pools and select the pool to be updated

#Region Disable the pool
Write-Output "Disabling the desktop pool."
if (!($objHvServer.IsConnected)){
    $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential
}
try {
    Set-HVPool -HvServer $objHvServer -PoolName $poolName -Disable -ErrorAction Stop
}
catch {
    Write-Warning "Error disabling the desktop pool $($objPools.Base.DisplayName)."
    Write-Warning "The script will now exit."
    Disconnect-Sessions -err $_ -vc $objvcenter -hv -$objHvServer
}
Write-Output "The pool has been disabled."
Write-Output ""
#EndRegion Disable the pool

#Region Remove the existing desktops from the pool and delete them from vCenter
Write-Output "Deleting the existing desktops."
if (!($objHvServer.IsConnected)){
    $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential
}
$hvMachines = Get-HVMachine -PoolName $poolName -HvServer $objHvServer
$list = $hvMachines.Base.Name
if (!($objvcenter.IsConnected)){
    $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential
}
#Save some vm details name for use later
$original = Get-VM -Name $list[0] -Server $objvcenter
$vmFolder = $original.Folder
$dsID = $original.DatastoreIdList
try {
    Remove-HVMachine -MachineNames $list -HVServer $objHvServer -DeleteFromDisk:$true -Confirm:$false -ErrorAction Stop
}
catch {
    Write-Warning "There was an error deleting the desktops."
    Write-Warning "Please confirm the desktops have been removed from the pool"
    Write-Warning "and manually delete any remaining desktops."
    Read-Host "Press Enter to continue."
}
Write-Output "The existing desktops have been deleted."
Write-Output ""
#EndRegion Remove the existing desktops from the pool and delete them from vCenter

#Region Clone new desktops
Write-Output "Ready to create the new desktops."
$numDesktops = Read-Host "How many desktops do you want to create? (1-20)"
if (!($objvcenter.IsConnected)){
    $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential
}
#Get a list of VM Templates to choose
$objTemplates = Get-Template -Server $objvcenter
if ($objTemplates.Count -gt 1){
    #Display Menu
    $tempChoice = Show-UserChoices -InputObject $objTemplates -LabelProperty Name -Caption "Select the VM template to use."
    $template = $objTemplates[$tempChoice]
}
else {
    #Only one template found
    Write-Output ""
    Write-Output "Only one template was found: $($objTemplates.Name)"
    Write-Output "This template will be used."
    $template = $objTemplates
}
#Get a list of available OS Customization Specifications from vCenter
$objCustSpec = Get-OSCustomizationSpec -Server $objvcenter
if ($objCustSpec.Count -gt 1){
    #Display Menu
    $specChoice = Show-UserChoices -InputObject $objCustSpec -LabelProperty Name -Caption "Which customization specification do you want to use?"
    $custSpec = $objCustSpec[$specChoice]
}
else {
    #Only one customization Specification found
    Write-Output ""
    Write-Output "Only one customization specification was found: $($objCustSpec.Name)"
    Write-Output "This customization specification will be used."
    $custSpec = $objCustSpec
}
#Ask to use the folder for the original desktops.
Write-Output ""
Write-Output "The original desktops were in the folder: $($vmFolder.Name)"
$folderAnswer = Read-Host "Do you want to use this same folder for the new desktops? (Y/N)"
if ($folderAnswer -eq 'N'){
    #Get a list of folders to choose
    $objFolders = Get-Folder -Type "VM" -Server $objvcenter
    $folderChoice = Show-UserChoices -InputObject $objFolders -LabelProperty Name -Caption "Choose the folder for the new desktops."
    $vmFolder = $objFolders[$folderChoice]
}
#Get a list of vSphere clusters to choose
$objClusters = Get-Cluster -Server $objvcenter
if ($objClusters.Count -gt 1){
    #Display Menu
    $clusterChoice = Show-UserChoices -InputObject $objClusters -LabelProperty Name -Caption "Choose a cluster for the new desktops."
    $cluster = $objClusters[$clusterChoice]
}
else {
    #only one cluster found
    Write-Output ""
    Write-Output "Only one cluster was found: $($objClusters.Name)"
    Write-Output "This cluster will be used."
    $cluster = $objClusters
}
#Get the original datastore
$ds = Get-Datastore -Id $dsID -Server $objvcenter
Write-Output ""
Write-Output "The original desktops were in the datastore: $($ds.Name)"
$dsAnswer = Read-Host "Do you want to use the same datastore for the new desktops? (Y/N)"
if ($dsAnswer -eq 'N'){
    $objDatastores = Get-Datastore -Server $objvcenter
    if ($objDatastores.Count -gt 1){
        #Display Menu
        $dsChoice = Show-UserChoices -InputObject $objDatastores -LabelProperty Name -Caption "Choose a datastore for the new desktops."
        $ds = $objDatastores[$dsChoice]
    }
    else {
        #Only one datastore found
        Write-Output ""
        Write-Output "Only one datastore was found: $($objDatastores.Name)"
        Write-Output "This datastore will be used."
    }
}
Write-Output ""
if (!($objvcenter.IsConnected)){
    $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential
}
foreach ($i in 1..$numDesktops){
    $newName = New-RandomVMName
    try {
        Write-Output "Creating desktop $newName ($i of $numDesktops)"
        New-VM -Name $newName -Template $template -Location $vmFolder -ResourcePool $cluster -Datastore $ds -OSCustomizationSpec $custSpec -Confirm:$false -ErrorAction Stop
        Start-VM -VM $newName
    }
    catch {
        Write-Warning "Error creating desktop $newName."
        Write-Warning "The desktop will not be created"
        Write-Error $_.Exception.Message
    }
}
Write-Output "Finished creating new desktops."
Write-Output ""
#EndRegion Clone new desktops

#Region Check for new desktops to finish SysPrep process
Write-Output "Waiting for the last desktop to finish the SysPrep process."
do {
    Write-Host -NoNewline "."
    Start-Sleep -Seconds 30
    $spVM = Get-VM -Name $newName
}
until ($spVM.Guest.HostName -like ($spVM.Name) + "*")
Write-Output "SysPrep process has finished."
Write-Output ""
#EndRegion Check for new desktops to finish SysPrep process

#Region Add new desktops to pool and wait for them to check in with the Connection Server
Write-Output "Adding new desktops to desktop pool."
if (!($objvcenter.IsConnected)){
    $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential
}
if (!($objHvServer.IsConnected)){
    $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential
}
$poolVMs = Get-VM -Location $vmFolder
try {
    Add-HVDesktop -PoolName $poolName -Machines $poolVMs -Vcenter $objvcenter -HvServer $objHvServer -Confirm:$false -ErrorAction Stop
}
catch {
    Write-Warning "Error adding desktops to pool."
    Write-Warning "You will have to add them manually."
    Read-Host "Press enter when you have added the desktops to the pool"
}
Write-Output "Desktops have been added to the pool."
Write-Output ""
Write-Output "Waiting for the desktops to check in with the Connection Server."
if (!($objHvServer.IsConnected)){
    $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential
}
do {
    Write-Host -NoNewline "."
    Start-Sleep -Seconds 30
    $csWait = Get-HVMachine -PoolName $poolName -HvServer $objHvServer | Where-Object {$_.Base.BasicState -ne "Available"}
}
while ($csWait.count -gt 0)
Write-Output "All desktops have checked in."
Write-Output ""
#EndRegion Add new desktops to pool and wait for them to check in with the Connection Server

#Region Set disk mode on new desktops
Write-Output "Shutting down desktops to update disk persistence."
if (!($objvcenter.IsConnected)){
    $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential
}
foreach ($vm in $poolVMs){
    $vmList = @()
    try {
        Stop-VMGuest -VM $vm -Confirm:$false -ErrorAction Stop
    }
    catch {
        Write-Warning "Error shutting down $($vm.Name)"
        $vmErr = $true
        $vmList += $vm.Name
    }
}
if ($vmErr){
    Write-Warning "There were errors shutting down some desktops."
    Write-Output "Please shutdown the following desktops:"
    $vmList
}
do {
    $vmOn = Get-VM -Location $vmFolder -Server $objvcenter | Where-Object {$_.PowerState -ne "PoweredOff"}
}
while ($vmOn.Count -gt 0)
Write-Output "Desktops have been shutdown."
Write-Output ""
Write-Output "Setting the desktop disk persistence."
if (!($objvcenter.IsConnected)){
    $objvcenter = Connect-VIServer -Server $vCenterServer -Credential $credential
}
foreach ($vm in $poolVMs){
    try {
        Get-HardDisk -VM $vm -Server $objvcenter | Set-HardDisk -Persistence IndependentNonPersistent -Confirm:$false -ErrorAction Stop
    }
    catch {
        Write-Warning "Error setting disk mode on $($vm.Name)"
        Write-Warning "You will have to set the disk mode manually."
    }
}
Write-Output "Disk persistence has been set."
Write-Output ""
Write-Output "Powering on the desktops."
foreach ($vm in $poolVMs){
    try {
        Start-VM -VM $vm -Server $objvcenter -Confirm:$false -ErrorAction Stop
    }
    catch {
        Write-Warning "Error starting $($vm.Name)"
    }
}
Write-Output "Desktops have been powered on."
Write-Output ""
#EndRegion Set disk mode on new desktops

#Region Enable the desktop pool
Write-Output "Enabling the desktop pool."
if (!($objHvServer.IsConnected)){
    $objHvServer = Connect-HVServer -Server $HvServer -Credential $credential
}
try {
    Set-HVPool -PoolName $poolName -Enable
}
catch {
    Write-Warning "Error enabling the desktop pool."
    Write-Warning "You will need to enable the pool manually."
}
#EndRegion Enable the desktop pool
Write-Output "**********"
Write-Output "Horizon View Desktop Pool Update Complete!"
Write-Output "**********"
#Region Cleanup the script environment
Disconnect-Sessions -vc $objvcenter -hv $objHvServer
#EndRegion Cleanup the script environment
}