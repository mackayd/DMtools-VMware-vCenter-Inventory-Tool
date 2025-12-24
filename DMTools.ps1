
<#
.SYNOPSIS
    DMTools.ps1 - VMware vSphere Inventory and Reporting Tool

.DESCRIPTION
    This script connects to a specified VMware vCenter Server and collects comprehensive inventory and configuration data from the entire vSphere environment. 
    It gathers detailed information about virtual machines, hosts, clusters, resource pools, datastores, networks, snapshots, orphaned files, hardware, and more.
    The collected data is exported into a multi-tabbed Excel workbook, with each worksheet representing a different aspect of the environment for easy review and analysis.

    Key features:
    - Checks for and installs required modules: VMware.PowerCLI, ImportExcel, psInlineProgress.
    - Prompts for vCenter address and credentials, and allows the user to select the Excel export location via a GUI dialog.
    - Connects to vCenter, collects data on VMs, CPU, memory, disks, partitions, SCSI, network, floppy, CD, snapshots, VMware Tools, resource pools, clusters, hosts, HBAs, NICs, vSwitches, port groups, distributed switches, VMkernel adapters, datastores, orphaned files, licenses, and recent health alarms.
    - Displays progress bars for each data collection phase.
    - Exports all data to a single Excel file with multiple worksheets, using the ImportExcel module.
    - Disconnects from vCenter upon completion.

.PARAMETER vCenter
    The FQDN or IP address of the vCenter Server to connect to.

.PARAMETER Credential
    Credentials with sufficient privileges to inventory the vSphere environment.

.OUTPUTS
    Excel Workbook (.xlsx) with multiple worksheets, each containing a different inventory or configuration report.

.NOTES
    - Requires PowerShell 5.1+ and Windows OS.
    - Requires network connectivity to the vCenter Server.
    - Script must be run with sufficient privileges to install modules and connect to vCenter.
    - May take several minutes to complete, depending on environment size.

.AUTHOR
    Drew Mackay
    https://github.com/mackayd

.VERSION
    1.0

.LICENSE
    MIT License

#>

# ---- SETUP AND MODULES ----

function Test-Module($ModuleName) {
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Module '$ModuleName' not found. Installing from PSGallery..." -ForegroundColor Yellow
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Module '$ModuleName' installed." -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install $ModuleName. Exiting." -ForegroundColor Red
            exit 1
        }
    }
}



function Get-ExcelFilePath-GUI {
    Add-Type -AssemblyName System.Windows.Forms
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Title = "Save DrewTools Excel Export"
    $saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    $saveFileDialog.FileName = "DMTools-Export-$((Get-Date).ToString('yyyyMMdd-HHmmss')).xlsx"
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $saveFileDialog.FileName
    }
    else {
        Write-Host "Export canceled by user." -ForegroundColor Yellow
        exit 1
    }
}

Test-Module -ModuleName VCF.PowerCLI
Test-Module -ModuleName ImportExcel
Test-Module -ModuleName psInlineProgress
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false | Out-Null

$vcenter = Read-Host "Enter vCenter Server address (FQDN or IP)"
$cred = Get-Credential -Message "Enter vCenter administrator credentials for $vcenter"
try {
    Connect-VIServer -Server $vcenter -Credential $cred -ErrorAction Stop | Out-Null
    Write-Host "Connected to $vcenter successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to $vcenter. Exiting." -ForegroundColor Red
    exit 1
}
$excelFile = Get-ExcelFilePath-GUI

# ---- ROOT DATA COLLECTION ----
$si = Get-View ServiceInstance
$about = $si.Content.About
$vms = Get-VM | Where-Object { $_.Name -notlike "*vCLS-*" }
$ESXhosts = Get-VMHost
$clusters = Get-Cluster
$rpools = Get-ResourcePool
$datastores = Get-Datastore
$VCC = $global:DefaultVIServer[0]

function Get-VMContext {
    param($vm)
    $cluster = $null
    try { $cluster = $vm | Get-Cluster } catch {}
    $datacenter = $null
    try { $datacenter = $vm | Get-Datacenter } catch {}
    $resourcePool = $null
    try { $resourcePool = $vm | Get-ResourcePool } catch {}
    $vapp = $null
    try { $vapp = $vm.VApp } catch {}
    $folder = $null
    try { $folder = $vm.Folder.Name } catch {}
    return @{
        Cluster      = $cluster.Name
        Datacenter   = $datacenter.Name
        ResourcePool = $resourcePool.Name
        vApp         = $vapp.Name
        Folder       = $folder
    }
}

# ---- FUNCTIONS FOR EACH TAB ----

function Get-vInfo {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vInfo $i of $total VMs" -PercentComplete ([int](($i / $total) * 100)) -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        [PSCustomObject]@{
            Name                                     = $vm.Name
            Annotation                               = $ed.Config.Annotation
            Powerstate                               = $vm.PowerState
            Template                                 = $ed.Config.Template
            vApp                                     = $ctx.vApp
            ResourcePool                             = $ctx.ResourcePool
            Folder                                   = $ctx.Folder
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $vm.VMHost.Name
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $vm.Guest.OSFullName
            IP                                       = $ed.Summary.Guest.IpAddress
            MacAddress                               = ($vm | Get-NetworkAdapter | Select-Object -First 1).MacAddress
            VMwareTools                              = $ed.Guest.ToolsStatus
            VMwareToolsVersion                       = $ed.Guest.ToolsVersion
            VMwareToolsVersionStatus                 = $ed.Guest.ToolsVersionStatus2
            VMwareToolsRunningStatus                 = $ed.Guest.ToolsRunningStatus
            ChangeVersion                            = $ed.Config.ChangeVersion
            ConfigStatus                             = $ed.OverallStatus
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            "VI SDK Server"                          = $about.FullName
            "VI SDK UUID"                            = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vInfo Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vCPU {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count    
    $i = 0

    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vCPU $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        [PSCustomObject]@{
            Name                                     = $vm.Name
            Annotation                               = $ed.Config.Annotation
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $vm.VMHost.Name
            Folder                                   = $ctx.Folder
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $vm.Guest.OSFullName
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            vCPUs                                    = $ed.Config.Hardware.NumCPU
            Sockets                                  = $ed.Config.Hardware.NumCPU / $ed.Config.Hardware.NumCoresPerSocket
            CoresPerSocket                           = $ed.Config.Hardware.NumCoresPerSocket
            MaxCPU_MHz                               = $ed.Runtime.MaxCpuUsage
            CPU_Usage_MHz                            = $ed.Summary.QuickStats.OverallCpuUsage
            SharesLevel                              = $ed.Config.CpuAllocation.Shares.Level
            Shares                                   = $ed.Config.CpuAllocation.Shares.Shares
            CPU_Reservation                          = $ed.Config.CpuAllocation.Reservation
            CPULimit                                 = if ($ed.Config.CpuAllocation.Limit -eq -1) { 0 } else { $ed.Config.CpuAllocation.Limit }
            EntitlementMHz                           = $ed.Summary.QuickStats.StaticCpuEntitlement
            DRSEntitlementMHz                        = $ed.Summary.QuickStats.DistributedCpuEntitlement
            CPUHotAdd                                = $ed.Config.CpuHotAddEnabled
            CPUHotRemove                             = $ed.Config.CpuHotRemoveEnabled
            "VI SDK Server"                          = $about.FullName
            "VI SDK UUID"                            = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vCPU Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null

}

function Get-vMemory {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0

    foreach ($vm in $vms) {
        $i++

        Write-InlineProgress -Activity "Collecting vMemory $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        [PSCustomObject]@{
            Name                                     = $vm.Name
            Annotation                               = $ed.Config.Annotation
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $vm.VMHost.Name
            Folder                                   = $ctx.Folder
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $vm.Guest.OSFullName
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            MemoryMB                                 = $ed.Config.Hardware.MemoryMB
            Overhead                                 = $ed.Config.MemoryAllocation.OverheadLimit
            MemLimitMB                               = if ($ed.Config.MemoryAllocation.Limit -eq -1) { 0 } else { $ed.Config.MemoryAllocation.Limit }
            MemReservation                           = $ed.Config.MemoryAllocation.Reservation
            MemSharesLevel                           = $ed.Config.MemoryAllocation.Shares.Level
            MemShares                                = $ed.Config.MemoryAllocation.Shares.Shares
            MemConsumedMB                            = $ed.Summary.QuickStats.HostMemoryUsage
            OverheadMB                               = $ed.Summary.QuickStats.ConsumedOverheadMemory
            PrivateMB                                = $ed.Summary.QuickStats.PrivateMemory
            SharedMB                                 = $ed.Summary.QuickStats.SharedMemory
            SwappedMB                                = $ed.Summary.QuickStats.SwappedMemory
            BalloonedMB                              = $ed.Summary.QuickStats.BalloonedMemory
            CompressedMB                             = [math]::Round($ed.Summary.QuickStats.CompressedMemory / 1024, 0)
            ActiveMB                                 = $ed.Summary.QuickStats.GuestMemoryUsage
            EntitlementMB                            = $ed.Summary.QuickStats.StaticMemoryEntitlement
            DRSEntitlementMB                         = $ed.Summary.QuickStats.DistributedMemoryEntitlement
            MemHotAdd                                = $ed.Config.MemoryHotAddEnabled
            "VI SDK Server"                          = $about.FullName
            "VI SDK UUID"                            = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vMemory Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vDisk {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count      
    $i = 0

    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vDisk $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null


        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        foreach ($disk in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] }) {
            $ctl = $ed.Config.Hardware.Device | Where-Object { $_.Key -eq $disk.ControllerKey }
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $vm.VMHost.Name
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                Disk                                     = $disk.DeviceInfo.Label
                CapacityMB                               = [int]($disk.CapacityInKB / 1024)
                Thin                                     = $disk.Backing.ThinProvisioned
                EagerZero                                = $disk.Backing.EagerlyScrub
                Mode                                     = $disk.Backing.DiskMode
                Controller                               = $ctl.GetType().Name
                ControllerBus                            = $ctl.BusNumber
                Unit                                     = $disk.UnitNumber
                Path                                     = $disk.Backing.FileName
                Raw                                      = ($disk.Backing -is [VMware.Vim.VirtualDiskRawDiskMappingVer1BackingInfo])
                LunUuid                                  = $disk.Backing.LunUuid
                RDMMode                                  = $disk.Backing.CompatibilityMode
                "VI SDK Server"                          = $about.FullName
                "VI SDK UUID"                            = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vDisk Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vPartition {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0

    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vPartition $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        if ($ed.Guest.Disk) {
            foreach ($gdisk in $ed.Guest.Disk) {
                [PSCustomObject]@{
                    Name                                     = $vm.Name
                    Annotation                               = $ed.Config.Annotation
                    Datacenter                               = $ctx.Datacenter
                    Cluster                                  = $ctx.Cluster
                    Host                                     = $vm.VMHost.Name
                    Folder                                   = $ctx.Folder
                    "OS according to the configuration file" = $ed.Config.GuestFullName
                    "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                    VM_ID                                    = $ed.MoRef.Value
                    UUID                                     = $ed.Config.Uuid
                    Disk                                     = $gdisk.DiskPath
                    CapacityMB                               = [math]::Round($gdisk.Capacity / 1MB, 0)
                    FreeMB                                   = [math]::Round($gdisk.FreeSpace / 1MB, 0)
                    FreePct                                  = if ($gdisk.Capacity -gt 0) { [math]::Round(($gdisk.FreeSpace / $gdisk.Capacity * 100), 0) } else { 0 }
                    Powerstate                               = $vm.PowerState
                    "VI SDK Server"                          = $about.FullName
                    "VI SDK UUID"                            = $about.InstanceUuid
                }
            }
        }
    }
    Write-InlineProgress -Activity 'vPartition Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vSCSI {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0

    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vSCSI $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
       if ($ed.guest.guestfamily -like "*windowsGuest") {
            if($($VM | Get-AdvancedSetting -name "disk.EnableUUID") -like "*TRUE*"){
                $diskData = $true
            }else{
                $diskData = $false
            }
       }
       Else{$diskData=$true}
       If($diskData){
            foreach ($gDisk in Get-VMGuestDisk -VM $vm) {

                # Map guest-visible disk/partition → HardDisk → SCSI controller
                $vDisk = Get-HardDisk -VMGuestDisk $gDisk -ErrorAction SilentlyContinue
                $ctlr = Get-ScsiController -VM $vm |
                Where-Object { $_.ExtensionData.Key -eq $vDisk.ExtensionData.ControllerKey }

                # Compose a clean SCSI “bus:unit” string
                $SCSIbus = $ctlr.ExtensionData.BusNumber
                $unit = $vDisk.ExtensionData.UnitNumber
                $scsi = "$SCSIbus : $unit"
                [PSCustomObject]@{
       
                    Name                                     = $vm.Name
                    Annotation                               = $ed.Config.Annotation
                    Datacenter                               = $ctx.Datacenter
                    Cluster                                  = $ctx.Cluster
                    Host                                     = $vm.VMHost.Name
                    Folder                                   = $ctx.Folder
                    "OS according to the configuration file" = $ed.Config.GuestFullName
                    "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                    VM_ID                                    = $ed.MoRef.Value
                    UUID                                     = $ed.Config.Uuid    
                    GuestDiskPath                            = $gDisk.DiskPath    
                    CapacityGB                               = [math]::Round($gDisk.Capacity / 1GB, 2)
                    FreeGB                                   = [math]::Round($gDisk.FreeSpace / 1GB, 2)
                    SCSIname                                 = $ctlr.Name
                    SCSIControler                            = $SCSIbus
                    UnitNumber                               = $unit
                    SCSI_Path                                = $scsi        
                    VMDK_Path                                = $vDisk.FileName             
                    "VI SDK Server"                          = $about.FullName
                    "VI SDK UUID"                            = $about.InstanceUuid
                }
            }
        }
        else{
             $scsi = "$SCSIbus : $unit"
                [PSCustomObject]@{
       
                    Name                                     = $vm.Name
                    Annotation                               = "WINDOWS disk data requires Disk.enableUUID advanced setting"
                    Datacenter                               = ""
                    Cluster                                  = ""
                    Host                                     = ""
                    Folder                                   = ""
                    "OS according to the configuration file" = ""
                    "OS according to the VMware Tools"       = ""
                    VM_ID                                    = ""
                    UUID                                     = ""
                    GuestDiskPath                            = ""
                    CapacityGB                               = ""
                    FreeGB                                   = ""
                    SCSIname                                 = ""
                    SCSIControler                            = ""
                    UnitNumber                               = ""
                    SCSI_Path                                = ""
                    VMDK_Path                                = ""
                    "VI SDK Server"                          = ""
                    "VI SDK UUID"                            = ""
                }
        }
        
    }
    Write-InlineProgress -Activity 'vSCSI Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}


function Get-vNetwork {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count 
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vNetwork $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        foreach ($nic in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualEthernetCard] }) {
            $net = $nic.Backing
            $networkName = $nic.DeviceInfo.Label
            $pgName = if ($net -is [VMware.Vim.VirtualEthernetCardNetworkBackingInfo]) { $net.DeviceName } else { $null }
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $vm.VMHost.Name
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                Adapter                                  = $networkName
                Network                                  = $pgName
                MAC                                      = $nic.MacAddress
                Type                                     = $nic.GetType().Name
                Connected                                = $nic.Connectable.Connected
                StartConnected                           = $nic.Connectable.StartConnected
                "VI SDK Server"                          = $about.FullName
                "VI SDK UUID"                            = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vNetwork Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null

}

function Get-vFloppy {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vFloppy $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null


        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        foreach ($fl in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualFloppy] }) {
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $vm.VMHost.Name
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                Label                                    = $fl.DeviceInfo.Label
                Connected                                = $fl.Connectable.Connected
                StartConnected                           = $fl.Connectable.StartConnected
                Type                                     = $fl.Backing.GetType().Name
                Backing                                  = $fl.Backing.Filename
                "VI SDK Server"                          = $about.FullName
                "VI SDK UUID"                            = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vFloppy Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vCD {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vCD $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null


        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        foreach ($cd in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualCdrom] }) {
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $vm.VMHost.Name
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                Label                                    = $cd.DeviceInfo.Label
                Connected                                = $cd.Connectable.Connected
                StartConnected                           = $cd.Connectable.StartConnected
                Type                                     = $cd.Backing.GetType().Name
                Backing                                  = $cd.Backing.Filename
                "VI SDK Server"                          = $about.FullName
                "VI SDK UUID"                            = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vCD Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vSnapshot {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vSnapshot $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        if ($ed.Snapshot) {
            foreach ($snap in Get-Snapshot -vm $vm ) {
                
                [PSCustomObject]@{
                    Name                                     = $vm.Name
                    Annotation                               = $ed.Config.Annotation
                    Datacenter                               = $ctx.Datacenter
                    Cluster                                  = $ctx.Cluster
                    Host                                     = $vm.VMHost.Name
                    Folder                                   = $ctx.Folder
                    "OS according to the configuration file" = $ed.Config.GuestFullName
                    "OS according to the VMware Tools"       = $vm.Guest.OSFullName
                    VM_ID                                    = $ed.MoRef.Value
                    UUID                                     = $ed.Config.Uuid
                    SnapshotName                             = $snap.Name
                    Description                              = $snap.Description
                    Created                                  = $snap.Create
                    isCurrent                                = $snap.isCurrent
                    Quiesced                                 = $snap.Quiesced
                    SizeGB                                   = $snap.SizeGB
                    "VI SDK Server"                          = $about.FullName
                    "VI SDK UUID"                            = $about.InstanceUuid
                }
            }
        }
    }
    Write-InlineProgress -Activity 'vSnapshot Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vTools {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vTools $i of $total VMs" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
       

        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        [PSCustomObject]@{
            Name                                     = $vm.Name
            Annotation                               = $ed.Config.Annotation
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $vm.VMHost.Name
            Folder                                   = $ctx.Folder
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $vm.Guest.OSFullName
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            VMVersion                                = $ed.Config.Version -replace "vmx-", ""
            Tools                                    = $ed.Guest.ToolsStatus
            ToolsVersion                             = $ed.Guest.ToolsVersion
            ToolsStatus2                             = $ed.Guest.ToolsVersionStatus2
            ToolsInstallType                         = $ed.Guest.ToolsInstallType
            SyncTime                                 = $ed.Config.Tools.SyncTime
            AppStatus                                = $ed.Guest.AppState
            AppHeartbeat                             = $ed.Guest.AppHeartbeatStatus
            KernelCrash                              = $ed.Guest.GuestKernelCrashed
            OpsReady                                 = $ed.Guest.GuestOperationsReady
            InteractiveReady                         = $ed.Guest.InteractiveGuestOperationsReady
            StateChangeSupported                     = $ed.Guest.GuestStateChangeSupported
            ToolsUpgradePolicy                       = $ed.Config.Tools.ToolsUpgradePolicy
            "VI SDK Server"                          = $about.FullName
            "VI SDK UUID"                            = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vTools Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vRP {
    param($rpools, $about)
    $total = $rpools.Count
    $i = 0
    foreach ($rp in $rpools) {
        $i++
        Write-InlineProgress -Activity "Collecting pools $i of $total resource pools" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        
        $ed = $rp.ExtensionData
        $qs = $ed.Runtime.QuickStats
        $vmsInPool = $rp | Get-VM

        [PSCustomObject]@{
            Name                              = $rp.Name
            Status                            = $ed.OverallStatus
            NumVMs                            = $vmsInPool.Count
            vCPUs                             = ($vmsInPool | Measure-Object -Property NumCPU -Sum).Sum
            MemConfigured                     = ($vmsInPool | Measure-Object -Property MemoryMB -Sum).Sum
            CPU_Limit                         = $ed.Config.CpuAllocation.Limit
            CPU_Reservation                   = $ed.Config.CpuAllocation.Reservation
            CPU_SharesLevel                   = $ed.Config.CpuAllocation.Shares.Level
            CPU_Shares                        = $ed.Config.CpuAllocation.Shares.Shares
            CPU_Expandable                    = $ed.Config.CpuAllocation.ExpandableReservation
            Mem_Limit                         = $ed.Config.MemoryAllocation.Limit
            Mem_Reservation                   = $ed.Config.MemoryAllocation.Reservation
            Mem_SharesLevel                   = $ed.Config.MemoryAllocation.Shares.Level
            Mem_Shares                        = $ed.Config.MemoryAllocation.Shares.Shares
            Mem_Expandable                    = $ed.Config.MemoryAllocation.ExpandableReservation
            "Mem maxUsage"                    = $ed.Runtime.MaxUsage
            "Mem overallUsage"                = $ed.Runtime.OverallUsage
            "Mem reservationUsed"             = $ed.Runtime.ReservationUsed
            "Mem reservationUsedForVm"        = $ed.Runtime.ReservationUsedForVm
            "Mem unreservedForPool"           = $ed.Runtime.UnreservedForPool
            "Mem unreservedForVm"             = $ed.Runtime.UnreservedForVm
            "QS overallCpuDemand"             = $qs.OverallCpuDemand
            "QS overallCpuUsage"              = $qs.OverallCpuUsage
            "QS staticCpuEntitlement"         = $qs.StaticCpuEntitlement
            "QS distributedCpuEntitlement"    = $qs.DistributedCpuEntitlement
            "QS balloonedMemory"              = $qs.BalloonedMemory
            "QS compressedMemory"             = $qs.CompressedMemory
            "QS consumedOverheadMemory"       = $qs.ConsumedOverheadMemory
            "QS distributedMemoryEntitlement" = $qs.DistributedMemoryEntitlement
            "QS guestMemoryUsage"             = $qs.GuestMemoryUsage
            "QS hostMemoryUsage"              = $qs.HostMemoryUsage
            "QS overheadMemory"               = $qs.OverheadMemory
            "QS privateMemory"                = $qs.PrivateMemory
            "QS sharedMemory"                 = $qs.SharedMemory
            "QS staticMemoryEntitlement"      = $qs.StaticMemoryEntitlement
            "QS swappedMemory"                = $qs.SwappedMemory
            "VI SDK Server"                   = $about.FullName
            "VI SDK UUID"                     = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vResourcePools Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}


function Get-vCluster {
    param($clusters, $about)
    $total = $clusters.Count
    $i = 0
    foreach ($cl in $clusters) {
        $i++
        Write-InlineProgress -Activity "Collecting vCluster $i of $total clusters" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null


        $ed = $cl.ExtensionData
        $das = $ed.Configuration.DasConfig
        $dasDef = $das.DefaultVmSettings
        $drs = $ed.Configuration.DrsConfig
        $dpm = $ed.Configuration.DpmConfig
        [PSCustomObject]@{
            Name                            = $cl.Name
            "Config status"                 = $ed.ConfigurationEx.Status
            OverallStatus                   = $ed.OverallStatus
            NumHosts                        = $ed.Summary.NumHosts
            numEffectiveHosts               = $ed.Summary.NumEffectiveHosts
            TotalCpu                        = $ed.Summary.TotalCpu
            NumCpuCores                     = $ed.Summary.NumCpuCores
            NumCpuThreads                   = $ed.Summary.NumCpuThreads
            "Effective Cpu"                 = $ed.Summary.EffectiveCpu
            TotalMemory                     = [math]::Round($ed.Summary.TotalMemory / 1MB, 0)
            "Effective Memory"              = $ed.Summary.EffectiveMemory
            "Num VMotions"                  = $ed.Summary.NumVmotions
            "HA enabled"                    = $das.Enabled
            "Failover Level"                = $das.FailoverLevel
            AdmissionControlEnabled         = $das.AdmissionControlEnabled
            "Host monitoring"               = $das.HostMonitoring
            "HB Datastore Candidate Policy" = $das.HBDatastoreCandidatePolicy
            "Isolation Response"            = $dasDef.IsolationResponse
            "Restart Priority"              = $dasDef.RestartPriority
            "Max Failures"                  = $das.MaxFailures
            "Max Failure Window"            = $das.MaxFailureWindow
            "Failure Interval"              = $das.FailureInterval
            "Min Up Time"                   = $das.MinUpTime
            "VM Monitoring"                 = $das.VmMonitoring
            "DRS enabled"                   = $drs.Enabled
            "DRS default VM behavior"       = $drs.DefaultVmBehavior
            "DRS vmotion rate"              = $drs.VmotionRate
            "DPM enabled"                   = $dpm.Enabled
            "DPM default behavior"          = $dpm.DefaultBehavior
            "DPM Host Power Action Rate"    = $dpm.HostPowerActionRate
            "VI SDK Server"                 = $about.FullName
            "VI SDK UUID"                   = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vCluster Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vHost {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vHost $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
    

        $ed = $ESXhost.ExtensionData
        $prod = $ed.Summary.Config.Product
        $sys = $ed.Hardware.SystemInfo
        $bios = $ed.Hardware.BiosInfo
        [PSCustomObject]@{
            Name               = $ESXhost.Name
            Datacenter         = ($ESXhost | Get-Datacenter).Name
            Cluster            = ($ESXhost | Get-Cluster).Name
            Status             = $ed.OverallStatus
            CPUModel           = $ed.Summary.Hardware.CpuModel
            SpeedMHz           = $ed.Summary.Hardware.CpuMhz
            HT_Available       = $ed.Config.HyperThread.Available
            HT_Active          = $ed.Config.HyperThread.Active
            CPUPackages        = $ed.Summary.Hardware.NumCpuPkgs
            CoresPerCPU        = $ed.Summary.Hardware.NumCpuCores / $ed.Summary.Hardware.NumCpuPkgs
            TotalCores         = $ed.Summary.Hardware.NumCpuCores
            CPUUsagePct        = [math]::Round(($ed.Summary.QuickStats.OverallCpuUsage / ($ed.Summary.Hardware.CpuMhz * $ed.Summary.Hardware.NumCpuCores) * 100), 1)
            MemoryGB           = [math]::Round($ed.Summary.Hardware.MemorySize / 1GB, 0)
            MemoryUsagePct     = [math]::Round(($ed.Summary.QuickStats.OverallMemoryUsage / ($ed.Summary.Hardware.MemorySize / 1MB) * 100), 1)
            NumNICs            = $ed.Config.Network.Pnic.Count
            NumHBAs            = $ed.Config.StorageDevice.HostBusAdapter.Count
            NumVMs             = ($ESXhost | Get-VM).Count
            VMsPerCore         = [math]::Round((($ESXhost | Get-VM).Count / $ed.Summary.Hardware.NumCpuCores), 2)
            ESXiVersion        = $prod.FullName
            DisabledTLSversion = (Get-AdvancedSetting -Entity $ESXHost -Name 'UserVars.ESXiVPsDisabledProtocols' -ErrorAction SilentlyContinue).Value
            BootTime           = $ed.Summary.Runtime.BootTime
            Vendor             = $sys.Vendor
            Model              = $sys.Model
            Serial             = $sys.SerialNumber
            BIOSVendor         = $bios.Vendor
            BIOSVersion        = $bios.BiosVersion
            BIOSDate           = $bios.ReleaseDate
            ObjectID           = $ed.MoRef.Value
            "VI SDK Server"    = $about.FullName
            "VI SDK UUID"      = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vHost Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}


function Get-vTLShosts {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vTLShosts $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null



        if ($ESXhost.Version -like "8.0.*") {

            $esxcli = Get-EsxCli -VMHost $ESXhost -V2

    
            $srvProfile = $esxcli.system.tls.server.get.Invoke().Profile

  
            $cliProfile = $esxcli.system.tls.client.get.Invoke().Profile
            $explicitProtocols = $null
            if ($srvProfile -eq 'MANUAL') {
                $ESXargs = $esxcli.system.tls.server.get.CreateArgs()
                $ESXargs.protocolversions = $true   # expose the individual versions
                $explicitProtocols = $esxcli.system.tls.server.get.Invoke($ESXargs).ProtocolVersions -join ','
            }
        }
        else {
            $srvProfile = "Only available from ESXI 8.0U3"
            $cliProfile = "Only available from ESXI 8.0U3"

        }
        $disabled = (Get-AdvancedSetting -Entity $ESXhost -Name 'UserVars.ESXiVPsDisabledProtocols' `
                -ErrorAction SilentlyContinue).Value


        [pscustomobject]@{
            Name              = $ESXhost.Name
            ServerTLSProfile  = $srvProfile
            ClientTLSProfile  = $cliProfile
            DisabledProtocols = $disabled
            ExplicitProtocols = $explicitProtocols
        }
    }
    Write-InlineProgress -Activity 'vTLShosts Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vHBA {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vHBA $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null


        $ed = $ESXhost.ExtensionData
        foreach ($hba in $ed.Config.StorageDevice.HostBusAdapter) {
            [PSCustomObject]@{
                Host              = $ESXhost.Name
                HBADevice         = $hba.Device
                Model             = $hba.Model
                Type              = $hba.GetType().Name
                Status            = $hba.Status
                NodeWorldWideName = $hba.NodeWorldWideName
                PortWorldWideName = $hba.PortWorldWideName
                Driver            = $hba.Driver
                PCI               = $hba.Pci
                "VI SDK Server"   = $about.FullName
                "VI SDK UUID"     = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vHBA Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vNIC {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vNIC $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
      
        $ed = $ESXhost.ExtensionData
        foreach ($pnic in $ed.Config.Network.Pnic) {
            [PSCustomObject]@{
                Host            = $ESXhost.Name
                PNICDevice      = $pnic.Device
                MAC             = $pnic.Mac
                LinkSpeed       = $pnic.LinkSpeed.SpeedMb
                Driver          = $pnic.Driver
                PCI             = $pnic.Pci
                WakeOnLAN       = $pnic.WakeOnLanSupported
                "VI SDK Server" = $about.FullName
                "VI SDK UUID"   = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vNIC Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vSwitch {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vSwitch $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

        foreach ($vs in Get-VirtualSwitch -VMHost $ESXhost) {
            $vsExt = $vs.ExtensionData
            [PSCustomObject]@{
                Host              = $ESXhost.Name
                vSwitch           = $vs.Name
                NumPorts          = $vsExt.NumPorts
                NumPortsAvailable = $vsExt.NumPortsAvailable
                MTU               = $vs.Mtu
                Nic               = ($vs.Nic -join ',')
                ActiveNic         = $vs.ExtensionData.Spec.Policy.NicTeaming.NicOrder.ActiveNic -join ','
                StandbyNic        = $vs.ExtensionData.Spec.Policy.NicTeaming.NicOrder.StandbyNic -join ','
                AllowPromiscuous  = $vs.ExtensionData.Spec.Policy.Security.AllowPromiscuous
                ForgedTransmits   = $vs.ExtensionData.Spec.Policy.Security.ForgedTransmits
                MacChanges        = $vs.ExtensionData.Spec.Policy.Security.MacChanges
                CheckBeacon       = $vs.ExtensionData.Spec.Policy.NicTeaming.FailureCriteria.CheckBeacon
                "VI SDK Server"   = $about.FullName
                "VI SDK UUID"     = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vSwitch Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vPort {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vPort $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        
        foreach ($pg in Get-VirtualPortGroup -VMHost $ESXhost) {
            $pgExt = $pg.ExtensionData
            [PSCustomObject]@{
                Host             = $ESXhost.Name
                PortGroup        = $pg.Name
                vSwitch          = $pg.VirtualSwitchName
                VLANId           = $pg.VlanId
                NumPorts         = $pgExt.NumPorts
                ActivePorts      = $pgExt.NumPortsActive
                AllowPromiscuous = $pgExt.Spec.Policy.Security.AllowPromiscuous.Value
                ForgedTransmits  = $pgExt.Spec.Policy.Security.ForgedTransmits.Value
                MacChanges       = $pgExt.Spec.Policy.Security.MacChanges.Value
                "VI SDK Server"  = $about.FullName
                "VI SDK UUID"    = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vPort Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vdvSwitch {
    param($about)
    $allDVS = Get-vdSwitch  
    $total = $allDVS.Count
    $i = 0    
    foreach ($dvs in $allDVS) {
        $i++
        Write-InlineProgress -Activity "Collecting dvSwitch $i of $total dvSwitches" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        $dvsExt = $dvs.ExtensionData
        [PSCustomObject]@{
            Name             = $dvs.Name
            Version          = $dvs.Version
            NumPorts         = $dvs.NumPorts
            NumUplinks       = $dvs.NumUplinks
            MTU              = $dvs.Mtu
            Description      = $dvs.Description
            Uplinks          = ($dvs.UplinkPortPolicy.UplinkPortName -join ',')
            AllowPromiscuous = $dvsExt.Config.DefaultPortConfig.SecurityPolicy.AllowPromiscuous.Value
            ForgedTransmits  = $dvsExt.Config.DefaultPortConfig.SecurityPolicy.ForgedTransmits.Value
            MacChanges       = $dvsExt.Config.DefaultPortConfig.SecurityPolicy.MacChanges.Value
            Datacenter       = $dvs.VMwareDatacenter.Name
            "VI SDK Server"  = $about.FullName
            "VI SDK UUID"    = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'dvSwitches Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vdvPort {
    param($about)
    $allDVport = Get-VDPortgroup 
    $total = $allDVport.Count
    $i = 0    
    foreach ($dvpg in $allDVport) {
        $i++
        Write-InlineProgress -Activity "Collecting dvPorts $i of $total dvPorts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        $dvpgExt = $dvpg.ExtensionData
        [PSCustomObject]@{
            Name            = $dvpg.Name
            VDSwitch        = $dvpg.VDSwitch.Name
            VLANId          = $dvpg.VlanConfiguration.VlanId
            NumPorts        = $dvpg.NumPorts
            ActivePorts     = $dvpgExt.RuntimeInfo.PortCount
            Type            = $dvpg.PortBinding
            Description     = $dvpg.Description
            "VI SDK Server" = $about.FullName
            "VI SDK UUID"   = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'dvPorts Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null

}


function Get-vSC_VMK {
    param($ESXhosts, $about)
    $total = $ESXhosts.Count
    $i = 0
    foreach ($ESXhost in $ESXhosts) {
        $i++
        Write-InlineProgress -Activity "Collecting vSC_VMK $i of $total hosts" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

     
        foreach ($vmk in Get-VMHostNetworkAdapter -VMKernel -VMHost $ESXhost) {
            [PSCustomObject]@{
                Host                  = $ESXhost.Name
                VMKernelAdapter       = $vmk.Name
                IPAddress             = $vmk.IP
                SubnetMask            = $vmk.SubnetMask
                MAC                   = $vmk.Mac
                MTU                   = $vmk.Mtu
                vSwitch               = $vmk.VirtualSwitch
                PortGroup             = $vmk.PortGroupName
                DHCPEnabled           = $vmk.DhcpEnabled
                IPv6                  = $vmk.IPv6[0].address
                Management            = $vmk.ManagementTrafficEnabled
                FaultTolerance        = $vmk.FaultToleranceLoggingEnabled
                vMotion               = $vmk.VMotionEnabled
                vSAN                  = $vmk.VsanTrafficEnabled
                Provisioning          = $vmk.ProvisioningTrafficEnabled
                VSphereReplication    = $vmk.VSphereReplicationTrafficEnabled
                VSphereReplicationNFC = $vmk.VSphereReplicationNfcEnabled
                "VI SDK Server"       = $about.FullName
                "VI SDK UUID"         = $about.InstanceUuid
            }
        }
    }
    Write-InlineProgress -Activity 'vSC_VMK Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vDatastore {
    param($datastores, $about)
    $total = $datastores.Count; $i = 0
    foreach ($ds in $datastores) {
        $i++
        Write-InlineProgress -Activity "Collecting vDatastore $i of $total datastores" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null

    
        $dsExt = $ds.ExtensionData
        $ESXhostsOnDS = ($ds | Get-VMHost) | Select-Object -ExpandProperty Name
        [PSCustomObject]@{
            Name            = $ds.Name
            Type            = $ds.Type
            FreeSpaceGB     = [math]::Round($ds.FreeSpaceGB, 2)
            CapacityGB      = [math]::Round($ds.CapacityGB, 2)
            Hosts           = ($ESXhostsOnDS -join ', ')
            NumVMs          = ($ds | Get-VM).Count
            Cluster         = ($ds | Get-VMHost | Get-Cluster).Name
            Datacenter      = ($ds | Get-VMHost | Get-Cluster | Get-Datacenter).Name 
            FileSystem      = $dsExt.Summary.Type
            URL             = $dsExt.Info.Url
            UUID            = $dsExt.Info.MembershipUuid
            "VI SDK Server" = $about.FullName
            "VI SDK UUID"   = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vDatastore Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}


function Get-VmwOrphan {
    param($Datastores, $about)
    $flags = New-Object VMware.Vim.FileQueryFlags
    $flags.FileSize = $true
    $flags.Modification = $true
    $generic = New-Object VMware.Vim.FileQuery
    $searchSpec = New-Object VMware.Vim.HostDatastoreBrowserSearchSpec
    $searchSpec.Query = $generic
    $searchSpec.Details = $flags
    $searchSpec.MatchPattern = @('*.vmdk', '*.vmx')
    
    $total = $datastores.Count
    $i = 0
    foreach ($ds in $Datastores) {
        $i++
        Write-InlineProgress -Activity "Collecting vZombie $i of $total datastores" `
            -PercentComplete ([int](($i / $total) * 100)) `
            -ProgressCharacter ([char]9632) `
            -ProgressFillCharacter ([char]9632) `
            -ProgressFill ([char]183) `
            -BarBracketStart $null `
            -BarBracketEnd $null
        if (($ds.Type -in 'VMFS', 'vsan') -and $ds.ExtensionData.Summary.MultipleHostAccess) {

               

            $browser = Get-View $ds.ExtensionData.Browser
            $rootPath = '[' + $ds.Name + ']'

            # ----- Build hash-set of all files referenced by VMs/Templates -----
            $vmFiles = @{}
            Get-VM -Datastore $ds -ErrorAction SilentlyContinue | ForEach-Object {
                $view = Get-View $_.Id
                $view.LayoutEx.File | ForEach-Object {
                    $_.Name.ToLower() | ForEach-Object { $vmFiles[$_] = $true }
                }
            }
            Get-Template -Datastore $ds -ErrorAction SilentlyContinue | ForEach-Object {
                $view = Get-View $_.Id
                $view.LayoutEx.File | ForEach-Object {
                    $_.Name.ToLower() | ForEach-Object { $vmFiles[$_] = $true }
                }
            }


            # ----- Enumerate every folder & file on the datastore -----
            $result = $browser.SearchDatastoreSubFolders($rootPath, $searchSpec)

            foreach ($folder in $result) {
                foreach ($f in $folder.File) {
                    $full = "$($folder.FolderPath)/$($f.Path)"
                    if (-not $vmFiles.ContainsKey($full.ToLower())) {
                        # ------ OUTPUT orphaned file object ------
                        [pscustomobject]@{
                            Datastore       = $ds.Name
                            FilePath        = $full
                            FileSizeGB      = [math]::Round($f.FileSize / 1GB, 2)
                            Modified        = $f.Modification
                            "VI SDK Server" = $about.FullName
                            "VI SDK UUID"   = $about.InstanceUuid

                        }
                    }
                }
            }
        }
    }
    Write-InlineProgress -Activity 'vZombie Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}


function Get-vLicense {
    param($about)
     
    foreach ($lic in $(Get-View LicenseManager -Server $vcc).Licenses) {
        [PSCustomObject]@{
            Name            = $lic.Name
            LicenseKey      = $lic.LicenseKey
            EditionKey      = $lic.EditionKey
            CostUnit        = $lic.costunit
            Total           = $lic.Total
            Used            = $lic.Used
            ExpirationDate  = $lic.ExpirationDate
            "VI SDK Server" = $about.FullName
            "VI SDK UUID"   = $about.InstanceUuid
        }
    }
   
}

function Get-vHealth {
    param($about)
   
    foreach ($vcAlarm in Get-VIEvent -Start (Get-Date).AddDays(-14) -MaxSamples ([int]::MaxValue) | Where-Object { $_ -is [VMware.Vim.AlarmStatusChangedEvent] -and ($_.To -eq "Yellow" -or $_.To -eq "Red") -and $_.To -ne "Gray" }) {
     
        [PSCustomObject]@{
                
            AlarmName       = $vcAlarm.alarm.name
            AlarmStatus     = $vcAlarm.to
            AlarmData       = $vcAlarm.createdtime
            "VI SDK Server" = $about.FullName
            "VI SDK UUID"   = $about.InstanceUuid
        }
        
    }
    
}



# ---- CALL FUNCTIONS AND EXPORT ALL TO EXCEL ----


$vinfo = Get-vInfo      -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vcpu = Get-vCPU       -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vmemory = Get-vMemory    -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vdisk = Get-vDisk      -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vpartition = Get-vPartition -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vSCSI = Get-vSCSI  -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vnetwork = Get-vNetwork   -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vfloppy = Get-vFloppy    -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vcd = Get-vCD        -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vsnapshot = Get-vSnapshot  -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vtools = Get-vTools     -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext}
$vrp = Get-vRP        -rpools $rpools -about $about -GetVMContextFn ${function:Get-VMContext}
$vcluster = Get-vCluster   -clusters $clusters -about $about
$vhost = Get-vHost      -ESXhosts $ESXhosts -about $about
$vTLShost = Get-vTLShosts  -ESXhosts $ESXhosts -about $about
$vhba = Get-vHBA       -ESXhosts $ESXhosts -about $about
$vnic = Get-vNIC       -ESXhosts $ESXhosts -about $about
$vswitch = Get-vSwitch    -ESXhosts $ESXhosts -about $about
$vport = Get-vPort      -ESXhosts $ESXhosts -about $about
$dvswitch = Get-vdvSwitch  -about $about
$dvport = Get-vdvPort    -about $about
$vsc_vmk = Get-vSC_VMK    -ESXhosts $ESXhosts -about $about
$vdatastore = Get-vDatastore -datastores $datastores -about $about
$vZombie = Get-VmwOrphan  -Datastores $datastores -about $about 
$vlicense = Get-vLicense   -about $about
$vhealth = Get-vHealth    -about $about

Write-Host "DMTools exporting to Excel file $excelFile" -ForegroundColor Blue
$vinfo      | Export-Excel $excelFile -WorksheetName 'vInfo'     -AutoSize -AutoFilter 
$vcpu       | Export-Excel $excelFile -WorksheetName 'vCPU'      -AutoSize -AutoFilter -Append
$vmemory    | Export-Excel $excelFile -WorksheetName 'vMemory'   -AutoSize -AutoFilter -Append
$vdisk      | Export-Excel $excelFile -WorksheetName 'vDisk'     -AutoSize -AutoFilter -Append
$vpartition | Export-Excel $excelFile -WorksheetName 'vPartition' -AutoSize -AutoFilter -Append
$vSCSI      | Export-Excel $excelFile -WorksheetName 'vSCSI'      -AutoSize -AutoFilter -Append
$vnetwork   | Export-Excel $excelFile -WorksheetName 'vNetwork'  -AutoSize -AutoFilter -Append
$vfloppy    | Export-Excel $excelFile -WorksheetName 'vFloppy'   -AutoSize -AutoFilter -Append
$vcd        | Export-Excel $excelFile -WorksheetName 'vCD'       -AutoSize -AutoFilter -Append
$vsnapshot  | Export-Excel $excelFile -WorksheetName 'vSnapshot' -AutoSize -AutoFilter -Append
$vtools     | Export-Excel $excelFile -WorksheetName 'vTools'    -AutoSize -AutoFilter -Append
$vrp        | Export-Excel $excelFile -WorksheetName 'vRP'       -AutoSize -AutoFilter -Append
$vcluster   | Export-Excel $excelFile -WorksheetName 'vCluster'  -AutoSize -AutoFilter -Append
$vhost      | Export-Excel $excelFile -WorksheetName 'vHost'     -AutoSize -AutoFilter -Append
$vTLShost   | Export-Excel $excelFile -WorksheetName 'vTLShost'     -AutoSize -AutoFilter -Append
$vhba       | Export-Excel $excelFile -WorksheetName 'vHBA'      -AutoSize -AutoFilter -Append
$vnic       | Export-Excel $excelFile -WorksheetName 'vNIC'      -AutoSize -AutoFilter -Append
$vswitch    | Export-Excel $excelFile -WorksheetName 'vSwitch'   -AutoSize -AutoFilter -Append
$vport      | Export-Excel $excelFile -WorksheetName 'vPort'     -AutoSize -AutoFilter -Append
$dvswitch   | Export-Excel $excelFile -WorksheetName 'dvSwitch'  -AutoSize -AutoFilter -Append
$dvport     | Export-Excel $excelFile -WorksheetName 'dvPort'    -AutoSize -AutoFilter -Append
$vsc_vmk    | Export-Excel $excelFile -WorksheetName 'vSC_VMK'   -AutoSize -AutoFilter -Append
$vdatastore | Export-Excel $excelFile -WorksheetName 'vDatastore'-AutoSize -AutoFilter -Append
$vZombie    | Export-Excel $excelFile -WorksheetName 'vZombieFiles' -AutoSize -AutoFilter -Append 
$vlicense   | Export-Excel $excelFile -WorksheetName 'vLicense'  -AutoSize -AutoFilter -Append
$vhealth    | Export-Excel $excelFile -WorksheetName 'vHealth'   -AutoSize -AutoFilter -Append

Disconnect-VIServer -Force -Confirm:$false 
Write-Host "DMTools Excel export complete: $excelFile" -ForegroundColor Green
