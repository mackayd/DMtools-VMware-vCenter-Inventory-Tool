param(
    [Parameter(Mandatory)]
    [string]$InputWorkbook,

    [string]$OutputHtml
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not installed."
    }
    Import-Module $Name -ErrorAction Stop | Out-Null
}

function Get-Sheet {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )
    try {
        return @(Import-Excel -Path $Path -WorksheetName $Name -ErrorAction Stop)
    }
    catch {
        return @()
    }
}

function Get-CellValue {
    param($Row, [string]$Name)
    try { return $Row.$Name } catch { return $null }
}

function Get-FirstValue {
    param(
        [Parameter(Mandatory)] $Row,
        [Parameter(Mandatory)][string[]]$Names
    )
    foreach ($name in $Names) {
        $value = Get-CellValue -Row $Row -Name $name
        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            return $value
        }
    }
    return $null
}

function Get-String {
    param($Value)
    if ($null -eq $Value) { return $null }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s.Trim()
}

function Parse-DatastoreFromPath {
    param([string]$DiskPath)
    if ([string]::IsNullOrWhiteSpace($DiskPath)) { return $null }
    $m = [regex]::Match($DiskPath, '^\[(?<ds>[^\]]+)\]')
    if ($m.Success) { return $m.Groups['ds'].Value }
    return $null
}

function New-Slug {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'unknown' }
    $slug = ($Value.ToLowerInvariant() -replace '[^a-z0-9]+','-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($slug)) { return 'unknown' }
    return $slug
}

function Get-SafeCount {
    param($Value)
    if ($null -eq $Value) { return 0 }
    if ($Value -is [string]) { return 1 }
    if ($Value -is [System.Collections.ICollection]) { return $Value.Count }
    if ($Value.PSObject.Properties.Name -contains 'Count') {
        try { return [int]$Value.Count } catch {}
    }
    return @($Value).Count
}

function Add-ToListIndex {
    param(
        [Parameter(Mandatory)][hashtable]$Index,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = New-Object System.Collections.ArrayList
    }
    $null = $Index[$Key].Add($Value)
}

function Add-UniqueStringToListIndex {
    param(
        [Parameter(Mandatory)][hashtable]$Index,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Key) -or [string]::IsNullOrWhiteSpace($Value)) { return }
    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = New-Object System.Collections.ArrayList
    }
    if (-not ($Index[$Key] -contains $Value)) {
        $null = $Index[$Key].Add($Value)
    }
}

function Set-NameIndex {
    param(
        [Parameter(Mandatory)][hashtable]$Index,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)]$Value
    )
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = $Value
    }
}

Ensure-Module -Name ImportExcel

if (-not (Test-Path -LiteralPath $InputWorkbook)) {
    throw "Input workbook not found: $InputWorkbook"
}

if ([string]::IsNullOrWhiteSpace($OutputHtml)) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InputWorkbook)
    $dir  = Split-Path -Parent $InputWorkbook
    $OutputHtml = Join-Path $dir ($base + '-InventoryReport.html')
}

$vInfo      = @(Get-Sheet -Path $InputWorkbook -Name 'vInfo')
$vTools     = @(Get-Sheet -Path $InputWorkbook -Name 'vTools')
$vDisk      = @(Get-Sheet -Path $InputWorkbook -Name 'vDisk')
$vPartition = @(Get-Sheet -Path $InputWorkbook -Name 'vPartition')
$vNetwork   = @(Get-Sheet -Path $InputWorkbook -Name 'vNetwork')
$vSnapshot  = @(Get-Sheet -Path $InputWorkbook -Name 'vSnapshot')
$vHost      = @(Get-Sheet -Path $InputWorkbook -Name 'vHost')
$vCluster   = @(Get-Sheet -Path $InputWorkbook -Name 'vCluster')
$vDatastore = @(Get-Sheet -Path $InputWorkbook -Name 'vDatastore')
$vNIC       = @(Get-Sheet -Path $InputWorkbook -Name 'vNIC')
$vSwitch    = @(Get-Sheet -Path $InputWorkbook -Name 'vSwitch')
$vPort      = @(Get-Sheet -Path $InputWorkbook -Name 'vPort')
$dvSwitch   = @(Get-Sheet -Path $InputWorkbook -Name 'dvSwitch')
$dvPort     = @(Get-Sheet -Path $InputWorkbook -Name 'dvPort')
$vSC_VMK    = @(Get-Sheet -Path $InputWorkbook -Name 'vSC_VMK')

$vmToolsByName = @{}
foreach ($row in $vTools) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('Name','VM'))
    if ($name) { $vmToolsByName[$name] = $row }
}

$vmNetworksByName = @{}
$networkVmIndex = @{}
foreach ($row in $vNetwork) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('VM','Name'))
    if (-not $name) { continue }

    $networkName = Get-String (Get-FirstValue -Row $row -Names @('Network','Portgroup','PortGroup'))
    $entry = [ordered]@{
        label      = Get-String (Get-CellValue $row 'Label')
        macAddress = Get-String (Get-FirstValue -Row $row -Names @('Mac Address','MAC'))
        network    = $networkName
        connected  = Get-CellValue $row 'Connected'
        type       = Get-String (Get-CellValue $row 'Type')
        ipAddress  = Get-String (Get-FirstValue -Row $row -Names @('IP Address','IPAddress'))
    }

    Add-ToListIndex -Index $vmNetworksByName -Key $name -Value $entry

    if ($networkName) {
        Add-ToListIndex -Index $networkVmIndex -Key $networkName -Value ([ordered]@{
            vm         = $name
            label      = $entry.label
            macAddress = $entry.macAddress
            connected  = $entry.connected
            type       = $entry.type
            ipAddress  = $entry.ipAddress
        })
    }
}

$vmDisksByName = @{}
$vmDatastoreIndex = @{}
foreach ($row in $vDisk) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('VM','Name'))
    if (-not $name) { continue }

    $diskPath = Get-String (Get-CellValue $row 'Disk Path')
    $datastore = Parse-DatastoreFromPath -DiskPath $diskPath

    $diskObj = [ordered]@{
        label        = Get-String (Get-CellValue $row 'Label')
        diskPath     = $diskPath
        datastore    = $datastore
        capacityMB   = Get-CellValue $row 'Capacity MB'
        persistence  = Get-String (Get-CellValue $row 'Persistence')
        thin         = Get-CellValue $row 'Thin'
        diskMode     = Get-String (Get-FirstValue -Row $row -Names @('Mode','Disk Mode'))
    }

    Add-ToListIndex -Index $vmDisksByName -Key $name -Value $diskObj

    if ($datastore) {
        Add-UniqueStringToListIndex -Index $vmDatastoreIndex -Key $datastore -Value $name
    }
}

$vmDiskIndex = @{}
foreach ($row in $vDisk) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('VM','Name'))
    if (-not $name) { continue }
    $diskKey = Get-FirstValue -Row $row -Names @('Disk Key','DiskKey')
    if ($null -eq $diskKey -or [string]::IsNullOrWhiteSpace([string]$diskKey)) { continue }
    $compositeKey = '{0}|{1}' -f $name, [string]$diskKey
    $vmDiskIndex[$compositeKey] = [ordered]@{
        diskKey    = $diskKey
        label      = Get-String (Get-FirstValue -Row $row -Names @('Disk','Hard Disk'))
        diskPath   = Get-String (Get-FirstValue -Row $row -Names @('Disk Path','Path'))
        capacityMB = Get-FirstValue -Row $row -Names @('Capacity MiB','Capacity MB')
        mode       = Get-String (Get-FirstValue -Row $row -Names @('Disk Mode','Mode'))
        thin       = Get-FirstValue -Row $row -Names @('Thin')
        datastore  = Parse-DatastoreFromPath -DiskPath (Get-String (Get-FirstValue -Row $row -Names @('Disk Path','Path')))
    }
}

$vmPartitionsByName = @{}
foreach ($row in $vPartition) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('VM','Name'))
    if (-not $name) { continue }

    $diskKey = Get-FirstValue -Row $row -Names @('Disk Key','DiskKey')
    $lookup = $null
    $compositeKey = if ($null -ne $diskKey) { '{0}|{1}' -f $name, [string]$diskKey } else { $null }
    if ($compositeKey -and $vmDiskIndex.ContainsKey($compositeKey)) { $lookup = $vmDiskIndex[$compositeKey] }

    Add-ToListIndex -Index $vmPartitionsByName -Key $name -Value ([ordered]@{
        partitionPath = Get-String (Get-FirstValue -Row $row -Names @('Partition Path','Disk'))
        diskKey       = $diskKey
        capacityMB    = Get-FirstValue -Row $row -Names @('Capacity MiB','Capacity MB')
        consumedMB    = Get-FirstValue -Row $row -Names @('Consumed MiB','Consumed MB')
        freeMB        = Get-FirstValue -Row $row -Names @('Free MiB','Free MB')
        freePct       = Get-FirstValue -Row $row -Names @('Free %','FreePct')
        vmdkLabel     = if ($lookup) { $lookup.label } else { Get-String (Get-FirstValue -Row $row -Names @('VMDK')) }
        vmdkPath      = if ($lookup) { $lookup.diskPath } else { Get-String (Get-FirstValue -Row $row -Names @('VMDK Path')) }
        datastore     = if ($lookup) { $lookup.datastore } else { Parse-DatastoreFromPath -DiskPath (Get-String (Get-FirstValue -Row $row -Names @('VMDK Path'))) }
    })
}

$vmSnapshotsByName = @{}
foreach ($row in $vSnapshot) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('VM','Name'))
    if (-not $name) { continue }

    Add-ToListIndex -Index $vmSnapshotsByName -Key $name -Value ([ordered]@{
        name        = Get-String (Get-CellValue $row 'Name')
        description = Get-String (Get-CellValue $row 'Description')
        created     = Get-String (Get-CellValue $row 'Created')
        sizeMB      = Get-CellValue $row 'Size MB'
    })
}

$hostPnicsByName = @{}
foreach ($row in $vNIC) {
    $vmHostName = Get-String (Get-FirstValue -Row $row -Names @('Host','Name'))
    if (-not $vmHostName) { continue }

    Add-ToListIndex -Index $hostPnicsByName -Key $vmHostName -Value ([ordered]@{
        device    = Get-String (Get-FirstValue -Row $row -Names @('PNICDevice','Device'))
        mac       = Get-String (Get-FirstValue -Row $row -Names @('MAC','Mac Address'))
        linkSpeed = Get-CellValue $row 'LinkSpeed'
        duplex    = Get-CellValue $row 'Duplex'
        driver    = Get-String (Get-CellValue $row 'Driver')
        switch    = Get-String (Get-FirstValue -Row $row -Names @('Switch','vSwitch'))
        uplink    = Get-String (Get-CellValue $row 'UplinkPort')
        pci       = Get-String (Get-CellValue $row 'PCI')
    })
}

$hostVSwitchesByName = @{}
foreach ($row in $vSwitch) {
    $vmHostName = Get-String (Get-FirstValue -Row $row -Names @('Host','Name'))
    if (-not $vmHostName) { continue }

    Add-ToListIndex -Index $hostVSwitchesByName -Key $vmHostName -Value ([ordered]@{
        name       = Get-String (Get-FirstValue -Row $row -Names @('vSwitch','Name'))
        mtu        = Get-CellValue $row 'MTU'
        nic        = Get-String (Get-CellValue $row 'Nic')
        activeNic  = Get-String (Get-CellValue $row 'ActiveNic')
        standbyNic = Get-String (Get-CellValue $row 'StandbyNic')
        promiscuous= Get-CellValue $row 'AllowPromiscuous'
        macChanges = Get-CellValue $row 'MacChanges'
        forged     = Get-CellValue $row 'ForgedTransmits'
    })
}

$hostPortGroupsByName = @{}
$networkPortGroupIndex = @{}
foreach ($row in $vPort) {
    $vmHostName = Get-String (Get-FirstValue -Row $row -Names @('Host','Name'))
    $portGroupName = Get-String (Get-FirstValue -Row $row -Names @('PortGroup','Portgroup','Name'))
    if (-not $vmHostName -or -not $portGroupName) { continue }

    $portObj = [ordered]@{
        host       = $vmHostName
        cluster    = Get-String (Get-CellValue $row 'Cluster')
        portGroup  = $portGroupName
        vSwitch    = Get-String (Get-CellValue $row 'vSwitch')
        vlanId     = Get-CellValue $row 'VLANId'
        numPorts   = Get-CellValue $row 'NumPorts'
        activePorts= Get-CellValue $row 'ActivePorts'
    }
    Add-ToListIndex -Index $hostPortGroupsByName -Key $vmHostName -Value $portObj
    Add-ToListIndex -Index $networkPortGroupIndex -Key $portGroupName -Value $portObj
}

$hostVmksByName = @{}
$networkVmkIndex = @{}
foreach ($row in $vSC_VMK) {
    $vmHostName = Get-String (Get-FirstValue -Row $row -Names @('Host','Name'))
    $portGroupName = Get-String (Get-FirstValue -Row $row -Names @('PortGroup','Portgroup','Name'))
    if (-not $vmHostName) { continue }

    $vmkObj = [ordered]@{
        host        = $vmHostName
        adapter     = Get-String (Get-FirstValue -Row $row -Names @('VMKernelAdapter','Name'))
        ipAddress   = Get-String (Get-FirstValue -Row $row -Names @('IPAddress','IP Address'))
        subnetMask  = Get-String (Get-CellValue $row 'SubnetMask')
        vSwitch     = Get-String (Get-CellValue $row 'vSwitch')
        portGroup   = $portGroupName
        management  = Get-CellValue $row 'Management'
        vMotion     = Get-CellValue $row 'vMotion'
        vSAN        = Get-CellValue $row 'vSAN'
        faultTol    = Get-CellValue $row 'FaultTolerance'
        provisioning= Get-CellValue $row 'Provisioning'
    }
    Add-ToListIndex -Index $hostVmksByName -Key $vmHostName -Value $vmkObj

    if ($portGroupName) {
        Add-ToListIndex -Index $networkVmkIndex -Key $portGroupName -Value $vmkObj
    }
}

$dvSwitchByName = @{}
foreach ($row in $dvSwitch) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('Name','vSwitch'))
    if (-not $name) { continue }
    $dvSwitchByName[$name] = [ordered]@{
        name        = $name
        datacenter  = Get-String (Get-CellValue $row 'Datacenter')
        version     = Get-String (Get-CellValue $row 'Version')
        mtu         = Get-CellValue $row 'MTU'
        uplinks     = Get-String (Get-CellValue $row 'Uplinks')
        numPorts    = Get-CellValue $row 'NumPorts'
        numUplinks  = Get-CellValue $row 'NumUplinks'
    }
}

$networkDvPortIndex = @{}
foreach ($row in $dvPort) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('Name','PortGroup'))
    if (-not $name) { continue }
    Add-ToListIndex -Index $networkDvPortIndex -Key $name -Value ([ordered]@{
        name       = $name
        vdSwitch   = Get-String (Get-FirstValue -Row $row -Names @('VDSwitch','dvSwitch'))
        vlanId     = Get-CellValue $row 'VLANId'
        numPorts   = Get-CellValue $row 'NumPorts'
        activePorts= Get-CellValue $row 'ActivePorts'
        type       = Get-String (Get-CellValue $row 'Type')
        description= Get-String (Get-CellValue $row 'Description')
    })
}

$vmInfoByName = @{}
$hostVmIndex = @{}
$clusterVmIndex = @{}
$clusterHostIndex = @{}
$vms = @(
    foreach ($row in $vInfo) {
        $name = Get-String (Get-FirstValue -Row $row -Names @('VM','Name'))
        if (-not $name) { continue }

        $vmHostName = Get-String (Get-FirstValue -Row $row -Names @('Host'))
        $clusterName = Get-String (Get-FirstValue -Row $row -Names @('Cluster','Name'))
        $datacenterName = Get-String (Get-FirstValue -Row $row -Names @('Datacenter'))
        $powerstate = Get-String (Get-CellValue $row 'Powerstate')
        $template = [bool](Get-CellValue $row 'Template')
        $vmSlug = New-Slug $name

        if ($vmHostName) {
            Add-UniqueStringToListIndex -Index $hostVmIndex -Key $vmHostName -Value $name
        }
        if ($clusterName) {
            Add-UniqueStringToListIndex -Index $clusterVmIndex -Key $clusterName -Value $name
            if ($vmHostName) {
                Add-UniqueStringToListIndex -Index $clusterHostIndex -Key $clusterName -Value $vmHostName
            }
        }

        $toolRow = $null
        if ($vmToolsByName.ContainsKey($name)) { $toolRow = $vmToolsByName[$name] }

        $disks = if ($vmDisksByName.ContainsKey($name)) { @($vmDisksByName[$name]) } else { @() }
        $networks = if ($vmNetworksByName.ContainsKey($name)) { @($vmNetworksByName[$name]) } else { @() }
        $snapshots = if ($vmSnapshotsByName.ContainsKey($name)) { @($vmSnapshotsByName[$name]) } else { @() }
        $partitions = if ($vmPartitionsByName.ContainsKey($name)) { @($vmPartitionsByName[$name]) } else { @() }

        $vmObj = [ordered]@{
            id             = $vmSlug
            name           = $name
            powerState     = $powerstate
            template       = $template
            datacenter     = $datacenterName
            cluster        = $clusterName
            host           = $vmHostName
            dnsName        = Get-String (Get-CellValue $row 'DNS Name')
            primaryIP      = Get-String (Get-FirstValue -Row $row -Names @('Primary IP Address','IPAddress'))
            osConfig       = Get-String (Get-CellValue $row 'OS according to the configuration file')
            osTools        = Get-String (Get-CellValue $row 'OS according to the VMware Tools')
            cpus           = Get-CellValue $row 'CPUs'
            memoryMB       = Get-CellValue $row 'Memory'
            nics           = Get-CellValue $row 'Nics'
            disksCount     = Get-CellValue $row 'Disks'
            annotation     = Get-String (Get-CellValue $row 'Annotation')
            configStatus   = Get-String (Get-CellValue $row 'Config status')
            toolsStatus    = if ($toolRow) { Get-String (Get-CellValue $toolRow 'Tools') } else { $null }
            toolsVersion   = if ($toolRow) { Get-CellValue $toolRow 'ToolsVersion' } else { $null }
            upgradeable    = if ($toolRow) { Get-String (Get-CellValue $toolRow 'Upgradeable') } else { $null }
            requiredTools  = if ($toolRow) { Get-String (Get-CellValue $toolRow 'Required Version') } else { $null }
            snapshotCount  = Get-SafeCount $snapshots
            datastores     = @($disks | ForEach-Object { $_.datastore } | Where-Object { $_ } | Sort-Object -Unique)
            networks       = @($networks)
            disks          = @($disks)
            snapshots      = @($snapshots)
            partitions     = @($partitions)
            statusClass    = if ($powerstate -eq 'PoweredOn') { 'PASS' } elseif ($powerstate -match 'PoweredOff|poweredOff') { 'INFO' } else { 'WARN' }
        }

        $vmInfoByName[$name] = $vmObj
        $vmObj
    }
)

$hostRowsByName = @{}
foreach ($row in $vHost) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('Host','Name'))
    if ($name) { Set-NameIndex -Index $hostRowsByName -Key $name -Value $row }
}

$clusterRowsByName = @{}
foreach ($row in $vCluster) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('Cluster','Name'))
    if ($name) { Set-NameIndex -Index $clusterRowsByName -Key $name -Value $row }
}

$datastoreRowsByName = @{}
foreach ($row in $vDatastore) {
    $name = Get-String (Get-FirstValue -Row $row -Names @('Datastore','Name'))
    if ($name) { Set-NameIndex -Index $datastoreRowsByName -Key $name -Value $row }
}

$allHostNames = @(
    @($hostRowsByName.Keys) +
    @($hostVmIndex.Keys) +
    @($hostPnicsByName.Keys) +
    @($hostVSwitchesByName.Keys) +
    @($hostPortGroupsByName.Keys) +
    @($hostVmksByName.Keys)
) | Where-Object { $_ } | Sort-Object -Unique

$hosts = @(
    foreach ($vmHostName in $allHostNames) {
        $row = $null
        if ($hostRowsByName.ContainsKey($vmHostName)) { $row = $hostRowsByName[$vmHostName] }

        $vmHostNameVMs   = if ($hostVmIndex.ContainsKey($vmHostName)) { @($hostVmIndex[$vmHostName] | Sort-Object -Unique) } else { @() }
        $hostPnics       = if ($hostPnicsByName.ContainsKey($vmHostName)) { @($hostPnicsByName[$vmHostName]) } else { @() }
        $hostVSwitches   = if ($hostVSwitchesByName.ContainsKey($vmHostName)) { @($hostVSwitchesByName[$vmHostName]) } else { @() }
        $hostPortGroups  = if ($hostPortGroupsByName.ContainsKey($vmHostName)) { @($hostPortGroupsByName[$vmHostName]) } else { @() }
        $hostVmks        = if ($hostVmksByName.ContainsKey($vmHostName)) { @($hostVmksByName[$vmHostName]) } else { @() }

        $connectedNetworks = @(
            @($hostPortGroups | ForEach-Object { $_.portGroup }) +
            @($hostVmks | ForEach-Object { $_.portGroup }) +
            @($vmHostNameVMs | ForEach-Object {
                if ($vmInfoByName.ContainsKey($_)) { $vmInfoByName[$_].networks | ForEach-Object { $_.network } }
            })
        ) | Where-Object { $_ } | Sort-Object -Unique

        [ordered]@{
            id           = New-Slug $vmHostName
            name         = $vmHostName
            cluster      = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Cluster')) } else { Get-String ((@($vms | Where-Object { $_.host -eq $vmHostName } | Select-Object -ExpandProperty cluster -First 1))) }
            datacenter   = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Datacenter')) } else { Get-String ((@($vms | Where-Object { $_.host -eq $vmHostName } | Select-Object -ExpandProperty datacenter -First 1))) }
            powerState   = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Powerstate','Status')) } else { $null }
            connection   = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Connection State','ConnectionState')) } else { $null }
            version      = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Version','ESXiVersion')) } else { $null }
            build        = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Build')) } else { $null }
            model        = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Model')) } else { $null }
            vendor       = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Vendor')) } else { $null }
            cpuModel     = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('CPU Model','CPUModel')) } else { $null }
            memoryMB     = if ($row) { Get-FirstValue -Row $row -Names @('Memory Total MB','Memory') } else { $null }
            vmCount      = Get-SafeCount $vmHostNameVMs
            vms          = @($vmHostNameVMs)
            pnics        = @($hostPnics)
            vswitches    = @($hostVSwitches)
            portgroups   = @($hostPortGroups)
            vmkernels    = @($hostVmks)
            networks     = @($connectedNetworks)
        }
    }
)

$allClusterNames = @(
    @($clusterRowsByName.Keys) +
    @($clusterVmIndex.Keys) +
    @($clusterHostIndex.Keys) +
    @($vms | ForEach-Object { $_.cluster })
) | Where-Object { $_ } | Sort-Object -Unique

$clusters = @(
    foreach ($clusterName in $allClusterNames) {
        $row = $null
        if ($clusterRowsByName.ContainsKey($clusterName)) { $row = $clusterRowsByName[$clusterName] }

        $clusterVMs = if ($clusterVmIndex.ContainsKey($clusterName)) { @($clusterVmIndex[$clusterName] | Sort-Object -Unique) } else { @() }
        $clusterHosts = if ($clusterHostIndex.ContainsKey($clusterName)) { @($clusterHostIndex[$clusterName] | Sort-Object -Unique) } else { @() }

        [ordered]@{
            id         = New-Slug $clusterName
            name       = $clusterName
            datacenter = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Datacenter')) } else { Get-String ((@($vms | Where-Object { $_.cluster -eq $clusterName } | Select-Object -ExpandProperty datacenter -First 1))) }
            hosts      = @($clusterHosts)
            vms        = @($clusterVMs)
            drsEnabled = if ($row) { Get-FirstValue -Row $row -Names @('DRS enabled') } else { $null }
            haEnabled  = if ($row) { Get-FirstValue -Row $row -Names @('HA enabled') } else { $null }
            vsanEnabled= if ($row) { Get-FirstValue -Row $row -Names @('vSAN enabled') } else { $null }
            evcMode    = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('EVC Mode')) } else { $null }
        }
    }
)

$allDatastoreNames = @(
    @($datastoreRowsByName.Keys) +
    @($vmDatastoreIndex.Keys)
) | Where-Object { $_ } | Sort-Object -Unique

$datastores = @(
    foreach ($datastoreName in $allDatastoreNames) {
        $row = $null
        if ($datastoreRowsByName.ContainsKey($datastoreName)) { $row = $datastoreRowsByName[$datastoreName] }
        $relatedVMs = if ($vmDatastoreIndex.ContainsKey($datastoreName)) { @($vmDatastoreIndex[$datastoreName] | Sort-Object -Unique) } else { @() }

        [ordered]@{
            id          = New-Slug $datastoreName
            name        = $datastoreName
            type        = if ($row) { Get-String (Get-FirstValue -Row $row -Names @('Type')) } else { $null }
            capacityMB  = if ($row) { Get-FirstValue -Row $row -Names @('Capacity MB') } else { $null }
            freeMB      = if ($row) { Get-FirstValue -Row $row -Names @('Free MB') } else { $null }
            provisioned = if ($row) { Get-FirstValue -Row $row -Names @('Provisioned MB') } else { $null }
            accessible  = if ($row) { Get-FirstValue -Row $row -Names @('Accessible') } else { $null }
            multiHost   = if ($row) { Get-FirstValue -Row $row -Names @('MHA') } else { $null }
            vmCount     = Get-SafeCount $relatedVMs
            vms         = @($relatedVMs)
        }
    }
)

$networkNames = @(
    @($networkVmIndex.Keys) +
    @($networkPortGroupIndex.Keys) +
    @($networkDvPortIndex.Keys) +
    @($networkVmkIndex.Keys)
) | Where-Object { $_ } | Sort-Object -Unique

$networks = @(
    foreach ($networkName in $networkNames) {
        $vmAttachments = if ($networkVmIndex.ContainsKey($networkName)) { @($networkVmIndex[$networkName]) } else { @() }
        $stdPortGroups = if ($networkPortGroupIndex.ContainsKey($networkName)) { @($networkPortGroupIndex[$networkName]) } else { @() }
        $dvPortGroups  = if ($networkDvPortIndex.ContainsKey($networkName)) { @($networkDvPortIndex[$networkName]) } else { @() }
        $vmkernels     = if ($networkVmkIndex.ContainsKey($networkName)) { @($networkVmkIndex[$networkName]) } else { @() }

        $attachedVMs = @($vmAttachments | ForEach-Object { $_.vm } | Where-Object { $_ } | Sort-Object -Unique)
        $attachedHosts = @(
            @($vmAttachments | ForEach-Object {
                if ($vmInfoByName.ContainsKey($_.vm)) { $vmInfoByName[$_.vm].host }
            }) +
            @($stdPortGroups | ForEach-Object { $_.host }) +
            @($vmkernels | ForEach-Object { $_.host })
        ) | Where-Object { $_ } | Sort-Object -Unique

        $standardSwitches = @($stdPortGroups | ForEach-Object { $_.vSwitch } | Where-Object { $_ } | Sort-Object -Unique)
        $distributedSwitches = @($dvPortGroups | ForEach-Object { $_.vdSwitch } | Where-Object { $_ } | Sort-Object -Unique)
        $vlanHints = @(
            @($stdPortGroups | ForEach-Object { $_.vlanId }) +
            @($dvPortGroups | ForEach-Object { $_.vlanId })
        ) | Where-Object { $_ -ne $null -and $_ -ne '' } | Sort-Object -Unique

        [ordered]@{
            id                  = New-Slug $networkName
            name                = $networkName
            attachedVMs         = @($attachedVMs)
            attachedHosts       = @($attachedHosts)
            vmAttachments       = @($vmAttachments)
            standardPortGroups  = @($stdPortGroups)
            distributedPortGroups = @($dvPortGroups)
            vmkernels           = @($vmkernels)
            standardSwitches    = @($standardSwitches)
            distributedSwitches = @($distributedSwitches)
            vlanHints           = @($vlanHints)
        }
    }
)

$platformName = $null
$apiVersion = $null
if ((Get-SafeCount $vInfo) -gt 0) {
    $platformName = Get-String (Get-CellValue $vInfo[0] 'VI SDK Server')
    $apiVersion   = Get-String (Get-CellValue $vInfo[0] 'VI SDK API Version')
}

$templates = @($vms | Where-Object { $_.template })
$vms = @($vms | Where-Object { -not $_.template })

$summary = [ordered]@{
    generatedAt    = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    vcenter        = $platformName
    apiVersion     = $apiVersion
    totalVMs       = Get-SafeCount $vms
    poweredOnVMs   = Get-SafeCount @($vms | Where-Object { $_.powerState -eq 'PoweredOn' })
    templates      = Get-SafeCount $templates
    hosts          = Get-SafeCount $hosts
    clusters       = Get-SafeCount $clusters
    datastores     = Get-SafeCount $datastores
    networks       = Get-SafeCount $networks
}

$reportData = [ordered]@{
    summary    = $summary
    vms        = @($vms)
    templates  = @($templates)
    hosts      = @($hosts)
    clusters   = @($clusters)
    datastores = @($datastores)
    networks   = @($networks)
}

$json = $reportData | ConvertTo-Json -Depth 12 -Compress

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>DMTools Inventory Report</title>
<style>
:root{--bg:#0b1220;--panel:#111a2b;--text:#e8eef8;--muted:#9fb0c8;--border:#243650;--pass:#2e7d32;--warn:#f9a825;--fail:#c62828;--info:#1565c0;--shadow:0 12px 40px rgba(0,0,0,.35)}
*{box-sizing:border-box}
body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:linear-gradient(180deg,#09101d,#0f1728 25%,#0d1321);color:var(--text)}
.container{max-width:1720px;margin:0 auto;padding:24px}
.hero{background:linear-gradient(135deg,#15315c,#101b31 55%,#113b34);border:1px solid var(--border);border-radius:24px;padding:28px;box-shadow:var(--shadow);margin-bottom:20px}
.hero-top{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap}
.hero h1{margin:0 0 8px;font-size:32px}
.hero p{margin:0;color:#d2def0}
.hero-note{font-size:13px;color:#aebfd6;margin-top:8px}
.badge{display:inline-flex;align-items:center;justify-content:center;min-width:70px;padding:6px 12px;border-radius:999px;font-weight:700;font-size:12px;letter-spacing:.4px;text-transform:uppercase}
.badge.PASS{background:#1b5e20;color:#d6ffd6}.badge.WARN{background:#7a5a00;color:#fff0b5}.badge.FAIL{background:#7f1d1d;color:#ffd6d6}.badge.INFO{background:#0d47a1;color:#dce9ff}.badge.large{padding:10px 16px;font-size:13px}
.hero-grid{display:grid;grid-template-columns:2fr 1fr;gap:20px;margin-top:24px}
.stack-card,.score-card,.sidebar-card,.panel-card,.detail-card,.stat-card,.entity-card{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:20px;padding:20px}
.stack-bar{display:flex;height:54px;border-radius:18px;overflow:hidden;margin-top:18px;background:#0f1728;border:1px solid rgba(255,255,255,.1)}
.stack-segment{display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;white-space:nowrap}.stack-segment span{padding:0 10px}
.stack-segment.PASS{background:var(--pass)}.stack-segment.INFO{background:var(--info)}
.score-card{display:flex;align-items:center;justify-content:center}
.score-ring{width:180px;height:180px;border-radius:50%;background:conic-gradient(var(--pass) 0 68%, var(--info) 68% 100%);display:flex;align-items:center;justify-content:center}
.score-ring-content{width:130px;height:130px;border-radius:50%;background:#0f1728;display:flex;flex-direction:column;align-items:center;justify-content:center}
.score-ring-value{font-size:44px;font-weight:800}.score-ring-label{color:var(--muted);font-size:13px;text-transform:uppercase;letter-spacing:.06em}
.stats{display:grid;grid-template-columns:repeat(6,1fr);gap:16px;margin:20px 0}
.stat-label{color:var(--muted);font-size:13px}.stat-value{font-size:34px;font-weight:800;margin-top:8px}
.main-grid{display:grid;grid-template-columns:340px 1fr;gap:20px;align-items:start}.sidebar,.content{min-width:0}.sidebar-card{position:sticky;top:20px}
.section-title{margin:0 0 6px;font-size:18px}.section-subtitle{margin:0 0 16px;color:var(--muted);font-size:13px}
.filter-bar,.view-mode-toggle,.entity-toggle{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0 18px}
button{cursor:pointer;border:none}
.filter-btn,.context-filter-btn,.view-btn,.entity-btn,.view-item-btn,.nav-link{padding:10px 14px;border-radius:12px;background:#18253b;color:var(--text);border:1px solid var(--border);font-weight:600}
.filter-btn.active,.context-filter-btn.active,.view-btn.active,.entity-btn.active{background:#1d4ed8;border-color:#3b82f6}
.item-nav{display:flex;flex-direction:column;gap:10px;max-height:calc(100vh - 320px);overflow:auto;padding-right:4px}
.item-nav button{width:100%;text-align:left;padding:12px 14px;border-radius:14px;background:#141f33;border:1px solid var(--border);color:var(--text);display:flex;flex-direction:column;gap:4px}
.item-nav button.active{outline:2px solid #3b82f6;background:#18253b}
.item-nav-title{font-weight:800}.item-nav-meta,.item-nav-sub{font-size:12px;color:var(--muted)}
.panel{display:none;flex-direction:column;gap:18px}.panel.active{display:flex}
.panel-header{display:flex;justify-content:space-between;align-items:flex-start;gap:14px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:20px;padding:20px}
.eyebrow{text-transform:uppercase;letter-spacing:.08em;font-size:11px;color:var(--muted);font-weight:700}.panel-header h2{margin:4px 0 6px;font-size:28px}.panel-subtitle{margin:0;color:var(--muted)}
.detail-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}
.detail-label{font-size:12px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);margin-bottom:8px}
.detail-value{font-size:15px;font-weight:700;word-break:break-word}
.panel-card h3{margin:0 0 14px;font-size:18px}
.table-wrap{overflow:auto}table{width:100%;border-collapse:collapse}th,td{padding:12px;border-bottom:1px solid var(--border);text-align:left;vertical-align:top}th{color:#c9d8ee;font-size:13px;background:#101a2d;position:sticky;top:0}td{font-size:14px;color:#e9f1fd}
.entity-overview{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:16px}
.entity-card{box-shadow:0 8px 24px rgba(0,0,0,.22)}
.entity-card-header{display:flex;justify-content:space-between;gap:16px;align-items:flex-start}
.entity-card h3{margin:4px 0 6px;font-size:22px}.entity-card p{margin:0;color:var(--muted);font-size:13px}
.mini-list{list-style:none;margin:16px 0 0;padding:0;display:flex;flex-direction:column;gap:10px}
.mini-list li{display:flex;justify-content:space-between;gap:12px;align-items:center;padding:10px 12px;background:#101a2d;border:1px solid var(--border);border-radius:12px}
.mini-name{font-size:13px;font-weight:600}
.footer-note{margin-top:18px;font-size:12px;color:var(--muted)}
.nav-link{padding:5px 10px;border-radius:999px;display:inline-flex;align-items:center;gap:6px;font-size:12px;background:#0f2242;border-color:#244676}
.nav-link.inline{padding:0;border:none;background:none;color:#8ac0ff;border-radius:0;font-size:14px;text-decoration:underline}
.link-list{display:flex;gap:8px;flex-wrap:wrap}
.network-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}
.topology-shell{position:relative;border:1px solid var(--border);border-radius:18px;background:#0b1324;overflow:hidden}
.topology-toolbar{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:12px 14px;border-bottom:1px solid var(--border);background:#0f1a31;flex-wrap:wrap}
.topology-help{font-size:12px;color:var(--muted)}
.topology-actions{display:flex;gap:8px;flex-wrap:wrap}
.topology-btn{padding:8px 12px;border-radius:10px;background:#18253b;color:var(--text);border:1px solid var(--border);font-weight:600}
.topology-stage{width:100%;height:560px;display:block;background:radial-gradient(circle at top,#14233f,#0b1324 65%)}
.topology-node{cursor:pointer}
.topology-node text{pointer-events:none}
.topology-label{font-size:12px;fill:#e8eef8;font-weight:700}
.topology-sublabel{font-size:10px;fill:#9fb0c8;font-weight:600;text-transform:uppercase}
.topology-edge{stroke:#4f6589;stroke-width:2;opacity:.35;fill:none;transition:opacity .15s ease,stroke-width .15s ease,stroke .15s ease}
.topology-edge.active{opacity:1;stroke:#ffeb3b;stroke-width:4}
.topology-edge.faded{opacity:.10}
.topology-node{opacity:1;transition:opacity .15s ease,transform .15s ease}
.topology-node.active{opacity:1}
.topology-node.faded{opacity:.18}
.topology-node .topology-focus-ring{opacity:0;transition:opacity .15s ease}
.topology-node.active .topology-focus-ring{opacity:1}
.topology-legend{display:flex;gap:8px;flex-wrap:wrap;padding:12px 14px;border-top:1px solid var(--border);background:#0f1a31}
.legend-chip{display:inline-flex;align-items:center;gap:8px;padding:6px 10px;border-radius:999px;background:#14233f;border:1px solid var(--border);font-size:12px;color:#dbe7fb}
.legend-dot{width:10px;height:10px;border-radius:50%}
@media (max-width:1180px){.hero-grid,.main-grid,.stats,.detail-grid,.network-grid{grid-template-columns:1fr}.sidebar-card{position:static}}
</style>
</head>
<body>
<div class="container">
  <section class="hero">
    <div class="hero-top">
      <div>
        <h1>DMTools Inventory Report</h1>
        <p>Interactive summary of vCenter inventory relationships across virtual machines, hosts, clusters, datastores, and virtual networks.</p>
        <div class="hero-note">Generated: <span id="generatedAt"></span></div>
        <div class="hero-note">vCenter: <span id="vcenterName"></span> · API Version: <span id="apiVersion"></span></div>
      </div>
      <div class="badge large INFO">vSphere Estate</div>
    </div>
    <div class="hero-grid">
      <div class="stack-card">
        <div class="eyebrow">Summary Graphic</div>
        <h2 style="margin:6px 0 0">VM power state distribution</h2>
        <p class="section-subtitle" style="color:rgba(255,255,255,.74);margin-top:8px">Use relationship links inside each detail view to jump directly between objects.</p>
        <div class="stack-bar" id="powerBar"></div>
      </div>
      <div class="score-card">
        <div class="score-ring">
          <div class="score-ring-content">
            <div class="score-ring-value" id="poweredOnCount">0</div>
            <div class="score-ring-label">Powered On VMs</div>
          </div>
        </div>
      </div>
    </div>
  </section>

  <section class="stats">
    <div class="stat-card"><div class="stat-label">Total VMs</div><div class="stat-value" id="statVMs">0</div></div>
    <div class="stat-card"><div class="stat-label">Hosts</div><div class="stat-value" id="statHosts">0</div></div>
    <div class="stat-card"><div class="stat-label">Clusters</div><div class="stat-value" id="statClusters">0</div></div>
    <div class="stat-card"><div class="stat-label">Datastores</div><div class="stat-value" id="statDatastores">0</div></div>
    <div class="stat-card"><div class="stat-label">Networks</div><div class="stat-value" id="statNetworks">0</div></div>
    <div class="stat-card"><div class="stat-label">Templates</div><div class="stat-value" id="statTemplates">0</div></div>
  </section>

  <section class="main-grid">
    <aside class="sidebar">
      <div class="sidebar-card">
        <h2 class="section-title">Inventory Navigator</h2>
        <p class="section-subtitle">Select an entity type, filter the list, and open a single-object detail view.</p>
        <div class="entity-toggle">
          <button type="button" class="entity-btn active" data-entity="vms">VMs</button>
          <button type="button" class="entity-btn" data-entity="hosts">Hosts</button><button type="button" class="entity-btn" data-entity="templates">Templates</button>
          <button type="button" class="entity-btn" data-entity="clusters">Clusters</button>
          <button type="button" class="entity-btn" data-entity="datastores">Datastores</button>
          <button type="button" class="entity-btn" data-entity="networks">Networks</button>
        </div>
        <div class="filter-bar">
          <button type="button" class="filter-btn active" data-filter="ALL">All</button>
          <button type="button" class="filter-btn" data-filter="PASS">PASS</button>
          <button type="button" class="filter-btn" data-filter="INFO">INFO</button>
          <button type="button" class="filter-btn" data-filter="WARN">WARN</button>
        </div>
        <div class="item-nav" id="itemNav"></div>
      </div>
    </aside>

    <main class="content">
      <div class="view-mode-toggle">
        <button type="button" class="view-btn active" data-view="ALL">Show All</button>
        <button type="button" class="view-btn" data-view="DETAIL">Single Detail</button>
      </div>
      <div class="filter-bar" id="contextFilterBar"></div>

      <section class="panel active" id="overviewPanel">
        <div class="panel-card">
          <h3 id="overviewTitle">Virtual Machines</h3>
          <p class="section-subtitle" id="overviewSubtitle">Overview cards remain filtered by the selected entity type and status filter.</p>
        </div>
        <div class="entity-overview" id="overviewGrid"></div>
      </section>

      <section class="panel" id="detailPanel"></section>

      <div class="footer-note">This report is generated from the DMTools Excel workbook and adds cross-linked navigation between related objects.</div>
    </main>
  </section>
</div>

<script>
const reportData = __REPORT_DATA__;
const entityMeta = {
  vms:        { title: 'Virtual Machines', subtitle: 'VM ownership, placement, networking, storage, and VMware Tools context.' },
  templates:  { title: 'VM Templates', subtitle: 'Template placement, storage, network attachments, and reusable build metadata.' },
  hosts:      { title: 'Hosts', subtitle: 'Host platform details, child VMs, and host network configuration.' },
  clusters:   { title: 'Clusters', subtitle: 'Cluster settings with member hosts and virtual machines.' },
  datastores: { title: 'Datastores', subtitle: 'Datastore capacity and attached virtual machine references.' },
  networks:   { title: 'Virtual Networks', subtitle: 'Virtual network relationships across VMs, hosts, port groups, and VMkernel adapters.' }
};

let currentEntity = 'vms';
let currentFilter = 'ALL';
let currentContextFilter = 'ALL';
let currentView = 'ALL';
let currentSelectedId = null;

const el = {
  generatedAt: document.getElementById('generatedAt'),
  vcenterName: document.getElementById('vcenterName'),
  apiVersion: document.getElementById('apiVersion'),
  statVMs: document.getElementById('statVMs'),
  statHosts: document.getElementById('statHosts'),
  statClusters: document.getElementById('statClusters'),
  statDatastores: document.getElementById('statDatastores'),
  statNetworks: document.getElementById('statNetworks'),
  statTemplates: document.getElementById('statTemplates'),
  poweredOnCount: document.getElementById('poweredOnCount'),
  powerBar: document.getElementById('powerBar'),
  itemNav: document.getElementById('itemNav'),
  overviewPanel: document.getElementById('overviewPanel'),
  overviewGrid: document.getElementById('overviewGrid'),
  overviewTitle: document.getElementById('overviewTitle'),
  overviewSubtitle: document.getElementById('overviewSubtitle'),
  detailPanel: document.getElementById('detailPanel'),
  contextFilterBar: document.getElementById('contextFilterBar')
};

function safe(value) {
  return value === null || value === undefined || value === '' ? '-' : String(value);
}
function badgeClass(status) {
  return status || 'INFO';
}
function getEntityItems() {
  return reportData[currentEntity] || [];
}
function entityStatus(item) {
  if (currentEntity === 'vms') return item.statusClass || 'INFO';
  if (currentEntity === 'templates') return 'INFO';
  if (currentEntity === 'hosts') return (item.vmCount || 0) > 0 ? 'PASS' : 'INFO';
  if (currentEntity === 'clusters') return (item.vms || []).length > 0 ? 'PASS' : 'INFO';
  if (currentEntity === 'datastores') return (item.vmCount || 0) > 0 ? 'PASS' : 'INFO';
  if (currentEntity === 'networks') return (item.attachedVMs || []).length > 0 ? 'PASS' : 'INFO';
  return 'INFO';
}
function getContextFilters(entity) {
  switch (entity) {
    case 'vms':
      return [
        { key: 'ALL', label: 'All VMs' },
        { key: 'WINDOWS', label: 'Windows' },
        { key: 'LINUX', label: 'Linux' },
        { key: 'OTHEROS', label: 'Other OS' },
        { key: 'POWEREDON', label: 'Powered On' },
        { key: 'POWEROFF', label: 'Powered Off' },
        { key: 'HASPARTITIONS', label: 'Has Partitions' },
        { key: 'NOPARTITIONS', label: 'No Partitions' }
      ];
    case 'templates':
      return [
        { key: 'ALL', label: 'All Templates' },
        { key: 'WINDOWS', label: 'Windows' },
        { key: 'LINUX', label: 'Linux' },
        { key: 'OTHEROS', label: 'Other OS' }
      ];
    case 'hosts':
      return [
        { key: 'ALL', label: 'All Hosts' },
        { key: 'HASVMS', label: 'Has VMs' },
        { key: 'HASNETWORKS', label: 'Has Networks' },
        { key: 'HASVMK', label: 'Has VMkernel' }
      ];
    case 'clusters':
      return [
        { key: 'ALL', label: 'All Clusters' },
        { key: 'HAVMS', label: 'Has VMs' },
        { key: 'HAENABLED', label: 'HA Enabled' },
        { key: 'DRSENABLED', label: 'DRS Enabled' },
        { key: 'VSANENABLED', label: 'vSAN Enabled' }
      ];
    case 'datastores':
      return [
        { key: 'ALL', label: 'All Datastores' },
        { key: 'HASVMS', label: 'Has VMs' },
        { key: 'ACCESSIBLE', label: 'Accessible' },
        { key: 'MULTIHOST', label: 'Multi-Host' }
      ];
    case 'networks':
      return [
        { key: 'ALL', label: 'All Networks' },
        { key: 'HASVMS', label: 'Has VMs' },
        { key: 'HASHOSTS', label: 'Has Hosts' },
        { key: 'HASVMK', label: 'Has VMkernel' },
        { key: 'STDPORTGROUP', label: 'Standard PG' },
        { key: 'DVPORTGROUP', label: 'Distributed PG' }
      ];
    default:
      return [{ key: 'ALL', label: 'All' }];
  }
}
function getOsFamily(item) {
  const a = [item.osTools, item.osConfig].filter(Boolean).join(' ').toLowerCase();
  if (a.includes('windows')) return 'WINDOWS';
  if (a.includes('linux') || a.includes('ubuntu') || a.includes('debian') || a.includes('suse') || a.includes('centos') || a.includes('photon') || a.includes('red hat') || a.includes('rhel')) return 'LINUX';
  return 'OTHEROS';
}
function matchesContextFilter(item) {
  switch (currentEntity) {
    case 'vms':
      switch (currentContextFilter) {
        case 'WINDOWS': return getOsFamily(item) === 'WINDOWS';
        case 'LINUX': return getOsFamily(item) === 'LINUX';
        case 'OTHEROS': return getOsFamily(item) === 'OTHEROS';
        case 'POWEREDON': return safe(item.powerState) === 'PoweredOn';
        case 'POWEROFF': return safe(item.powerState) === 'PoweredOff';
        case 'HASPARTITIONS': return (item.partitions || []).length > 0;
        case 'NOPARTITIONS': return (item.partitions || []).length === 0;
        default: return true;
      }
    case 'templates':
      switch (currentContextFilter) {
        case 'WINDOWS': return getOsFamily(item) === 'WINDOWS';
        case 'LINUX': return getOsFamily(item) === 'LINUX';
        case 'OTHEROS': return getOsFamily(item) === 'OTHEROS';
        default: return true;
      }
    case 'hosts':
      switch (currentContextFilter) {
        case 'HASVMS': return (item.vmCount || 0) > 0;
        case 'HASNETWORKS': return (item.networks || []).length > 0;
        case 'HASVMK': return (item.vmkernels || []).length > 0;
        default: return true;
      }
    case 'clusters':
      switch (currentContextFilter) {
        case 'HAVMS': return (item.vms || []).length > 0;
        case 'HAENABLED': return String(item.haEnabled).toLowerCase() === 'true';
        case 'DRSENABLED': return String(item.drsEnabled).toLowerCase() === 'true';
        case 'VSANENABLED': return String(item.vsanEnabled).toLowerCase() === 'true';
        default: return true;
      }
    case 'datastores':
      switch (currentContextFilter) {
        case 'HASVMS': return (item.vmCount || 0) > 0;
        case 'ACCESSIBLE': return String(item.accessible).toLowerCase() === 'true';
        case 'MULTIHOST': return String(item.multiHost).toLowerCase() === 'true';
        default: return true;
      }
    case 'networks':
      switch (currentContextFilter) {
        case 'HASVMS': return (item.attachedVMs || []).length > 0;
        case 'HASHOSTS': return (item.attachedHosts || []).length > 0;
        case 'HASVMK': return (item.vmkernels || []).length > 0;
        case 'STDPORTGROUP': return (item.standardPortGroups || []).length > 0;
        case 'DVPORTGROUP': return (item.distributedPortGroups || []).length > 0;
        default: return true;
      }
    default:
      return true;
  }
}
function filteredItems() {
  const items = getEntityItems().slice().sort((a,b)=>safe(a.name).localeCompare(safe(b.name)));
  return items.filter(x => (currentFilter === 'ALL' || entityStatus(x) === currentFilter) && matchesContextFilter(x));
}
function renderContextFilters() {
  const filters = getContextFilters(currentEntity);
  el.contextFilterBar.innerHTML = '';
  filters.forEach(filter => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'context-filter-btn' + (filter.key === currentContextFilter ? ' active' : '');
    btn.textContent = filter.label;
    btn.addEventListener('click', () => {
      currentContextFilter = filter.key;
      currentSelectedId = null;
      refreshEntityView();
    });
    el.contextFilterBar.appendChild(btn);
  });
}
function refreshEntityView() {
  renderContextFilters();
  renderNav();
  renderPanels();
}
function setEntityButtonState() {
  document.querySelectorAll('.entity-btn').forEach(x => x.classList.toggle('active', x.dataset.entity === currentEntity));
}
function setFilterButtonState() {
  document.querySelectorAll('.filter-btn').forEach(x => x.classList.toggle('active', x.dataset.filter === currentFilter));
}
function setViewButtonState() {
  document.querySelectorAll('.view-btn').forEach(x => x.classList.toggle('active', x.dataset.view === currentView));
}
function setSummary() {
  el.generatedAt.textContent = safe(reportData.summary.generatedAt);
  el.vcenterName.textContent = safe(reportData.summary.vcenter);
  el.apiVersion.textContent = safe(reportData.summary.apiVersion);
  el.statVMs.textContent = safe(reportData.summary.totalVMs);
  el.statHosts.textContent = safe(reportData.summary.hosts);
  el.statClusters.textContent = safe(reportData.summary.clusters);
  el.statDatastores.textContent = safe(reportData.summary.datastores);
  el.statNetworks.textContent = safe(reportData.summary.networks);
  el.statTemplates.textContent = safe(reportData.summary.templates);
  el.poweredOnCount.textContent = safe(reportData.summary.poweredOnVMs);

  const total = Math.max(reportData.summary.totalVMs || 0, 1);
  const onCount = reportData.summary.poweredOnVMs || 0;
  const offCount = Math.max((reportData.summary.totalVMs || 0) - onCount, 0);
  const onPct = (onCount / total) * 100;
  const offPct = 100 - onPct;

  el.powerBar.innerHTML = `
    <div class="stack-segment PASS" style="width:${onPct}%"><span>Powered On ${onCount}</span></div>
    <div class="stack-segment INFO" style="width:${offPct}%"><span>Other ${offCount}</span></div>`;
}
function navSub(item) {
  switch (currentEntity) {
    case 'vms': return `${safe(item.host)} · ${safe(item.cluster)}`;
    case 'templates': return `${safe(item.datacenter)} · ${(item.datastores || []).length} datastores`;
    case 'hosts': return `${safe(item.cluster)} · ${item.vmCount || 0} VMs`;
    case 'clusters': return `${(item.hosts || []).length} hosts · ${(item.vms || []).length} VMs`;
    case 'datastores': return `${safe(item.type)} · ${item.vmCount || 0} VMs`;
    case 'networks': return `${(item.attachedHosts || []).length} hosts · ${(item.attachedVMs || []).length} VMs`;
    default: return '';
  }
}
function navLink(entity, label, idOverride) {
  if (!label || label === '-') return safe(label);
  const id = idOverride || slugify(label);
  return `<button type="button" class="nav-link inline" data-nav-entity="${entity}" data-nav-id="${id}">${safe(label)}</button>`;
}
function chipLink(entity, label, idOverride) {
  if (!label || label === '-') return '';
  const id = idOverride || slugify(label);
  return `<button type="button" class="nav-link" data-nav-entity="${entity}" data-nav-id="${id}">${safe(label)}</button>`;
}
function linkList(entity, values) {
  const arr = (values || []).filter(v => v);
  if (!arr.length) return '<span>-</span>';
  return `<div class="link-list">${arr.map(v => chipLink(entity, v)).join('')}</div>`;
}
function renderNav() {
  el.itemNav.innerHTML = '';
  filteredItems().forEach(item => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = currentSelectedId === item.id ? 'active' : '';
    btn.innerHTML = `
      <span class="item-nav-title">${safe(item.name)}</span>
      <span class="item-nav-meta">${currentEntity.slice(0,-1)} · ${entityStatus(item)}</span>
      <span class="item-nav-sub">${navSub(item)}</span>`;
    btn.addEventListener('click', () => {
      currentSelectedId = item.id;
      currentView = 'DETAIL';
      setViewButtonState();
      renderNav();
      renderPanels();
    });
    el.itemNav.appendChild(btn);
  });
}
function tableRows(values, colspan=5) {
  if (!values || !values.length) {
    return `<tr><td colspan="${colspan}">No related records.</td></tr>`;
  }
  return values.join('');
}
function slugify(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,'') || 'unknown';
}
function overviewMarkup(item) {
  const status = entityStatus(item);
  switch (currentEntity) {
    case 'vms':
      return `
      <div class="entity-card-header">
        <div>
          <div class="eyebrow">Virtual Machine</div>
          <h3>${safe(item.name)}</h3>
          <p>${safe(item.osConfig)} · Host: ${safe(item.host)} · Cluster: ${safe(item.cluster)}</p>
        </div>
        <div>
          <span class="badge large ${badgeClass(status)}">${safe(item.powerState)}</span><br/><br/>
          <button type="button" class="view-item-btn">Open details</button>
        </div>
      </div>
      <ul class="mini-list">
        <li><span class="mini-name">Primary IP</span><span>${safe(item.primaryIP)}</span></li>
        <li><span class="mini-name">Datastores</span><span>${(item.datastores || []).length}</span></li>
        <li><span class="mini-name">Networks</span><span>${(item.networks || []).length}</span></li>
        <li><span class="mini-name">Required Tools</span><span>${safe(item.requiredTools)}</span></li>
      </ul>`;
    case 'templates':
      return `
      <div class="entity-card-header">
        <div>
          <div class="eyebrow">VM Template</div>
          <h3>${safe(item.name)}</h3>
          <p>${safe(item.osConfig)} · Datacenter: ${safe(item.datacenter)} · Cluster: ${safe(item.cluster)}</p>
        </div>
        <div>
          <span class="badge large INFO">Template</span><br/><br/>
          <button type="button" class="view-item-btn">Open details</button>
        </div>
      </div>
      <ul class="mini-list">
        <li><span class="mini-name">Datastores</span><span>${(item.datastores || []).length}</span></li>
        <li><span class="mini-name">Networks</span><span>${(item.networks || []).length}</span></li>
        <li><span class="mini-name">Required Tools</span><span>${safe(item.requiredTools)}</span></li>
        <li><span class="mini-name">Config Status</span><span>${safe(item.configStatus)}</span></li>
      </ul>`;
    case 'hosts':
      return `
      <div class="entity-card-header">
        <div>
          <div class="eyebrow">Host</div>
          <h3>${safe(item.name)}</h3>
          <p>${safe(item.cluster)} · ${safe(item.version)} build ${safe(item.build)}</p>
        </div>
        <div>
          <span class="badge large ${badgeClass(status)}">${safe(item.vmCount)} VMs</span><br/><br/>
          <button type="button" class="view-item-btn">Open details</button>
        </div>
      </div>
      <ul class="mini-list">
        <li><span class="mini-name">Model</span><span>${safe(item.model)}</span></li>
        <li><span class="mini-name">Vendor</span><span>${safe(item.vendor)}</span></li>
        <li><span class="mini-name">pNICs</span><span>${(item.pnics || []).length}</span></li>
        <li><span class="mini-name">Virtual Networks</span><span>${(item.networks || []).length}</span></li>
      </ul>`;
    case 'clusters':
      return `
      <div class="entity-card-header">
        <div>
          <div class="eyebrow">Cluster</div>
          <h3>${safe(item.name)}</h3>
          <p>${safe(item.datacenter)}</p>
        </div>
        <div>
          <span class="badge large ${badgeClass(status)}">${(item.vms || []).length} VMs</span><br/><br/>
          <button type="button" class="view-item-btn">Open details</button>
        </div>
      </div>
      <ul class="mini-list">
        <li><span class="mini-name">Hosts</span><span>${(item.hosts || []).length}</span></li>
        <li><span class="mini-name">HA enabled</span><span>${safe(item.haEnabled)}</span></li>
        <li><span class="mini-name">DRS enabled</span><span>${safe(item.drsEnabled)}</span></li>
        <li><span class="mini-name">vSAN enabled</span><span>${safe(item.vsanEnabled)}</span></li>
      </ul>`;
    case 'datastores':
      return `
      <div class="entity-card-header">
        <div>
          <div class="eyebrow">Datastore</div>
          <h3>${safe(item.name)}</h3>
          <p>${safe(item.type)} · Accessible: ${safe(item.accessible)}</p>
        </div>
        <div>
          <span class="badge large ${badgeClass(status)}">${safe(item.vmCount)} VMs</span><br/><br/>
          <button type="button" class="view-item-btn">Open details</button>
        </div>
      </div>
      <ul class="mini-list">
        <li><span class="mini-name">Capacity MB</span><span>${safe(item.capacityMB)}</span></li>
        <li><span class="mini-name">Free MB</span><span>${safe(item.freeMB)}</span></li>
        <li><span class="mini-name">Provisioned MB</span><span>${safe(item.provisioned)}</span></li>
        <li><span class="mini-name">Attached VMs</span><span>${safe(item.vmCount)}</span></li>
      </ul>`;
    case 'networks':
      return `
      <div class="entity-card-header">
        <div>
          <div class="eyebrow">Virtual Network</div>
          <h3>${safe(item.name)}</h3>
          <p>Hosts: ${(item.attachedHosts || []).length} · VMs: ${(item.attachedVMs || []).length}</p>
        </div>
        <div>
          <span class="badge large ${badgeClass(status)}">${(item.attachedVMs || []).length} VMs</span><br/><br/>
          <button type="button" class="view-item-btn">Open details</button>
        </div>
      </div>
      <ul class="mini-list">
        <li><span class="mini-name">Standard Port Groups</span><span>${(item.standardPortGroups || []).length}</span></li>
        <li><span class="mini-name">Distributed Port Groups</span><span>${(item.distributedPortGroups || []).length}</span></li>
        <li><span class="mini-name">VMkernel Adapters</span><span>${(item.vmkernels || []).length}</span></li>
        <li><span class="mini-name">VLAN Hints</span><span>${(item.vlanHints || []).join(', ') || '-'}</span></li>
      </ul>`;
    default:
      return '';
  }
}
function renderOverview() {
  const meta = entityMeta[currentEntity];
  el.overviewTitle.textContent = meta.title;
  el.overviewSubtitle.textContent = meta.subtitle;
  el.overviewGrid.innerHTML = '';

  filteredItems().forEach(item => {
    const card = document.createElement('article');
    card.className = 'entity-card';
    card.innerHTML = overviewMarkup(item);
    el.overviewGrid.appendChild(card);
    const btn = card.querySelector('.view-item-btn');
    if (btn) {
      btn.addEventListener('click', () => {
        currentSelectedId = item.id;
        currentView = 'DETAIL';
        setViewButtonState();
        renderNav();
        renderPanels();
      });
    }
  });
}
function detailMarkup(item) {
  switch (currentEntity) {
    case 'vms':
      return `
      <div class="panel-header">
        <div>
          <div class="eyebrow">Virtual Machine</div>
          <h2>${safe(item.name)}</h2>
          <p class="panel-subtitle">${safe(item.osConfig)} · Host ${navLink('hosts', item.host)} · Cluster ${navLink('clusters', item.cluster)}</p>
        </div>
        <span class="badge large ${badgeClass(item.statusClass)}">${safe(item.powerState)}</span>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="detail-label">Datacenter</div><div class="detail-value">${safe(item.datacenter)}</div></div>
        <div class="detail-card"><div class="detail-label">DNS Name</div><div class="detail-value">${safe(item.dnsName)}</div></div>
        <div class="detail-card"><div class="detail-label">Primary IP</div><div class="detail-value">${safe(item.primaryIP)}</div></div>
        <div class="detail-card"><div class="detail-label">CPUs</div><div class="detail-value">${safe(item.cpus)}</div></div>
        <div class="detail-card"><div class="detail-label">Memory MB</div><div class="detail-value">${safe(item.memoryMB)}</div></div>
        <div class="detail-card"><div class="detail-label">Required Tools</div><div class="detail-value">${safe(item.requiredTools)}</div></div>
      </div>
      <div class="panel-card"><h3>Relationships</h3>
        <div class="detail-grid">
          <div class="detail-card"><div class="detail-label">Host</div><div class="detail-value">${navLink('hosts', item.host)}</div></div>
          <div class="detail-card"><div class="detail-label">Cluster</div><div class="detail-value">${navLink('clusters', item.cluster)}</div></div>
          <div class="detail-card"><div class="detail-label">Datastores</div><div class="detail-value">${linkList('datastores', item.datastores)}</div></div>
        </div>
      </div>
      <div class="panel-card"><h3>Networks</h3><div class="table-wrap"><table><thead><tr><th>Label</th><th>Network</th><th>MAC</th><th>Connected</th><th>IP</th></tr></thead><tbody>
      ${tableRows((item.networks || []).map(n => `<tr><td>${safe(n.label)}</td><td>${navLink('networks', n.network)}</td><td>${safe(n.macAddress)}</td><td>${safe(n.connected)}</td><td>${safe(n.ipAddress)}</td></tr>`))}
      </tbody></table></div></div>
      <div class="panel-card"><h3>Disks</h3><div class="table-wrap"><table><thead><tr><th>Label</th><th>Datastore</th><th>Disk Path</th><th>Thin</th><th>Persistence</th></tr></thead><tbody>
      ${tableRows((item.disks || []).map(d => `<tr><td>${safe(d.label)}</td><td>${navLink('datastores', d.datastore)}</td><td>${safe(d.diskPath)}</td><td>${safe(d.thin)}</td><td>${safe(d.persistence)}</td></tr>`))}
      </tbody></table></div></div>
      <div class="panel-card"><h3>Snapshots</h3><div class="table-wrap"><table><thead><tr><th>Name</th><th>Created</th><th>Size MB</th><th>Description</th></tr></thead><tbody>
      ${tableRows((item.snapshots || []).map(s => `<tr><td>${safe(s.name)}</td><td>${safe(s.created)}</td><td>${safe(s.sizeMB)}</td><td>${safe(s.description)}</td></tr>`),4)}
      </tbody></table></div></div>
      <div class="panel-card"><h3>VM File System</h3><div class="table-wrap"><table><thead><tr><th>Guest Path</th><th>VMDK</th><th>Datastore</th><th>Capacity MiB</th><th>Free MiB</th><th>Free %</th></tr></thead><tbody>
      ${tableRows((item.partitions || []).map(p => `<tr><td>${safe(p.partitionPath)}</td><td>${safe(p.vmdkLabel)}</td><td>${navLink('datastores', p.datastore)}</td><td>${safe(p.capacityMB)}</td><td>${safe(p.freeMB)}</td><td>${safe(p.freePct)}</td></tr>`),6)}
      </tbody></table></div></div>`;
    case 'templates':
      return `
      <div class="panel-header">
        <div>
          <div class="eyebrow">VM Template</div>
          <h2>${safe(item.name)}</h2>
          <p class="panel-subtitle">${safe(item.osConfig)} · Datacenter ${safe(item.datacenter)} · Cluster ${navLink('clusters', item.cluster)}</p>
        </div>
        <span class="badge large INFO">Template</span>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="detail-label">DNS Name</div><div class="detail-value">${safe(item.dnsName)}</div></div>
        <div class="detail-card"><div class="detail-label">Primary IP</div><div class="detail-value">${safe(item.primaryIP)}</div></div>
        <div class="detail-card"><div class="detail-label">Config Status</div><div class="detail-value">${safe(item.configStatus)}</div></div>
        <div class="detail-card"><div class="detail-label">CPUs</div><div class="detail-value">${safe(item.cpus)}</div></div>
        <div class="detail-card"><div class="detail-label">Memory MB</div><div class="detail-value">${safe(item.memoryMB)}</div></div>
        <div class="detail-card"><div class="detail-label">Required Tools</div><div class="detail-value">${safe(item.requiredTools)}</div></div>
      </div>
      <div class="panel-card"><h3>Relationships</h3>
        <div class="detail-grid">
          <div class="detail-card"><div class="detail-label">Cluster</div><div class="detail-value">${navLink('clusters', item.cluster)}</div></div>
          <div class="detail-card"><div class="detail-label">Datastores</div><div class="detail-value">${linkList('datastores', item.datastores)}</div></div>
          <div class="detail-card"><div class="detail-label">Networks</div><div class="detail-value">${linkList('networks', (item.networks || []).map(n => n.network).filter(Boolean))}</div></div>
        </div>
      </div>
      <div class="panel-card"><h3>Network Adapters</h3><div class="table-wrap"><table><thead><tr><th>Label</th><th>Network</th><th>MAC</th><th>Connected</th><th>IP</th></tr></thead><tbody>
      ${tableRows((item.networks || []).map(n => `<tr><td>${safe(n.label)}</td><td>${navLink('networks', n.network)}</td><td>${safe(n.macAddress)}</td><td>${safe(n.connected)}</td><td>${safe(n.ipAddress)}</td></tr>`))}
      </tbody></table></div></div>
      <div class="panel-card"><h3>Template Disks</h3><div class="table-wrap"><table><thead><tr><th>Label</th><th>Datastore</th><th>Disk Path</th><th>Thin</th><th>Persistence</th></tr></thead><tbody>
      ${tableRows((item.disks || []).map(d => `<tr><td>${safe(d.label)}</td><td>${navLink('datastores', d.datastore)}</td><td>${safe(d.diskPath)}</td><td>${safe(d.thin)}</td><td>${safe(d.persistence)}</td></tr>`))}
      </tbody></table></div></div>`;
    case 'hosts':
      return `
      <div class="panel-header">
        <div><div class="eyebrow">Host</div><h2>${safe(item.name)}</h2><p class="panel-subtitle">Cluster ${navLink('clusters', item.cluster)} · Datacenter ${safe(item.datacenter)}</p></div>
        <span class="badge large ${badgeClass(entityStatus(item))}">${safe(item.vmCount)} VMs</span>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="detail-label">Vendor</div><div class="detail-value">${safe(item.vendor)}</div></div>
        <div class="detail-card"><div class="detail-label">Model</div><div class="detail-value">${safe(item.model)}</div></div>
        <div class="detail-card"><div class="detail-label">CPU Model</div><div class="detail-value">${safe(item.cpuModel)}</div></div>
        <div class="detail-card"><div class="detail-label">Version</div><div class="detail-value">${safe(item.version)}</div></div>
        <div class="detail-card"><div class="detail-label">Build</div><div class="detail-value">${safe(item.build)}</div></div>
        <div class="detail-card"><div class="detail-label">Connection</div><div class="detail-value">${safe(item.connection)}</div></div>
      </div>
      <div class="panel-card"><h3>Child VMs</h3><div class="table-wrap"><table><thead><tr><th>VM</th><th>Power State</th><th>Primary IP</th><th>Networks</th></tr></thead><tbody>
      ${tableRows((item.vms || []).map(vmName => {
        const vm = [...(reportData.vms || []), ...(reportData.templates || [])].find(x => x.name === vmName);
        const networks = vm ? (vm.networks || []).map(n => n.network).filter(Boolean).join(', ') : '-';
        return `<tr><td>${navLink((reportData.templates || []).some(t => t.name === vmName) ? 'templates' : 'vms', vmName)}</td><td>${safe(vm ? vm.powerState : '')}</td><td>${safe(vm ? vm.primaryIP : '')}</td><td>${safe(networks)}</td></tr>`;
      }),4)}
      </tbody></table></div></div>
      <div class="panel-card"><h3>Physical NICs</h3><div class="table-wrap"><table><thead><tr><th>Device</th><th>MAC</th><th>Speed</th><th>Duplex</th><th>Switch</th></tr></thead><tbody>
      ${tableRows((item.pnics || []).map(p => `<tr><td>${safe(p.device)}</td><td>${safe(p.mac)}</td><td>${safe(p.linkSpeed)}</td><td>${safe(p.duplex)}</td><td>${safe(p.switch)}</td></tr>`))}
      </tbody></table></div></div>
      <div class="network-grid">
        <div class="panel-card"><h3>Standard vSwitches</h3><div class="table-wrap"><table><thead><tr><th>vSwitch</th><th>MTU</th><th>Uplinks</th><th>Active</th><th>Standby</th></tr></thead><tbody>
        ${tableRows((item.vswitches || []).map(s => `<tr><td>${safe(s.name)}</td><td>${safe(s.mtu)}</td><td>${safe(s.nic)}</td><td>${safe(s.activeNic)}</td><td>${safe(s.standbyNic)}</td></tr>`))}
        </tbody></table></div></div>
        <div class="panel-card"><h3>Connected Virtual Networks</h3><div class="detail-value">${linkList('networks', item.networks)}</div></div>
      </div>
      <div class="panel-card"><h3>Port Groups</h3><div class="table-wrap"><table><thead><tr><th>Port Group</th><th>vSwitch</th><th>VLAN</th><th>Active Ports</th></tr></thead><tbody>
      ${tableRows((item.portgroups || []).map(pg => `<tr><td>${navLink('networks', pg.portGroup)}</td><td>${safe(pg.vSwitch)}</td><td>${safe(pg.vlanId)}</td><td>${safe(pg.activePorts)}</td></tr>`),4)}
      </tbody></table></div></div>
      <div class="panel-card"><h3>VMkernel Adapters</h3><div class="table-wrap"><table><thead><tr><th>Adapter</th><th>IP</th><th>Port Group</th><th>vSwitch</th><th>Traffic</th></tr></thead><tbody>
      ${tableRows((item.vmkernels || []).map(vmk => {
        const traffic = [
          vmk.management ? 'Mgmt' : null,
          vmk.vMotion ? 'vMotion' : null,
          vmk.vSAN ? 'vSAN' : null,
          vmk.faultTol ? 'FT' : null,
          vmk.provisioning ? 'Provisioning' : null
        ].filter(Boolean).join(', ');
        return `<tr><td>${safe(vmk.adapter)}</td><td>${safe(vmk.ipAddress)}</td><td>${navLink('networks', vmk.portGroup)}</td><td>${safe(vmk.vSwitch)}</td><td>${safe(traffic)}</td></tr>`;
      }),5)}
      </tbody></table></div></div>`;
    case 'clusters':
      return `
      <div class="panel-header">
        <div><div class="eyebrow">Cluster</div><h2>${safe(item.name)}</h2><p class="panel-subtitle">${safe(item.datacenter)}</p></div>
        <span class="badge large ${badgeClass(entityStatus(item))}">${(item.vms || []).length} VMs</span>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="detail-label">HA Enabled</div><div class="detail-value">${safe(item.haEnabled)}</div></div>
        <div class="detail-card"><div class="detail-label">DRS Enabled</div><div class="detail-value">${safe(item.drsEnabled)}</div></div>
        <div class="detail-card"><div class="detail-label">vSAN Enabled</div><div class="detail-value">${safe(item.vsanEnabled)}</div></div>
        <div class="detail-card"><div class="detail-label">EVC Mode</div><div class="detail-value">${safe(item.evcMode)}</div></div>
        <div class="detail-card"><div class="detail-label">Host Count</div><div class="detail-value">${safe((item.hosts || []).length)}</div></div>
        <div class="detail-card"><div class="detail-label">VM Count</div><div class="detail-value">${safe((item.vms || []).length)}</div></div>
      </div>
      <div class="panel-card"><h3>Member Hosts</h3><div class="table-wrap"><table><thead><tr><th>Host</th></tr></thead><tbody>
      ${tableRows((item.hosts || []).map(h => `<tr><td>${navLink('hosts', h)}</td></tr>`),1)}
      </tbody></table></div></div>
      <div class="panel-card"><h3>Member VMs</h3><div class="table-wrap"><table><thead><tr><th>VM</th></tr></thead><tbody>
      ${tableRows((item.vms || []).map(v => `<tr><td>${navLink((reportData.templates || []).some(t => t.name === v) ? 'templates' : 'vms', v)}</td></tr>`),1)}
      </tbody></table></div></div>`;
    case 'datastores':
      return `
      <div class="panel-header">
        <div><div class="eyebrow">Datastore</div><h2>${safe(item.name)}</h2><p class="panel-subtitle">${safe(item.type)} · Accessible ${safe(item.accessible)}</p></div>
        <span class="badge large ${badgeClass(entityStatus(item))}">${safe(item.vmCount)} VMs</span>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="detail-label">Capacity MB</div><div class="detail-value">${safe(item.capacityMB)}</div></div>
        <div class="detail-card"><div class="detail-label">Free MB</div><div class="detail-value">${safe(item.freeMB)}</div></div>
        <div class="detail-card"><div class="detail-label">Provisioned MB</div><div class="detail-value">${safe(item.provisioned)}</div></div>
        <div class="detail-card"><div class="detail-label">Accessible</div><div class="detail-value">${safe(item.accessible)}</div></div>
        <div class="detail-card"><div class="detail-label">MHA</div><div class="detail-value">${safe(item.multiHost)}</div></div>
        <div class="detail-card"><div class="detail-label">Attached VMs</div><div class="detail-value">${safe(item.vmCount)}</div></div>
      </div>
      <div class="panel-card"><h3>Attached VMs</h3><div class="table-wrap"><table><thead><tr><th>VM</th><th>Host</th><th>Cluster</th></tr></thead><tbody>
      ${tableRows((item.vms || []).map(vmName => {
        const vm = [...(reportData.vms || []), ...(reportData.templates || [])].find(x => x.name === vmName);
        return `<tr><td>${navLink((reportData.templates || []).some(t => t.name === vmName) ? 'templates' : 'vms', vmName)}</td><td>${navLink('hosts', vm ? vm.host : '')}</td><td>${navLink('clusters', vm ? vm.cluster : '')}</td></tr>`;
      }),3)}
      </tbody></table></div></div>`;
    case 'networks':
      return `
      <div class="panel-header">
        <div><div class="eyebrow">Virtual Network</div><h2>${safe(item.name)}</h2><p class="panel-subtitle">Hosts ${linkList('hosts', item.attachedHosts)} · VMs ${linkList('vms', item.attachedVMs)}</p></div>
        <span class="badge large ${badgeClass(entityStatus(item))}">${(item.attachedVMs || []).length} VMs</span>
      </div>
      <div class="detail-grid">
        <div class="detail-card"><div class="detail-label">Attached Hosts</div><div class="detail-value">${safe((item.attachedHosts || []).length)}</div></div>
        <div class="detail-card"><div class="detail-label">Attached VMs</div><div class="detail-value">${safe((item.attachedVMs || []).length)}</div></div>
        <div class="detail-card"><div class="detail-label">VLAN Hints</div><div class="detail-value">${safe((item.vlanHints || []).join(', '))}</div></div>
        <div class="detail-card"><div class="detail-label">Standard Switches</div><div class="detail-value">${safe((item.standardSwitches || []).join(', '))}</div></div>
        <div class="detail-card"><div class="detail-label">Distributed Switches</div><div class="detail-value">${safe((item.distributedSwitches || []).join(', '))}</div></div>
        <div class="detail-card"><div class="detail-label">VMkernel Adapters</div><div class="detail-value">${safe((item.vmkernels || []).length)}</div></div>
      </div>
      <div class="panel-card"><h3>Attached VMs</h3><div class="table-wrap"><table><thead><tr><th>VM</th><th>Host</th><th>MAC</th><th>Connected</th><th>IP</th></tr></thead><tbody>
      ${tableRows((item.vmAttachments || []).map(v => {
        const vm = [...(reportData.vms || []), ...(reportData.templates || [])].find(x => x.name === v.vm);
        return `<tr><td>${navLink((reportData.templates || []).some(t => t.name === v.vm) ? 'templates' : 'vms', v.vm)}</td><td>${navLink('hosts', vm ? vm.host : '')}</td><td>${safe(v.macAddress)}</td><td>${safe(v.connected)}</td><td>${safe(v.ipAddress)}</td></tr>`;
      }),5)}
      </tbody></table></div></div>
      <div class="network-grid">
        <div class="panel-card"><h3>Standard Port Groups</h3><div class="table-wrap"><table><thead><tr><th>Host</th><th>vSwitch</th><th>Port Group</th><th>VLAN</th></tr></thead><tbody>
        ${tableRows((item.standardPortGroups || []).map(pg => `<tr><td>${navLink('hosts', pg.host)}</td><td>${safe(pg.vSwitch)}</td><td>${safe(pg.portGroup)}</td><td>${safe(pg.vlanId)}</td></tr>`),4)}
        </tbody></table></div></div>
        <div class="panel-card"><h3>Distributed Port Groups</h3><div class="table-wrap"><table><thead><tr><th>Name</th><th>dvSwitch</th><th>VLAN</th><th>Type</th></tr></thead><tbody>
        ${tableRows((item.distributedPortGroups || []).map(pg => `<tr><td>${safe(pg.name)}</td><td>${safe(pg.vdSwitch)}</td><td>${safe(pg.vlanId)}</td><td>${safe(pg.type)}</td></tr>`),4)}
        </tbody></table></div></div>
      </div>
      <div class="panel-card"><h3>VMkernel Adapters</h3><div class="table-wrap"><table><thead><tr><th>Host</th><th>Adapter</th><th>IP</th><th>vSwitch</th><th>Traffic</th></tr></thead><tbody>
      ${tableRows((item.vmkernels || []).map(vmk => {
        const traffic = [
          vmk.management ? 'Mgmt' : null,
          vmk.vMotion ? 'vMotion' : null,
          vmk.vSAN ? 'vSAN' : null,
          vmk.faultTol ? 'FT' : null,
          vmk.provisioning ? 'Provisioning' : null
        ].filter(Boolean).join(', ');
        return `<tr><td>${navLink('hosts', vmk.host)}</td><td>${safe(vmk.adapter)}</td><td>${safe(vmk.ipAddress)}</td><td>${safe(vmk.vSwitch)}</td><td>${safe(traffic)}</td></tr>`;
      }),5)}
      </tbody></table></div></div>`;
    default:
      return '';
  }
}
const topologyExpandedGroups = {};

function topologyShell() {
  return `
  <div class="panel-card">
    <h3>Object Connection Map</h3>
    <div class="topology-shell">
      <div class="topology-toolbar">
        <div class="topology-help">A hierarchical topology map of the selected object and its related cluster, host, switch, network, datastore, and workload objects. Click grouped workload nodes to expand or collapse them, drag to pan, use the mouse wheel to zoom, or click object nodes to jump to their detail view.</div>
        <div class="topology-actions">
          <button type="button" class="topology-btn" id="topologyZoomIn">Zoom In</button>
          <button type="button" class="topology-btn" id="topologyZoomOut">Zoom Out</button>
          <button type="button" class="topology-btn" id="topologyReset">Reset</button>
        </div>
      </div>
      <svg id="topologySvg" class="topology-stage" viewBox="0 0 1280 760" preserveAspectRatio="xMidYMid meet"></svg>
      <div class="topology-legend">
        <span class="legend-chip"><span class="legend-dot" style="background:#4fc3f7"></span>Selected</span>
        <span class="legend-chip"><span class="legend-dot" style="background:#66bb6a"></span>VM / Template</span>
        <span class="legend-chip"><span class="legend-dot" style="background:#ffb74d"></span>Host</span>
        <span class="legend-chip"><span class="legend-dot" style="background:#ba68c8"></span>Cluster</span>
        <span class="legend-chip"><span class="legend-dot" style="background:#90a4ae"></span>Datastore</span>
        <span class="legend-chip"><span class="legend-dot" style="background:#90caf9"></span>Network / Switch</span>
        <span class="legend-chip"><span class="legend-dot" style="background:#26a69a"></span>Expandable Workload Group</span>
      </div>
    </div>
  </div>`;
}
function getCurrentDetailItem() {
  const items = filteredItems();
  return items.find(x => x.id === currentSelectedId) || items[0] || null;
}
function getGroupExpansionKey(expandId) {
  return `${currentEntity}|${currentSelectedId || 'none'}|${expandId}`;
}
function isGroupExpanded(expandId) {
  return !!topologyExpandedGroups[getGroupExpansionKey(expandId)];
}
function toggleTopologyGroup(expandId) {
  const key = getGroupExpansionKey(expandId);
  topologyExpandedGroups[key] = !topologyExpandedGroups[key];
  const item = getCurrentDetailItem();
  if (item) {
    renderTopology(item, currentEntity);
  }
}
function renderDetail() {
  const items = filteredItems();
  const item = items.find(x => x.id === currentSelectedId) || items[0];
  if (!item) {
    el.detailPanel.innerHTML = '';
    return;
  }
  currentSelectedId = item.id;
  el.detailPanel.innerHTML = detailMarkup(item) + topologyShell();
  renderTopology(item, currentEntity);
}

function addNode(map, entity, name, label, kind, options = {}) {
  if (!name) return null;
  const id = options.id || `${entity}:${name}`;
  if (!map.has(id)) {
    map.set(id, {
      id,
      entity,
      name,
      label: label || name,
      kind,
      emphasis: !!options.emphasis,
      meta: options.meta || '',
      layerHint: options.layerHint ?? null,
      fixedX: options.fixedX,
      fixedY: options.fixedY,
      externalUrl: options.externalUrl || null,
      expandId: options.expandId || null,
      count: options.count ?? null,
      fanGroupId: options.fanGroupId || null,
      fanIndex: options.fanIndex ?? null,
      fanCount: options.fanCount ?? null
    });
  }
  return id;
}
function addEdge(edges, from, to, label = '') {
  if (!from || !to || from === to) return;
  if (!edges.some(e => (e.from === from && e.to === to && e.label === label) || (e.from === to && e.to === from && e.label === label))) {
    edges.push({ from, to, label });
  }
}
function getVmEntity(vmName) {
  return (reportData.templates || []).some(t => t.name === vmName) ? 'templates' : 'vms';
}
function getVmObject(vmName) {
  return [...(reportData.vms || []), ...(reportData.templates || [])].find(v => v.name === vmName);
}
function getNetworkObject(name) {
  return (reportData.networks || []).find(n => n.name === name);
}
function getHostObject(name) {
  return (reportData.hosts || []).find(h => h.name === name);
}
function getClusterObject(name) {
  return (reportData.clusters || []).find(c => c.name === name);
}
function addWorkloadGroup(nodes, edges, parentId, parentName, vmNames, options = {}) {
  const children = (vmNames || []).filter(Boolean);
  if (!children.length) return null;

  const groupLabel = options.label || 'Virtual Machines';
  const expandId = options.expandId || `${parentName}:workloads`;
  const expanded = isGroupExpanded(expandId);
  const groupMeta = `${children.length} object${children.length === 1 ? '' : 's'} · ${expanded ? 'click to collapse' : 'click to expand'}`;
  const groupId = addNode(nodes, 'none', `${expandId}:group`, groupLabel, 'group', {
    id: `group:${expandId}`,
    meta: groupMeta,
    layerHint: options.layerHint ?? 3,
    expandId,
    count: children.length
  });
  addEdge(edges, parentId, groupId, options.edgeLabel || 'contains');

  if (expanded) {
    children.forEach((vmName, idx) => {
      const vmObj = getVmObject(vmName);
      const vmEntity = getVmEntity(vmName);
      const vmId = addNode(nodes, vmEntity, vmName, vmName, vmEntity === 'templates' ? 'template' : 'vm', {
        meta: vmEntity === 'templates' ? 'Template' : 'Guest',
        layerHint: (options.layerHint ?? 3) + 1,
        fanGroupId: groupId,
        fanIndex: idx,
        fanCount: children.length
      });
      addEdge(edges, groupId, vmId, vmEntity === 'templates' ? 'holds' : 'runs');
      if (options.includeStorage && vmObj) {
        (vmObj.datastores || []).slice(0, 4).forEach(ds => {
          const dsId = addNode(nodes, 'datastores', ds, ds, 'datastore', { meta: 'Datastore', layerHint: (options.layerHint ?? 3) + 2 });
          addEdge(edges, dsId, vmId, 'stores');
        });
      }
    });
  }
  return groupId;
}

function addFilesystemGroup(nodes, edges, vmId, vmObj, options = {}) {
  const children = (vmObj && vmObj.partitions ? vmObj.partitions : []).filter(p => p && p.partitionPath);
  if (!children.length) { return null; }

  const groupLabel = options.label || 'VmFileSystem';
  const expandId = options.expandId || `${vmObj.name}:filesystem`;
  const expanded = isGroupExpanded(expandId);
  const groupMeta = `${children.length} object${children.length === 1 ? '' : 's'} · ${expanded ? 'click to collapse' : 'click to expand'}`;
  const groupId = addNode(nodes, 'none', `${expandId}:group`, groupLabel, 'group', {
    id: `group:${expandId}`,
    meta: groupMeta,
    layerHint: options.layerHint ?? 3,
    expandId,
    count: children.length
  });
  addEdge(edges, vmId, groupId, '');

  if (expanded) {
    children.forEach((part, idx) => {
      const label = part.partitionPath;
      const meta = part.vmdkLabel ? part.vmdkLabel : 'File System';
      const pId = addNode(nodes, 'none', `${vmObj.name}:partition:${label}`, label, 'filesystem', {
        id: `filesystem:${vmObj.name}:${label}`,
        meta,
        layerHint: (options.layerHint ?? 3) + 1,
        fanGroupId: groupId,
        fanIndex: idx,
        fanCount: children.length
      });
      addEdge(edges, groupId, pId, '');

      if (part.datastore) {
        const dsId = addNode(nodes, 'datastores', part.datastore, part.datastore, 'datastore', {
          meta: 'Datastore',
          layerHint: (options.layerHint ?? 3) + 2
        });
        addEdge(edges, dsId, pId, '');
      }
      else {
        const warningId = addNode(nodes, 'none', `${vmObj.name}:disk-enableuuid-warning`, 'Warning: additional VM setting required', 'warning', {
          id: `warning:${vmObj.name}:disk-enableuuid`,
          meta: 'disk.EnableUUID · opens KB',
          layerHint: (options.layerHint ?? 3) + 2,
          externalUrl: 'https://knowledge.broadcom.com/external/article/432156/mapping-disk-in-windows-to-its-correspon.html'
        });
        addEdge(edges, warningId, pId, '');
      }
    });
  }
  return groupId;
}

function buildTopologyGraph(item, entity) {
  const nodes = new Map();
  const edges = [];

  const addVmHierarchy = (vmObj, vmEntity, selected = false) => {
    const vmId = addNode(nodes, vmEntity, vmObj.name, vmObj.name, vmEntity === 'templates' ? 'template' : 'vm', {
      emphasis: selected,
      meta: vmEntity === 'templates' ? 'Template' : 'Guest',
      layerHint: 2
    });

    const clusterObj = vmObj.cluster ? getClusterObject(vmObj.cluster) : null;
    const hostObj = vmObj.host ? getHostObject(vmObj.host) : null;
    const clusterId = vmObj.cluster ? addNode(nodes, 'clusters', vmObj.cluster, vmObj.cluster, 'cluster', { meta: 'Cluster', layerHint: 0 }) : null;
    const hostId = vmObj.host ? addNode(nodes, 'hosts', vmObj.host, vmObj.host, 'host', { meta: 'ESXi Host', layerHint: 1 }) : null;

    if (clusterId && hostId) addEdge(edges, clusterId, hostId, 'contains');
    if (hostId) addEdge(edges, hostId, vmId, vmEntity === 'templates' ? 'holds' : 'runs');

    (vmObj.datastores || []).forEach(ds => {
      const dsId = addNode(nodes, 'datastores', ds, ds, 'datastore', { meta: 'Datastore', layerHint: 4 });
      addEdge(edges, dsId, vmId, '');
    });

    addFilesystemGroup(nodes, edges, vmId, vmObj, { label: 'VmFileSystem', expandId: `${vmObj.name}:filesystem`, layerHint: 3 });

    (vmObj.networks || []).forEach(n => {
      if (!n || !n.network) return;
      const networkObj = getNetworkObject(n.network);
      const netId = addNode(nodes, 'networks', n.network, n.network, 'network', { meta: 'Port Group', layerHint: 4 });
      const switchNames = [
        ...((networkObj?.standardSwitches) || []).map(s => ({ name: s, type: 'vSwitch' })),
        ...((networkObj?.distributedSwitches) || []).map(s => ({ name: s, type: 'dvSwitch' }))
      ];
      if (switchNames.length) {
        switchNames.forEach(sw => {
          const swId = addNode(nodes, 'none', `${n.network}:${sw.type}:${sw.name}`, sw.name, 'switch', {
            id: `switch:${n.network}:${sw.type}:${sw.name}`,
            meta: sw.type,
            layerHint: 3
          });
          if (hostId) addEdge(edges, hostId, swId, 'uplinks');
          addEdge(edges, swId, netId, 'presents');
        });
      } else if (hostId) {
        addEdge(edges, hostId, netId, 'presents');
      }
      addEdge(edges, netId, vmId, 'connected');
    });

    return vmId;
  };

  if (entity === 'vms' || entity === 'templates') {
    addVmHierarchy(item, entity, true);
  }
  else if (entity === 'hosts') {
    const hostId = addNode(nodes, 'hosts', item.name, item.name, 'host', { emphasis: true, meta: 'ESXi Host', layerHint: 1 });
    if (item.cluster) {
      const clusterId = addNode(nodes, 'clusters', item.cluster, item.cluster, 'cluster', { meta: 'Cluster', layerHint: 0 });
      addEdge(edges, clusterId, hostId, 'contains');
    }
    (item.vswitches || []).forEach(sw => {
      const swId = addNode(nodes, 'none', `${item.name}:vSwitch:${sw.name}`, sw.name, 'switch', {
        id: `switch:${item.name}:vSwitch:${sw.name}`,
        meta: 'vSwitch',
        layerHint: 2
      });
      addEdge(edges, hostId, swId, 'uplinks');
    });
    (item.networks || []).slice(0, 12).forEach(netName => {
      const netId = addNode(nodes, 'networks', netName, netName, 'network', { meta: 'Port Group', layerHint: 3 });
      const relatedSwitches = [...new Set((item.portgroups || []).filter(pg => pg.portGroup === netName).map(pg => pg.vSwitch).filter(Boolean))];
      if (relatedSwitches.length) {
        relatedSwitches.forEach(swName => {
          const swId = addNode(nodes, 'none', `${item.name}:vSwitch:${swName}`, swName, 'switch', {
            id: `switch:${item.name}:vSwitch:${swName}`,
            meta: 'vSwitch',
            layerHint: 2
          });
          addEdge(edges, hostId, swId, 'uplinks');
          addEdge(edges, swId, netId, 'presents');
        });
      } else {
        addEdge(edges, hostId, netId, 'presents');
      }
    });
    addWorkloadGroup(nodes, edges, hostId, item.name, item.vms || [], {
      label: 'Virtual Machines',
      expandId: `host:${item.name}:vms`,
      edgeLabel: 'runs',
      layerHint: 3,
      includeStorage: false
    });
  }
  else if (entity === 'clusters') {
    const clusterId = addNode(nodes, 'clusters', item.name, item.name, 'cluster', { emphasis: true, meta: 'Cluster', layerHint: 0 });
    (item.hosts || []).forEach(hostName => {
      const hostId = addNode(nodes, 'hosts', hostName, hostName, 'host', { meta: 'ESXi Host', layerHint: 1 });
      addEdge(edges, clusterId, hostId, 'contains');
      const hostObj = getHostObject(hostName);
      if (hostObj) {
        addWorkloadGroup(nodes, edges, hostId, hostName, hostObj.vms || [], {
          label: 'Virtual Machines',
          expandId: `cluster:${item.name}:host:${hostName}:vms`,
          edgeLabel: 'runs',
          layerHint: 2,
          includeStorage: false
        });
      }
    });
  }
  else if (entity === 'datastores') {
    const dsId = addNode(nodes, 'datastores', item.name, item.name, 'datastore', { emphasis: true, meta: 'Datastore', layerHint: 2 });
    const vmNames = item.vms || [];
    const hosts = [...new Set(vmNames.map(vmName => getVmObject(vmName)?.host).filter(Boolean))];
    const clusters = [...new Set(vmNames.map(vmName => getVmObject(vmName)?.cluster).filter(Boolean))];
    clusters.forEach(clusterName => {
      const clusterId = addNode(nodes, 'clusters', clusterName, clusterName, 'cluster', { meta: 'Cluster', layerHint: 0 });
      hosts.filter(h => getVmObject(vmNames.find(v => (getVmObject(v)?.host === h && getVmObject(v)?.cluster === clusterName))) )
        .forEach(hostName => {
          const hostId = addNode(nodes, 'hosts', hostName, hostName, 'host', { meta: 'ESXi Host', layerHint: 1 });
          addEdge(edges, clusterId, hostId, 'contains');
        });
    });
    addWorkloadGroup(nodes, edges, dsId, item.name, vmNames, {
      label: 'Stored VMs',
      expandId: `datastore:${item.name}:vms`,
      edgeLabel: 'stores',
      layerHint: 3,
      includeStorage: false
    });
  }
  else if (entity === 'networks') {
    const netId = addNode(nodes, 'networks', item.name, item.name, 'network', { emphasis: true, meta: 'Port Group', layerHint: 3 });
    const stdSwitches = item.standardSwitches || [];
    const dvSwitches = item.distributedSwitches || [];
    (item.attachedHosts || []).forEach(hostName => {
      const hostObj = getHostObject(hostName);
      const hostId = addNode(nodes, 'hosts', hostName, hostName, 'host', { meta: 'ESXi Host', layerHint: 1 });
      if (hostObj?.cluster) {
        const clusterId = addNode(nodes, 'clusters', hostObj.cluster, hostObj.cluster, 'cluster', { meta: 'Cluster', layerHint: 0 });
        addEdge(edges, clusterId, hostId, 'contains');
      }
      if (stdSwitches.length) {
        stdSwitches.forEach(swName => {
          const swId = addNode(nodes, 'none', `${hostName}:vSwitch:${swName}`, swName, 'switch', {
            id: `switch:${hostName}:vSwitch:${swName}`,
            meta: 'vSwitch',
            layerHint: 2
          });
          addEdge(edges, hostId, swId, 'uplinks');
          addEdge(edges, swId, netId, 'presents');
        });
      } else {
        addEdge(edges, hostId, netId, 'presents');
      }
    });
    dvSwitches.forEach(swName => {
      const swId = addNode(nodes, 'none', `dvSwitch:${swName}`, swName, 'switch', {
        id: `switch:dvSwitch:${swName}`,
        meta: 'dvSwitch',
        layerHint: 2
      });
      addEdge(edges, swId, netId, 'presents');
    });
    addWorkloadGroup(nodes, edges, netId, item.name, item.attachedVMs || [], {
      label: 'Connected VMs',
      expandId: `network:${item.name}:vms`,
      edgeLabel: 'connected',
      layerHint: 4,
      includeStorage: false
    });
  }

  return { nodes: [...nodes.values()], edges };
}
function nodeStyle(kind, emphasis) {
  const palette = {
    vm:        { fill: '#43a047', accent: '#1b5e20' },
    template:  { fill: '#7cb342', accent: '#33691e' },
    host:      { fill: '#ffb74d', accent: '#ef6c00' },
    cluster:   { fill: '#ba68c8', accent: '#7b1fa2' },
    datastore: { fill: '#90a4ae', accent: '#546e7a' },
    network:   { fill: '#90caf9', accent: '#1565c0' },
    filesystem:{ fill: '#ce93d8', accent: '#8e24aa' },
    warning:   { fill: '#ff8a65', accent: '#d84315' },
    switch:    { fill: '#80cbc4', accent: '#00695c' },
    group:     { fill: '#26a69a', accent: '#00695c' },
    default:   { fill: '#90a4ae', accent: '#546e7a' }
  };
  const style = palette[kind] || palette.default;
  return { ...style, emphasis };
}
function getLayerConfig(entity) {
  switch (entity) {
    case 'vms':
    case 'templates':
      return { cluster: 0, host: 1, vm: 2, template: 2, switch: 3, group: 3, filesystem: 4, network: 4, datastore: 4, warning: 4 };
    case 'hosts':
      return { cluster: 0, host: 1, switch: 2, network: 3, group: 3, vm: 4, template: 4, datastore: 5 };
    case 'clusters':
      return { cluster: 0, host: 1, group: 2, vm: 3, template: 3, datastore: 4, network: 4, switch: 2 };
    case 'datastores':
      return { cluster: 0, host: 1, datastore: 2, group: 3, vm: 4, template: 4, network: 5, switch: 2 };
    case 'networks':
      return { cluster: 0, host: 1, switch: 2, network: 3, group: 4, vm: 5, template: 5, datastore: 6 };
    default:
      return { cluster: 0, host: 1, vm: 2, template: 2, switch: 3, group: 3, filesystem: 4, network: 4, datastore: 4, warning: 4 };
  }
}
function layoutTopology(nodes, entity) {
  const config = getLayerConfig(entity);
  const width = 1280;
  const height = 760;
  const padX = 120;
  const padY = 110;

  const layers = {};
  nodes.forEach(node => {
    const layer = node.layerHint ?? config[node.kind] ?? 2;
    node.layer = layer;
    if (!layers[layer]) layers[layer] = [];
    layers[layer].push(node);
  });

  const layerKeys = Object.keys(layers).map(Number).sort((a, b) => a - b);
  const maxLayer = Math.max(...layerKeys, 1);

  layerKeys.forEach(layer => {
    const layerNodes = layers[layer].slice().sort((a, b) => {
      if (a.emphasis && !b.emphasis) return 1;
      if (!a.emphasis && b.emphasis) return -1;
      return String(a.label).localeCompare(String(b.label));
    });

    const x = padX + ((width - (padX * 2)) * layer / Math.max(maxLayer, 1));
    const count = layerNodes.length;
    const spreadHeight = Math.max(200, height - (padY * 2));
    const step = count <= 1 ? 0 : Math.min(125, spreadHeight / Math.max(count - 1, 1));
    const totalSpan = step * Math.max(count - 1, 0);
    const startY = (height / 2) - (totalSpan / 2);

    layerNodes.forEach((node, idx) => {
      node.x = node.fixedX ?? x;
      node.y = node.fixedY ?? (count === 1 ? height / 2 : startY + (idx * step));
    });

    const emphasisNode = layerNodes.find(n => n.emphasis);
    if (emphasisNode) {
      emphasisNode.y = height / 2;
      const siblings = layerNodes.filter(n => n !== emphasisNode);
      const upper = siblings.filter((_, i) => i % 2 === 0);
      const lower = siblings.filter((_, i) => i % 2 === 1);
      upper.forEach((node, i) => node.y = (height / 2) - 140 - (i * 90));
      lower.forEach((node, i) => node.y = (height / 2) + 140 + (i * 90));
    }
  });
}
function applyGroupFanout(nodes) {
  const groups = nodes.filter(n => n.kind === 'group');
  groups.forEach(group => {
    const children = nodes
      .filter(n => n.fanGroupId === group.id)
      .sort((a, b) => String(a.label).localeCompare(String(b.label)));
    const count = children.length;
    if (!count) return;

    if (count <= 8) {
      const angleStart = -72;
      const angleEnd = 72;
      const radius = 210;
      children.forEach((child, idx) => {
        const angle = count === 1 ? 0 : angleStart + ((angleEnd - angleStart) * idx / Math.max(count - 1, 1));
        const radians = angle * (Math.PI / 180);
        child.x = group.x + Math.cos(radians) * radius;
        child.y = group.y + Math.sin(radians) * (radius * 0.92);
      });
      return;
    }

    const maxRows = 8;
    const cols = Math.min(4, Math.max(2, Math.ceil(count / maxRows)));
    const rows = Math.ceil(count / cols);
    const colSpacing = 168;
    const rowSpacing = rows > 7 ? 72 : 84;
    const startX = group.x + 170;
    const startY = group.y - ((rows - 1) * rowSpacing / 2);

    children.forEach((child, idx) => {
      const col = Math.floor(idx / rows);
      const row = idx % rows;
      child.x = startX + (col * colSpacing);
      child.y = startY + (row * rowSpacing);
    });
  });
}
function drawVmIcon(group, node, style, isTemplate = false) {
  const body = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  body.setAttribute('x', -34); body.setAttribute('y', -24);
  body.setAttribute('width', 68); body.setAttribute('height', 44);
  body.setAttribute('rx', 10); body.setAttribute('fill', style.fill);
  body.setAttribute('stroke', '#d6e4ff'); body.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(body);

  const header = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  header.setAttribute('x', -34); header.setAttribute('y', -24);
  header.setAttribute('width', 68); header.setAttribute('height', 10);
  header.setAttribute('rx', 10); header.setAttribute('fill', style.accent);
  group.appendChild(header);

  for (let i = 0; i < 3; i++) {
    const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    dot.setAttribute('cx', -24 + (i * 8)); dot.setAttribute('cy', -19); dot.setAttribute('r', 1.8);
    dot.setAttribute('fill', '#d6e4ff'); group.appendChild(dot);
  }

  if (isTemplate) {
    const fold = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    fold.setAttribute('d', 'M 18 -24 L 34 -8 L 18 -8 Z');
    fold.setAttribute('fill', '#ffffffaa');
    group.appendChild(fold);
  }
}
function drawHostIcon(group, node, style) {
  const chassis = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  chassis.setAttribute('x', -26); chassis.setAttribute('y', -28);
  chassis.setAttribute('width', 52); chassis.setAttribute('height', 56);
  chassis.setAttribute('rx', 8); chassis.setAttribute('fill', style.fill);
  chassis.setAttribute('stroke', '#d6e4ff'); chassis.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(chassis);
  for (let i = 0; i < 3; i++) {
    const slot = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    slot.setAttribute('x', -16); slot.setAttribute('y', -14 + (i * 12));
    slot.setAttribute('width', 32); slot.setAttribute('height', 4);
    slot.setAttribute('rx', 2); slot.setAttribute('fill', style.accent);
    group.appendChild(slot);
  }
}
function drawClusterIcon(group, node, style) {
  [[-18,-8],[0,16],[18,-8]].forEach(([x,y]) => {
    const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    rect.setAttribute('x', x - 12); rect.setAttribute('y', y - 10);
    rect.setAttribute('width', 24); rect.setAttribute('height', 20);
    rect.setAttribute('rx', 5); rect.setAttribute('fill', style.fill);
    rect.setAttribute('stroke', '#d6e4ff'); rect.setAttribute('stroke-width', node.emphasis ? '2' : '1.2');
    group.appendChild(rect);
  });
  const link = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  link.setAttribute('d', 'M -6 -2 L 0 6 L 6 -2');
  link.setAttribute('stroke', style.accent); link.setAttribute('stroke-width', '3'); link.setAttribute('fill', 'none');
  group.appendChild(link);
}
function drawDatastoreIcon(group, node, style) {
  const top = document.createElementNS('http://www.w3.org/2000/svg', 'ellipse');
  top.setAttribute('cx', 0); top.setAttribute('cy', -18); top.setAttribute('rx', 24); top.setAttribute('ry', 8);
  top.setAttribute('fill', style.fill); top.setAttribute('stroke', '#d6e4ff'); top.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(top);
  const body = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  body.setAttribute('x', -24); body.setAttribute('y', -18); body.setAttribute('width', 48); body.setAttribute('height', 34);
  body.setAttribute('fill', style.fill); body.setAttribute('stroke', '#d6e4ff'); body.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(body);
  const bottom = document.createElementNS('http://www.w3.org/2000/svg', 'ellipse');
  bottom.setAttribute('cx', 0); bottom.setAttribute('cy', 16); bottom.setAttribute('rx', 24); bottom.setAttribute('ry', 8);
  bottom.setAttribute('fill', style.fill); bottom.setAttribute('stroke', '#d6e4ff'); bottom.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(bottom);
}
function drawNetworkIcon(group, node, style) {
  const left = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  left.setAttribute('x', -30); left.setAttribute('y', -14); left.setAttribute('width', 24); left.setAttribute('height', 28);
  left.setAttribute('rx', 6); left.setAttribute('fill', style.fill); left.setAttribute('stroke', '#d6e4ff'); left.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(left);
  const right = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  right.setAttribute('x', 6); right.setAttribute('y', -14); right.setAttribute('width', 24); right.setAttribute('height', 28);
  right.setAttribute('rx', 6); right.setAttribute('fill', style.fill); right.setAttribute('stroke', '#d6e4ff'); right.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(right);
  const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  line.setAttribute('x1', -6); line.setAttribute('y1', 0); line.setAttribute('x2', 6); line.setAttribute('y2', 0);
  line.setAttribute('stroke', style.accent); line.setAttribute('stroke-width', '4');
  group.appendChild(line);
}


function drawWarningIcon(group, node, style) {
  const tri = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  tri.setAttribute('d', 'M 0 -30 L 30 22 L -30 22 Z');
  tri.setAttribute('fill', style.fill);
  tri.setAttribute('stroke', '#d6e4ff');
  tri.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(tri);

  const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bar.setAttribute('x', -3);
  bar.setAttribute('y', -10);
  bar.setAttribute('width', 6);
  bar.setAttribute('height', 18);
  bar.setAttribute('rx', 2);
  bar.setAttribute('fill', style.accent);
  group.appendChild(bar);

  const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  dot.setAttribute('cx', 0);
  dot.setAttribute('cy', 14);
  dot.setAttribute('r', 3.2);
  dot.setAttribute('fill', style.accent);
  group.appendChild(dot);
}

function drawFileSystemIcon(group, node, style) {
  const drive = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  drive.setAttribute('x', -30); drive.setAttribute('y', -18);
  drive.setAttribute('width', 60); drive.setAttribute('height', 36);
  drive.setAttribute('rx', 8); drive.setAttribute('fill', style.fill);
  drive.setAttribute('stroke', '#d6e4ff'); drive.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(drive);
  const bar = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bar.setAttribute('x', -30); bar.setAttribute('y', -18);
  bar.setAttribute('width', 60); bar.setAttribute('height', 8);
  bar.setAttribute('rx', 8); bar.setAttribute('fill', style.accent);
  group.appendChild(bar);
  const led = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  led.setAttribute('cx', 20); led.setAttribute('cy', 0); led.setAttribute('r', 3);
  led.setAttribute('fill', '#d6e4ff');
  group.appendChild(led);
}

function drawSwitchIcon(group, node, style) {
  const base = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  base.setAttribute('x', -34); base.setAttribute('y', -18); base.setAttribute('width', 68); base.setAttribute('height', 36);
  base.setAttribute('rx', 8); base.setAttribute('fill', style.fill); base.setAttribute('stroke', '#d6e4ff'); base.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(base);
  for (let i = 0; i < 6; i++) {
    const port = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    port.setAttribute('x', -24 + (i * 8)); port.setAttribute('y', -6); port.setAttribute('width', 5); port.setAttribute('height', 10);
    port.setAttribute('rx', 1.5); port.setAttribute('fill', style.accent);
    group.appendChild(port);
  }
}
function drawGroupIcon(group, node, style) {
  const back = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  back.setAttribute('x', -28); back.setAttribute('y', -16);
  back.setAttribute('width', 54); back.setAttribute('height', 34);
  back.setAttribute('rx', 8); back.setAttribute('fill', style.fill); back.setAttribute('opacity', '0.7');
  back.setAttribute('stroke', '#d6e4ff'); back.setAttribute('stroke-width', node.emphasis ? '2' : '1.2');
  group.appendChild(back);

  const mid = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  mid.setAttribute('x', -18); mid.setAttribute('y', -24);
  mid.setAttribute('width', 54); mid.setAttribute('height', 34);
  mid.setAttribute('rx', 8); mid.setAttribute('fill', style.fill); mid.setAttribute('opacity', '0.82');
  mid.setAttribute('stroke', '#d6e4ff'); mid.setAttribute('stroke-width', node.emphasis ? '2' : '1.2');
  group.appendChild(mid);

  const front = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  front.setAttribute('x', -8); front.setAttribute('y', -32);
  front.setAttribute('width', 54); front.setAttribute('height', 34);
  front.setAttribute('rx', 8); front.setAttribute('fill', style.fill);
  front.setAttribute('stroke', '#d6e4ff'); front.setAttribute('stroke-width', node.emphasis ? '2' : '1.4');
  group.appendChild(front);

  const badge = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
  badge.setAttribute('cx', 38); badge.setAttribute('cy', -24); badge.setAttribute('r', 14);
  badge.setAttribute('fill', style.accent); badge.setAttribute('stroke', '#d6e4ff'); badge.setAttribute('stroke-width', '1.2');
  group.appendChild(badge);

  const countText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  countText.setAttribute('x', 38); countText.setAttribute('y', -20);
  countText.setAttribute('text-anchor', 'middle');
  countText.setAttribute('class', 'topology-label');
  countText.setAttribute('font-size', '11');
  countText.textContent = String(node.count ?? '');
  group.appendChild(countText);
}
function drawNodeGlyph(group, node, style) {
  switch (node.kind) {
    case 'vm': return drawVmIcon(group, node, style, false);
    case 'template': return drawVmIcon(group, node, style, true);
    case 'host': return drawHostIcon(group, node, style);
    case 'cluster': return drawClusterIcon(group, node, style);
    case 'datastore': return drawDatastoreIcon(group, node, style);
    case 'network': return drawNetworkIcon(group, node, style);
    case 'filesystem': return drawFileSystemIcon(group, node, style);
    case 'warning': return drawWarningIcon(group, node, style);
    case 'switch': return drawSwitchIcon(group, node, style);
    case 'group': return drawGroupIcon(group, node, style);
    default: return drawVmIcon(group, node, style, false);
  }
}
function enableSvgPanZoom(svg, viewport, panSurface) {
  let scale = 1;
  let tx = 0;
  let ty = 0;
  let dragging = false;
  let moved = false;
  let start = null;
  const apply = () => viewport.setAttribute('transform', `translate(${tx} ${ty}) scale(${scale})`);
  apply();

  svg.onwheel = (e) => {
    e.preventDefault();
    const factor = e.deltaY < 0 ? 1.12 : 0.9;
    scale = Math.max(0.45, Math.min(2.4, scale * factor));
    apply();
  };

  const panTarget = panSurface || svg;

  panTarget.onpointerdown = (e) => {
    if (e.target !== panTarget) return;
    dragging = true;
    moved = false;
    start = { x: e.clientX, y: e.clientY, tx, ty };
    try { panTarget.setPointerCapture(e.pointerId); } catch {}
  };

  panTarget.onpointermove = (e) => {
    if (!dragging || !start) return;
    const dx = e.clientX - start.x;
    const dy = e.clientY - start.y;
    if (Math.abs(dx) > 4 || Math.abs(dy) > 4) moved = true;
    tx = start.tx + dx;
    ty = start.ty + dy;
    apply();
  };

  panTarget.onpointerup = (e) => {
    dragging = false;
    start = null;
    try { panTarget.releasePointerCapture(e.pointerId); } catch {}
  };

  panTarget.onpointerleave = () => {
    dragging = false;
    start = null;
  };

  const zoomIn = document.getElementById('topologyZoomIn');
  const zoomOut = document.getElementById('topologyZoomOut');
  const reset = document.getElementById('topologyReset');
  if (zoomIn) zoomIn.onclick = () => { scale = Math.min(2.4, scale * 1.15); apply(); };
  if (zoomOut) zoomOut.onclick = () => { scale = Math.max(0.45, scale * 0.88); apply(); };
  if (reset) reset.onclick = () => { scale = 1; tx = 0; ty = 0; apply(); };
}

function buildAdjacency(graph) {
  const adjacency = new Map();
  graph.nodes.forEach(node => adjacency.set(node.id, new Set()));
  graph.edges.forEach((edge, idx) => {
    if (!adjacency.has(edge.from)) adjacency.set(edge.from, new Set());
    if (!adjacency.has(edge.to)) adjacency.set(edge.to, new Set());
    adjacency.get(edge.from).add(edge.to);
    adjacency.get(edge.to).add(edge.from);
    edge._index = idx;
  });
  return adjacency;
}
function traverseConnected(startId, adjacency) {
  const visited = new Set();
  const queue = [startId];
  while (queue.length) {
    const current = queue.shift();
    if (!current || visited.has(current)) continue;
    visited.add(current);
    const neighbors = adjacency.get(current) || new Set();
    neighbors.forEach(n => { if (!visited.has(n)) queue.push(n); });
  }
  return visited;
}
function getImmediateNeighborhood(startId, adjacency) {
  const visible = new Set([startId]);
  const neighbors = adjacency.get(startId) || new Set();
  neighbors.forEach(n => visible.add(n));
  return visible;
}
function applyTopologyState(svg, focusId, mode='select') {
  const viewport = svg.querySelector('g');
  if (!viewport) return;
  const nodes = [...viewport.querySelectorAll('.topology-node')];
  const edges = [...viewport.querySelectorAll('.topology-edge')];
  if (!focusId) {
    nodes.forEach(n => n.classList.remove('active','faded'));
    edges.forEach(e => e.classList.remove('active','faded'));
    return;
  }
  const graphData = window.__dmtoolsTopologyGraph;
  if (!graphData) return;

  const activeNodes = mode === 'hover'
    ? getImmediateNeighborhood(focusId, graphData.adjacency)
    : traverseConnected(focusId, graphData.adjacency);

  nodes.forEach(node => {
    const id = node.getAttribute('data-node-id');
    const isActive = activeNodes.has(id);
    node.classList.toggle('active', isActive);
    node.classList.toggle('faded', !isActive);
  });

  edges.forEach(edge => {
    const from = edge.getAttribute('data-from');
    const to = edge.getAttribute('data-to');
    const isActive = mode === 'hover'
      ? (from === focusId || to === focusId)
      : (activeNodes.has(from) && activeNodes.has(to));
    edge.classList.toggle('active', isActive);
    edge.classList.toggle('faded', !isActive);
  });
}

function renderTopology(item, entity) {
  const svg = document.getElementById('topologySvg');
  if (!svg) return;

  const graph = buildTopologyGraph(item, entity);
  layoutTopology(graph.nodes, entity);
  applyGroupFanout(graph.nodes);

  svg.innerHTML = '';

  const panSurface = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  panSurface.setAttribute('x', '0');
  panSurface.setAttribute('y', '0');
  panSurface.setAttribute('width', '1280');
  panSurface.setAttribute('height', '760');
  panSurface.setAttribute('fill', 'transparent');
  panSurface.setAttribute('pointer-events', 'all');
  svg.appendChild(panSurface);

  const viewport = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  svg.appendChild(viewport);

  graph.edges.forEach(edge => {
    const from = graph.nodes.find(n => n.id === edge.from);
    const to = graph.nodes.find(n => n.id === edge.to);
    if (!from || !to) return;

    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    line.setAttribute('x1', from.x);
    line.setAttribute('y1', from.y);
    line.setAttribute('x2', to.x);
    line.setAttribute('y2', to.y);
    line.setAttribute('class', 'topology-edge');
    line.setAttribute('data-from', from.id);
    line.setAttribute('data-to', to.id);
    viewport.appendChild(line);
  });

  graph.nodes.forEach(node => {
    const style = nodeStyle(node.kind, node.emphasis);
    const group = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    group.setAttribute('class', 'topology-node');
    group.setAttribute('transform', `translate(${node.x} ${node.y})`);
    group.setAttribute('data-node-id', node.id);

    const focusRing = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    focusRing.setAttribute('cx', 0); focusRing.setAttribute('cy', 0); focusRing.setAttribute('r', 48);
    focusRing.setAttribute('fill', '#ffd54f14');
    focusRing.setAttribute('stroke', '#ffd54f');
    focusRing.setAttribute('stroke-width', '2.5');
    focusRing.setAttribute('class', 'topology-focus-ring');
    group.appendChild(focusRing);

    if (node.expandId) {
      group.setAttribute('data-expand-id', node.expandId);
      group.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        window.__dmtoolsTopologySelectedNodeId = node.id;
        toggleTopologyGroup(node.expandId);
      });
    } else if (node.externalUrl) {
      group.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        window.__dmtoolsTopologySelectedNodeId = node.id;
        applyTopologyState(svg, node.id, 'select');
        window.open(node.externalUrl, '_blank', 'noopener');
      });
    } else if (node.entity && node.entity !== 'none') {
      group.setAttribute('data-nav-entity', node.entity);
      group.setAttribute('data-nav-id', slugify(node.name));
      group.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        window.__dmtoolsTopologySelectedNodeId = node.id;
        applyTopologyState(svg, node.id, 'select');
        navigateTo(node.entity, slugify(node.name));
      });
    } else {
      group.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        window.__dmtoolsTopologySelectedNodeId = node.id;
        applyTopologyState(svg, node.id, 'select');
      });
    }

    group.addEventListener('mouseenter', () => {
      if (window.__dmtoolsTopologySelectedNodeId) return;
      applyTopologyState(svg, node.id, 'hover');
    });
    group.addEventListener('mouseleave', () => {
      if (window.__dmtoolsTopologySelectedNodeId) {
        applyTopologyState(svg, window.__dmtoolsTopologySelectedNodeId, 'select');
      } else {
        applyTopologyState(svg, null);
      }
    });

    if (node.emphasis) {
      const halo = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      halo.setAttribute('cx', 0); halo.setAttribute('cy', 0); halo.setAttribute('r', 44);
      halo.setAttribute('fill', '#4fc3f733');
      group.appendChild(halo);
    }

    drawNodeGlyph(group, node, style);

    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', 0); text.setAttribute('y', -42);
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('class', 'topology-label');
    text.textContent = node.label.length > 22 ? node.label.slice(0, 21) + '…' : node.label;
    group.appendChild(text);

    if (node.meta) {
      const sub = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      sub.setAttribute('x', 0); sub.setAttribute('y', -28);
      sub.setAttribute('text-anchor', 'middle');
      sub.setAttribute('class', 'topology-sublabel');
      sub.textContent = node.meta;
      group.appendChild(sub);
    }

    viewport.appendChild(group);
  });

  window.__dmtoolsTopologyGraph = { nodes: graph.nodes, edges: graph.edges, adjacency: buildAdjacency(graph) };
  window.__dmtoolsTopologySelectedNodeId = null;
  applyTopologyState(svg, null, 'select');
  enableSvgPanZoom(svg, viewport, panSurface);
}

function renderPanels() {
  renderContextFilters();
  el.overviewPanel.classList.toggle('active', currentView === 'ALL');
  el.detailPanel.classList.toggle('active', currentView === 'DETAIL');
  el.overviewPanel.style.display = currentView === 'ALL' ? 'flex' : 'none';
  el.detailPanel.style.display = currentView === 'DETAIL' ? 'flex' : 'none';
  if (currentView === 'ALL') renderOverview(); else renderDetail();
}
function navigateTo(entity, id) {
  if (!entity || !id) return;
  window.__dmtoolsTopologySelectedNodeId = null;
  currentEntity = entity;
  currentFilter = 'ALL';
  currentView = 'DETAIL';
  currentSelectedId = id;
  setEntityButtonState();
  setFilterButtonState();
  setViewButtonState();
  renderNav();
  renderPanels();
}
document.addEventListener('click', (event) => {
  const nav = event.target.closest('[data-nav-entity]');
  if (nav) {
    event.preventDefault();
    navigateTo(nav.dataset.navEntity, nav.dataset.navId);
  }
});
document.querySelectorAll('.entity-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    currentEntity = btn.dataset.entity;
    currentSelectedId = null;
    currentView = 'ALL';
    setEntityButtonState();
    setViewButtonState();
    renderNav();
    renderPanels();
  });
});
document.querySelectorAll('.filter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    currentFilter = btn.dataset.filter;
    currentSelectedId = null;
    setFilterButtonState();
    renderNav();
    renderPanels();
  });
});
document.querySelectorAll('.view-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    currentView = btn.dataset.view;
    if (currentView === 'DETAIL' && !currentSelectedId) {
      const first = filteredItems()[0];
      currentSelectedId = first ? first.id : null;
    }
    setViewButtonState();
    renderPanels();
  });
});

setSummary();
setEntityButtonState();
setFilterButtonState();
setViewButtonState();
refreshEntityView();
</script>
</body>
</html>
'@

$html = $html.Replace('__REPORT_DATA__', $json)

$null = New-Item -ItemType Directory -Force -Path (Split-Path -Parent $OutputHtml)
[System.IO.File]::WriteAllText($OutputHtml, $html, [System.Text.UTF8Encoding]::new($false))

Write-Host "HTML report written to $OutputHtml" -ForegroundColor Green
