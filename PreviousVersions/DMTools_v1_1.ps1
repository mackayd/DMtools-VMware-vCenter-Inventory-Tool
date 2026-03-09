
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
    - Requires PowerShell 7 and Windows OS.
    - Requires network connectivity to the vCenter Server.
    - Script must be run with sufficient privileges to install modules and connect to vCenter.
    - May take several minutes to complete, depending on environment size.

.AUTHOR
    Drew Mackay
    https://github.com/mackayd

.VERSION
    1.1

.LICENSE
    MIT License

#>

param(
    [switch]$RedactVMNames,
    [switch]$RedactFqdnDomain,
    [switch]$RedactIPAddresses
)
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

function Write-DMConsoleBanner {
    param(
        [string]$Title = 'DMTools VMware Inventory Export'
    )

    Write-Host ''
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host (" " + $Title) -ForegroundColor Cyan
    Write-Host '============================================================' -ForegroundColor DarkCyan
    Write-Host ''
}

function Write-DMConsoleSection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    Write-Host $Title -ForegroundColor Yellow
    Write-Host ('-' * $Title.Length) -ForegroundColor DarkYellow
}

function Read-RedactionChoice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,
        [bool]$Default = $false,
        [ConsoleColor]$PromptColor = 'Gray'
    )

    $suffix = if ($Default) { "Y/n" } else { "y/N" }
    Write-Host (("{0} [{1}]: " -f $Prompt, $suffix)) -ForegroundColor $PromptColor -NoNewline
    $answer = Read-Host
    if ([string]::IsNullOrWhiteSpace($answer)) {
        return $Default
    }
    return $answer.Trim().ToLowerInvariant().StartsWith("y")
}

function Remove-FqdnDomain {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -match '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') {
        return ($Value -split '\.')[0]
    }
    return $Value
}

function Mask-IPAddresses {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $masked = $Value -replace '(?<!\d)(?:\d{1,3}\.){3}\d{1,3}(?!\d)', '[REDACTED-IPV4]'
    $masked = $masked -replace '(?i)\b(?:[0-9a-f]{1,4}:){2,7}[0-9a-f]{1,4}\b', '[REDACTED-IPV6]'
    return $masked
}

function New-VMRedactionMap {
    param($vms)
    $map = @{}
    $index = 1
    foreach ($vmName in ($vms | Select-Object -ExpandProperty Name | Sort-Object -Unique)) {
        $map[$vmName] = ('VM-REDACTED-{0:D4}' -f $index)
        $index++
    }
    return $map
}


function Invoke-Redaction {
    param(
        $Data,
        [Parameter(Mandatory = $true)]
        [string]$SheetName,
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        [hashtable]$VMNameMap,
        [string[]]$VMNameSheets
    )

    if ($null -eq $Data) { return $Data }

    $rows = @($Data)
    foreach ($row in $rows) {
        if ($null -eq $row) { continue }

        foreach ($prop in $row.PSObject.Properties) {
            $name = $prop.Name
            $value = $prop.Value
            if ($null -eq $value) { continue }

            if ($Config.RedactVMNames -and $name -eq 'Name' -and ($VMNameSheets -contains $SheetName)) {
                $vmName = [string]$value
                if ($VMNameMap.ContainsKey($vmName)) {
                    $prop.Value = $VMNameMap[$vmName]
                }
                else {
                    $prop.Value = 'VM-REDACTED-UNKNOWN'
                }
                continue
            }

            if ($value -is [string]) {
                $newValue = $value

                if ($Config.RedactVMNames -and $VMNameMap.Count -gt 0) {
                    foreach ($vmName in $VMNameMap.Keys | Sort-Object Length -Descending) {
                        if ([string]::IsNullOrWhiteSpace($vmName)) { continue }
                        $replacement = [string]$VMNameMap[$vmName]
                        $escapedVmName = [regex]::Escape($vmName)
                        $newValue = [regex]::Replace($newValue, $escapedVmName, $replacement, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    }
                }

                if ($Config.RedactFqdnDomain) {
                    $newValue = Remove-FqdnDomain -Value $newValue
                }
                if ($Config.RedactIPAddresses) {
                    $newValue = Mask-IPAddresses -Value $newValue
                }
                if ($newValue -ne $value) {
                    $prop.Value = $newValue
                }
            }
        }
    }
    return $Data
}
function Get-NormalizedColumnName {
    param([string]$ColumnName)
    if ([string]::IsNullOrWhiteSpace($ColumnName)) { return '' }
    return (($ColumnName -replace '[^A-Za-z0-9]', '').ToLowerInvariant())
}

function Get-OutputHeaderMap {
    return @{
        'vInfo' = @('VM','Powerstate','Template','SRM Placeholder','Config status','DNS Name','Connection state','Guest state','Heartbeat','Consolidation Needed','PowerOn','Suspended To Memory','Suspend time','Suspend Interval','Creation date','Change Version','CPUs','Overall Cpu Readiness','Memory','Active Memory','NICs','Disks','Total disk capacity MiB','Fixed Passthru HotPlug','min Required EVC Mode Key','Latency Sensitivity','Op Notification Timeout','EnableUUID','CBT','Primary IP Address','Network #1','Network #2','Network #3','Network #4','Network #5','Network #6','Network #7','Network #8','Num Monitors','Video Ram KiB','Resource pool','Folder ID','Folder','vApp','DAS protection','FT State','FT Role','FT Latency','FT Bandwidth','FT Sec. Latency','Vm Failover In Progress','Provisioned MiB','In Use MiB','Unshared MiB','HA Restart Priority','HA Isolation Response','HA VM Monitoring','Cluster rule(s)','Cluster rule name(s)','Boot Required','Boot delay','Boot retry delay','Boot retry enabled','Boot BIOS setup','Reboot PowerOff','EFI Secure boot','Firmware','HW version','HW upgrade status','HW upgrade policy','HW target','Path','Log directory','Snapshot directory','Suspend directory','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','OS according to the configuration file','OS according to the VMware Tools','Customization Info','Guest Detailed Data','VM ID','SMBIOS UUID','VM UUID','VI SDK Server type','VI SDK API Version','VI SDK Server','VI SDK UUID')
        'vCPU' = @('VM','Powerstate','Template','SRM Placeholder','CPUs','Sockets','Cores p/s','Max','Overall','Level','Shares','Reservation','Entitlement','DRS Entitlement','Limit','Hot Add','Hot Remove','Numa Hotadd Exposed','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vMemory' = @('VM','Powerstate','Template','SRM Placeholder','Size MiB','Memory Reservation Locked To Max','Overhead','Max','Consumed','Consumed Overhead','Private','Shared','Swapped','Ballooned','Active','Entitlement','DRS Entitlement','Level','Shares','Reservation','Limit','Hot Add','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vDisk' = @('VM','Powerstate','Template','SRM Placeholder','Disk','Disk Key','Disk UUID','Disk Path','Capacity MiB','Raw','Disk Mode','Sharing mode','Thin','Eagerly Scrub','Split','Write Through','Level','Shares','Reservation','Limit','Controller','Label','SCSI Unit #','Unit #','Shared Bus','Path','Raw LUN ID','Raw Comp. Mode','Internal Sort Column','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vPartition' = @('VM','Powerstate','Template','SRM Placeholder','Disk Key','Disk','Capacity MiB','Consumed MiB','Free MiB','Free %','Internal Sort Column','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vNetwork' = @('VM','Powerstate','Template','SRM Placeholder','NIC label','Adapter','Network','Switch','Connected','Starts Connected','Mac Address','Type','IPv4 Address','IPv6 Address','Direct Path IO','Internal Sort Column','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vCD' = @('VM','Powerstate','Template','SRM Placeholder','Device Node','Connected','Starts Connected','Device Type','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VMRef','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vSnapshot' = @('VM','Powerstate','Name','Description','Date / time','Filename','Size MiB (vmsn)','Size MiB (total)','Quiesced','State','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vTools' = @('VM','Powerstate','Template','SRM Placeholder','VM Version','Tools','Tools Version','Required Version','Upgradeable','Upgrade Policy','Sync time','App status','Heartbeat status','Kernel Crash state','Operation Ready','State change support','Interactive Guest','Annotation','com.vrlcm.snapshot','Datacenter','Cluster','Host','Folder','OS according to the configuration file','OS according to the VMware Tools','VMRef','VM ID','VM UUID','VI SDK Server','VI SDK UUID')
        'vRP' = @('Resource Pool name','Resource Pool path','Status','# VMs total','# VMs','# vCPUs','CPU limit','CPU overheadLimit','CPU reservation','CPU level','CPU shares','CPU expandableReservation','CPU maxUsage','CPU overallUsage','CPU reservationUsed','CPU reservationUsedForVm','CPU unreservedForPool','CPU unreservedForVm','Mem Configured','Mem limit','Mem overheadLimit','Mem reservation','Mem level','Mem shares','Mem expandableReservation','Mem maxUsage','Mem overallUsage','Mem reservationUsed','Mem reservationUsedForVm','Mem unreservedForPool','Mem unreservedForVm','QS overallCpuDemand','QS overallCpuUsage','QS staticCpuEntitlement','QS distributedCpuEntitlement','QS balloonedMemory','QS compressedMemory','QS consumedOverheadMemory','QS distributedMemoryEntitlement','QS guestMemoryUsage','QS hostMemoryUsage','QS overheadMemory','QS privateMemory','QS sharedMemory','QS staticMemoryEntitlement','QS swappedMemory','Object ID','VI SDK Server','VI SDK UUID')
        'vCluster' = @('Name','Config status','OverallStatus','NumHosts','numEffectiveHosts','TotalCpu','NumCpuCores','NumCpuThreads','Effective Cpu','TotalMemory','Effective Memory','Num VMotions','HA enabled','Failover Level','AdmissionControlEnabled','Host monitoring','HB Datastore Candidate Policy','Isolation Response','Restart Priority','Cluster Settings','Max Failures','Max Failure Window','Failure Interval','Min Up Time','VM Monitoring','DRS enabled','DRS default VM behavior','DRS vmotion rate','DPM enabled','DPM default behavior','DPM Host Power Action Rate','Object ID','com.vmware.vcenter.cluster.edrs.upgradeHostAdded','com.vrlcm.snapshot','VI SDK Server','VI SDK UUID')
        'vHost' = @('Host','Datacenter','Cluster','Config status','Compliance Check State','in Maintenance Mode','in Quarantine Mode','vSAN Fault Domain Name','CPU Model','Speed','HT Available','HT Active','# CPU','Cores per CPU','# Cores','CPU usage %','# Memory','Memory Tiering Type','Memory usage %','Console','# NICs','# HBAs','# VMs total','# VMs','VMs per Core','# vCPUs','vCPUs per Core','vRAM','VM Used memory','VM Memory Swapped','VM Memory Ballooned','VMotion support','Storage VMotion support','Current EVC','Max EVC','Assigned License(s)','ATS Heartbeat','ATS Locking','Current CPU power man. policy','Supported CPU power man.','Host Power Policy','ESX Version','Boot time','DNS Servers','DHCP','Domain','Domain List','DNS Search Order','NTP Server(s)','NTPD running','Time Zone','Time Zone Name','GMT Offset','Vendor','Model','Serial number','Service tag','OEM specific string','BIOS Vendor','BIOS Version','BIOS Date','Certificate Issuer','Certificate Start Date','Certificate Expiry Date','Certificate Status','Certificate Subject','Object ID','AutoDeploy.MachineIdentity','com.vrlcm.snapshot','UUID','VI SDK Server','VI SDK UUID')
        'vHBA' = @('Host','Datacenter','Cluster','Device','Type','Status','Bus','Pci','Driver','Model','WWN','VI SDK Server','VI SDK UUID')
        'vNIC' = @('Host','Datacenter','Cluster','Network Device','Driver','Speed','Duplex','MAC','Switch','Uplink port','PCI','WakeOn','VI SDK Server','VI SDK UUID')
        'vSwitch' = @('Host','Datacenter','Cluster','Switch','# Ports','Free Ports','Promiscuous Mode','Mac Changes','Forged Transmits','Traffic Shaping','Width','Peak','Burst','Policy','Reverse Policy','Notify Switch','Rolling Order','Offload','TSO','Zero Copy Xmit','MTU','VI SDK Server','VI SDK UUID')
        'vPort' = @('Host','Datacenter','Cluster','Port Group','Switch','VLAN','Promiscuous Mode','Mac Changes','Forged Transmits','Traffic Shaping','Width','Peak','Burst','Policy','Reverse Policy','Notify Switch','Rolling Order','Offload','TSO','Zero Copy Xmit','VI SDK Server','VI SDK UUID')
        'dvSwitch' = @('Switch','Datacenter','Name','Vendor','Version','Description','Created','Host members','Max Ports','# Ports','# VMs','In Traffic Shaping','In Avg','In Peak','In Burst','Out Traffic Shaping','Out Avg','Out Peak','Out Burst','CDP Type','CDP Operation','LACP Name','LACP Mode','LACP Load Balance Alg.','Max MTU','Contact','Admin Name','Object ID','com.vrlcm.snapshot','VI SDK Server','VI SDK UUID')
        'dvPort' = @('Port','Switch','Type','# Ports','VLAN','Speed','Full Duplex','Blocked','Allow Promiscuous','Mac Changes','Active Uplink','Standby Uplink','Policy','Forged Transmits','In Traffic Shaping','In Avg','In Peak','In Burst','Out Traffic Shaping','Out Avg','Out Peak','Out Burst','Reverse Policy','Notify Switch','Rolling Order','Check Beacon','Live Port Moving','Check Duplex','Check Error %','Check Speed','Percentage','Block Override','Config Reset','Shaping Override','Vendor Config Override','Sec. Policy Override','Teaming Override','Vlan Override','Object ID','VI SDK Server','VI SDK UUID')
        'vSC_VMK' = @('Host','Datacenter','Cluster','Port Group','Device','Mac Address','DHCP','IP Address','IP 6 Address','Subnet mask','Gateway','IP 6 Gateway','MTU','VI SDK Server','VI SDK UUID')
        'vDatastore' = @('Name','Config status','Address','Accessible','Type','# VMs total','# VMs','Capacity MiB','Provisioned MiB','In Use MiB','Free MiB','Free %','SIOC enabled','SIOC Threshold','# Hosts','Hosts','Cluster name','Cluster capacity MiB','Cluster free space MiB','Block size','Max Blocks','# Extents','Major Version','Version','VMFS Upgradeable','MHA','URL','Object ID','com.vrlcm.snapshot','VI SDK Server','VI SDK UUID')
        'vLicense' = @('Name','Key','Labels','Cost Unit','Total','Used','Expiration Date','Features','VI SDK Server','VI SDK UUID')
        'vHealth' = @('Name','Message','Message type','VI SDK Server','VI SDK UUID')
    }
}

function Convert-ToOutputSchema {
    param(
        $Data,
        [Parameter(Mandatory = $true)]
        [string]$SheetName,
        [Parameter(Mandatory = $true)]
        [hashtable]$OutputHeaderMap,
        [Parameter(Mandatory = $true)]
        [hashtable]$AliasMap
    )

    if (-not $OutputHeaderMap.ContainsKey($SheetName)) { return $Data }

    $targetHeaders = @($OutputHeaderMap[$SheetName])
    $rows = @($Data)
    if ($rows.Count -eq 0) { return $rows }

    $sheetAlias = if ($AliasMap.ContainsKey($SheetName)) { $AliasMap[$SheetName] } else { @{} }
    $result = New-Object System.Collections.Generic.List[object]

    foreach ($row in $rows) {
        $normalizedLookup = @{}
        foreach ($prop in $row.PSObject.Properties) {
            $norm = Get-NormalizedColumnName -ColumnName $prop.Name
            if (-not [string]::IsNullOrWhiteSpace($norm) -and -not $normalizedLookup.ContainsKey($norm)) {
                $normalizedLookup[$norm] = $prop.Name
            }
        }

        $aligned = [ordered]@{}
        foreach ($targetHeader in $targetHeaders) {
            $value = ''
            if ($row.PSObject.Properties.Name -contains $targetHeader) {
                $value = $row.$targetHeader
            }
            elseif ($sheetAlias.ContainsKey($targetHeader) -and ($row.PSObject.Properties.Name -contains $sheetAlias[$targetHeader])) {
                $value = $row.($sheetAlias[$targetHeader])
            }
            else {
                $targetNorm = Get-NormalizedColumnName -ColumnName $targetHeader
                if ($normalizedLookup.ContainsKey($targetNorm)) {
                    $value = $row.($normalizedLookup[$targetNorm])
                }
            }
            $aligned[$targetHeader] = $value
        }

        $result.Add([PSCustomObject]$aligned)
    }

    return $result
}
Test-Module -ModuleName VMware.PowerCLI
Test-Module -ModuleName ImportExcel
Test-Module -ModuleName psInlineProgress
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false -InvalidCertificateAction Ignore -DisplayDeprecationWarnings $false | Out-Null
$doRedactVMNames = [bool]$RedactVMNames
$doRedactFqdnDomain = [bool]$RedactFqdnDomain
$doRedactIPAddresses = [bool]$RedactIPAddresses
$doPromptRedactionOptions = $doRedactVMNames -or $doRedactFqdnDomain -or $doRedactIPAddresses

$redactionParamsProvided = (
    $PSBoundParameters.ContainsKey('RedactVMNames') -or
    $PSBoundParameters.ContainsKey('RedactFqdnDomain') -or
    $PSBoundParameters.ContainsKey('RedactIPAddresses')
)

Write-DMConsoleBanner

Write-DMConsoleSection -Title 'Export Options'
if (-not $redactionParamsProvided) {
    $doPromptRedactionOptions = Read-RedactionChoice -Prompt 'Apply redaction to the Excel export?' -Default $false -PromptColor Cyan
}

if ($doPromptRedactionOptions) {
    Write-Host ''
    Write-DMConsoleSection -Title 'Redaction Options'
    if (-not $PSBoundParameters.ContainsKey('RedactVMNames')) {
        $doRedactVMNames = Read-RedactionChoice -Prompt 'Redact VM names?' -Default $false -PromptColor Gray
    }
    if (-not $PSBoundParameters.ContainsKey('RedactFqdnDomain')) {
        $doRedactFqdnDomain = Read-RedactionChoice -Prompt 'Redact FQDN domain suffixes?' -Default $false -PromptColor Gray
    }
    if (-not $PSBoundParameters.ContainsKey('RedactIPAddresses')) {
        $doRedactIPAddresses = Read-RedactionChoice -Prompt 'Redact IP addresses?' -Default $false -PromptColor Gray
    }
}

Write-Host ''
Write-DMConsoleSection -Title 'Connection'
Write-Host 'vCenter Server (FQDN or IP): ' -ForegroundColor Cyan -NoNewline
$vcenter = Read-Host
$cred = Get-Credential -Message "DMTools connection for $vcenter"
try {
    Connect-VIServer -Server $vcenter -Credential $cred -ErrorAction Stop | Out-Null
    Write-Host "Connected to $vcenter successfully." -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to $vcenter. Exiting." -ForegroundColor Red
    exit 1
}
$excelFile = Get-ExcelFilePath-GUI
$DMexecstart = Get-date
# ---- ROOT DATA COLLECTION ----
$si = Get-View ServiceInstance
$about = $si.Content.About
$vms = @(
    Get-VM -ErrorAction SilentlyContinue
    Get-Template -ErrorAction SilentlyContinue
) | Sort-Object -Property Id -Unique
$ESXhosts = Get-VMHost
$clusters = Get-Cluster
$rpools = Get-ResourcePool
$datastores = Get-Datastore
$VCC = $global:DefaultVIServer[0]


function Get-SafeValue {
    param(
        [Parameter(Mandatory = $false)] $Value,
        [Parameter(Mandatory = $false)] $Default = $null
    )
    if ($null -eq $Value) { return $Default }
    return $Value
}

function Join-NonEmpty {
    param(
        [Parameter(Mandatory = $false)] [object[]]$Values,
        [string]$Separator = ', '
    )
    if ($null -eq $Values) { return '' }
    $filtered = @($Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    if ($filtered.Count -eq 0) { return '' }
    return ($filtered -join $Separator)
}

function Get-VMAdvancedSettingValue {
    param(
        [Parameter(Mandatory = $true)] $VM,
        [Parameter(Mandatory = $true)] [string]$Name
    )

    if (-not $script:VMAdvancedSettingCache) { $script:VMAdvancedSettingCache = @{} }

    $vmKey = $null
    try { $vmKey = $VM.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($vmKey)) { try { $vmKey = $VM.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($vmKey)) { $vmKey = [string]$VM.Name }
    $cacheKey = "$vmKey|$Name"

    if ($script:VMAdvancedSettingCache.ContainsKey($cacheKey)) {
        return $script:VMAdvancedSettingCache[$cacheKey]
    }

    $value = $null
    try { $value = (Get-AdvancedSetting -Entity $VM -Name $Name -ErrorAction Stop).Value } catch {}
    $script:VMAdvancedSettingCache[$cacheKey] = $value
    return $value
}
function Get-VMGuestNetworkDetails {
    param(
        [Parameter(Mandatory = $true)] $VM,
        [Parameter(Mandatory = $true)] $Nic
    )

    if (-not $script:VMGuestNetCache) { $script:VMGuestNetCache = @{} }

    $vmKey = $null
    try { $vmKey = $VM.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($vmKey)) { try { $vmKey = $VM.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($vmKey)) { $vmKey = [string]$VM.Name }

    if (-not $script:VMGuestNetCache.ContainsKey($vmKey)) {
        $macMap = @{}
        try {
            foreach ($guestNet in @($VM.ExtensionData.Guest.Net)) {
                $guestMac = [string]$guestNet.MacAddress
                if ([string]::IsNullOrWhiteSpace($guestMac)) { continue }
                $macKey = $guestMac.ToLowerInvariant()

                $ipv4 = @()
                $ipv6 = @()
                foreach ($ip in @($guestNet.IpAddress)) {
                    if ([string]::IsNullOrWhiteSpace($ip)) { continue }
                    if ($ip -match ':') { $ipv6 += $ip } else { $ipv4 += $ip }
                }

                $dnsName = $null
                try {
                    if ($guestNet.DnsConfig -and -not [string]::IsNullOrWhiteSpace($guestNet.DnsConfig.HostName)) {
                        $dnsName = $guestNet.DnsConfig.HostName
                    }
                }
                catch {}

                $macMap[$macKey] = [PSCustomObject]@{
                    IPv4    = Join-NonEmpty -Values $ipv4
                    IPv6    = Join-NonEmpty -Values $ipv6
                    DnsName = $dnsName
                }
            }
        }
        catch {}
        $script:VMGuestNetCache[$vmKey] = $macMap
    }

    $nicKey = [string]$Nic.MacAddress
    if (-not [string]::IsNullOrWhiteSpace($nicKey)) {
        $nicKey = $nicKey.ToLowerInvariant()
        $macMap = $script:VMGuestNetCache[$vmKey]
        if ($macMap -and $macMap.ContainsKey($nicKey)) {
            return $macMap[$nicKey]
        }
    }

    return [PSCustomObject]@{ IPv4 = $null; IPv6 = $null; DnsName = $null }
}
function Get-VMRuleSummary {
    param(
        [Parameter(Mandatory = $true)] $VM
    )

    if (-not $script:ClusterRuleCache) { $script:ClusterRuleCache = @{} }

    $clusterId = $null
    try {
        if (-not (Test-IsTemplateLikeObject -vm $VM)) {
            $cluster = $VM | Get-Cluster -ErrorAction Stop
            if ($cluster) { $clusterId = $cluster.Id }
        }
    }
    catch {}

    if ([string]::IsNullOrWhiteSpace($clusterId)) {
        return [PSCustomObject]@{ RuleNames = $null; EnabledRuleNames = $null }
    }

    if (-not $script:ClusterRuleCache.ContainsKey($clusterId)) {
        $ruleIndex = @{}
        try {
            $clusterView = Get-View -Id $clusterId -ErrorAction Stop
            foreach ($rule in @($clusterView.ConfigurationEx.Rule)) {
                $vmRefs = @()
                if ($rule.PSObject.Properties.Name -contains 'Vm') { $vmRefs = @($rule.Vm) }
                foreach ($ref in $vmRefs) {
                    $vmId = $null
                    try { $vmId = $ref.Value } catch {}
                    if ([string]::IsNullOrWhiteSpace($vmId)) { continue }
                    if (-not $ruleIndex.ContainsKey($vmId)) {
                        $ruleIndex[$vmId] = [PSCustomObject]@{ Names = New-Object System.Collections.Generic.List[string]; Enabled = New-Object System.Collections.Generic.List[string] }
                    }
                    $ruleIndex[$vmId].Names.Add([string]$rule.Name)
                    if ($rule.Enabled) { $ruleIndex[$vmId].Enabled.Add([string]$rule.Name) }
                }
            }
        }
        catch {}
        $script:ClusterRuleCache[$clusterId] = $ruleIndex
    }

    $vmMoRef = $null
    try { $vmMoRef = $VM.ExtensionData.MoRef.Value } catch {}
    if ([string]::IsNullOrWhiteSpace($vmMoRef)) { return [PSCustomObject]@{ RuleNames = $null; EnabledRuleNames = $null } }

    $index = $script:ClusterRuleCache[$clusterId]
    if (-not $index.ContainsKey($vmMoRef)) {
        return [PSCustomObject]@{ RuleNames = $null; EnabledRuleNames = $null }
    }

    $entry = $index[$vmMoRef]
    return [PSCustomObject]@{
        RuleNames = Join-NonEmpty -Values $entry.Names
        EnabledRuleNames = Join-NonEmpty -Values $entry.Enabled
    }
}
function Get-VMContext {
    param($vm)

    if (-not $script:VMContextCache) { $script:VMContextCache = @{} }

    $vmKey = $null
    try { $vmKey = $vm.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($vmKey)) { try { $vmKey = $vm.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($vmKey)) { $vmKey = [string]$vm.Name }

    if ($script:VMContextCache.ContainsKey($vmKey)) {
        return $script:VMContextCache[$vmKey]
    }

    $clusterName = $null
    $datacenterName = $null
    $resourcePoolName = $null
    $vappName = $null
    $folderName = $null

    try { $folderName = $vm.Folder.Name } catch {}
    try { $resourcePoolName = $vm.ResourcePool.Name } catch {}
    try { $vappName = $vm.VApp.Name } catch {}

    try {
        $parentRef = $vm.ExtensionData.Parent
        $safety = 0
        while ($parentRef -and $safety -lt 20) {
            $parentView = Get-View -Id $parentRef -Property Name,Parent,ResourcePool -ErrorAction SilentlyContinue
            if (-not $parentView) { break }

            switch ($parentView.MoRef.Type) {
                'Folder' {
                    if (-not $folderName) { $folderName = $parentView.Name }
                }
                'VirtualApp' {
                    if (-not $vappName) { $vappName = $parentView.Name }
                    try {
                        if (-not $resourcePoolName -and $parentView.ResourcePool) {
                            $rpView = Get-View -Id $parentView.ResourcePool -Property Name -ErrorAction SilentlyContinue
                            if ($rpView) { $resourcePoolName = $rpView.Name }
                        }
                    } catch {}
                }
                'ResourcePool' {
                    if (-not $resourcePoolName) { $resourcePoolName = $parentView.Name }
                }
                'ClusterComputeResource' {
                    if (-not $clusterName) { $clusterName = $parentView.Name }
                }
                'Datacenter' {
                    if (-not $datacenterName) { $datacenterName = $parentView.Name }
                }
            }

            $parentRef = $parentView.Parent
            $safety++
        }
    }
    catch {}

    if (-not $clusterName -or -not $datacenterName) {
        try {
            $hostRef = $vm.ExtensionData.Runtime.Host
            if ($hostRef) {
                $hostView = Get-View -Id $hostRef -Property Name,Parent -ErrorAction SilentlyContinue
                $parentRef = if ($hostView) { $hostView.Parent } else { $null }
                $safety = 0
                while ($parentRef -and $safety -lt 15) {
                    $parentView = Get-View -Id $parentRef -Property Name,Parent -ErrorAction SilentlyContinue
                    if (-not $parentView) { break }
                    switch ($parentView.MoRef.Type) {
                        'ClusterComputeResource' { if (-not $clusterName) { $clusterName = $parentView.Name } }
                        'Datacenter' { if (-not $datacenterName) { $datacenterName = $parentView.Name } }
                    }
                    $parentRef = $parentView.Parent
                    $safety++
                }
            }
        }
        catch {}
    }

    $ctx = @{
        Cluster      = $clusterName
        Datacenter   = $datacenterName
        ResourcePool = $resourcePoolName
        vApp         = $vappName
        Folder       = $folderName
    }

    $script:VMContextCache[$vmKey] = $ctx
    return $ctx
}
function Test-IsTemplateLikeObject {
    param($vm)
    try {
        if ($vm.ExtensionData -and $vm.ExtensionData.Config -and $vm.ExtensionData.Config.Template) { return $true }
    }
    catch {}
    try {
        if ($vm.GetType().Name -match 'Template') { return $true }
    }
    catch {}
    return $false
}

function Get-VMRuntimeHostName {
    param($vm)

    if (-not $script:VMHostNameCache) { $script:VMHostNameCache = @{} }
    $vmKey = $null
    try { $vmKey = $vm.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($vmKey)) { try { $vmKey = $vm.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($vmKey)) { $vmKey = [string]$vm.Name }

    if ($script:VMHostNameCache.ContainsKey($vmKey)) { return $script:VMHostNameCache[$vmKey] }

    $name = $null
    try { if ($vm.VMHost -and $vm.VMHost.Name) { $name = $vm.VMHost.Name } } catch {}
    if (-not $name) {
        try {
            $hostRef = $vm.ExtensionData.Runtime.Host
            if ($hostRef) {
                $hostView = Get-View -Id $hostRef -Property Name -ErrorAction SilentlyContinue
                if ($hostView) { $name = $hostView.Name }
            }
        }
        catch {}
    }

    $script:VMHostNameCache[$vmKey] = $name
    return $name
}
function Get-VMEffectivePowerState {
    param($vm)

    if (-not $script:VMPowerStateCache) { $script:VMPowerStateCache = @{} }
    $vmKey = $null
    try { $vmKey = $vm.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($vmKey)) { try { $vmKey = $vm.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($vmKey)) { $vmKey = [string]$vm.Name }

    if ($script:VMPowerStateCache.ContainsKey($vmKey)) { return $script:VMPowerStateCache[$vmKey] }

    $state = $null
    try {
        if ($null -ne $vm.PowerState) {
            $state = [string]$vm.PowerState
        }
    }
    catch {}

    if ([string]::IsNullOrWhiteSpace($state)) {
        try {
            if ($null -ne $vm.ExtensionData.Runtime.PowerState) {
                $state = [string]$vm.ExtensionData.Runtime.PowerState
            }
        }
        catch {}
    }

    $script:VMPowerStateCache[$vmKey] = $state
    return $state
}
function Get-VMEffectiveToolsOS {
    param($vm)

    if (-not $script:VMToolsOSCache) { $script:VMToolsOSCache = @{} }
    $vmKey = $null
    try { $vmKey = $vm.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($vmKey)) { try { $vmKey = $vm.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($vmKey)) { $vmKey = [string]$vm.Name }

    if ($script:VMToolsOSCache.ContainsKey($vmKey)) { return $script:VMToolsOSCache[$vmKey] }

    $osName = $null
    try { if ($vm.Guest -and -not [string]::IsNullOrWhiteSpace($vm.Guest.OSFullName)) { $osName = $vm.Guest.OSFullName } } catch {}
    if (-not $osName) { try { if ($vm.ExtensionData.Guest -and -not [string]::IsNullOrWhiteSpace($vm.ExtensionData.Guest.GuestFullName)) { $osName = $vm.ExtensionData.Guest.GuestFullName } } catch {} }
    if (-not $osName) { try { if ($vm.ExtensionData.Config -and -not [string]::IsNullOrWhiteSpace($vm.ExtensionData.Config.GuestFullName)) { $osName = $vm.ExtensionData.Config.GuestFullName } } catch {} }

    $script:VMToolsOSCache[$vmKey] = $osName
    return $osName
}

function Get-PlatformToolsRequiredVersion {
    param($VMHosts, $VMs)

    if ($script:PlatformToolsRequiredVersionResolved) {
        return $script:PlatformToolsRequiredVersion
    }

    $script:PlatformToolsRequiredVersionResolved = $true
    $script:PlatformToolsRequiredVersion = $null

    # Primary method: map ESXi host build(s) to the VMware Tools version exposed by the platform.
    # vCenter compares guest Tools against the Tools payload on the host (/productLocker),
    # so the host build is the best platform-side anchor for a cached Required Version value.
    try {
        $hostFacts = @($VMHosts | Where-Object { $_ -and $_.Version -and $_.Build } | ForEach-Object {
            [pscustomobject]@{
                Version = [string]$_.Version
                Build   = [int64]$_.Build
            }
        })

        if ($hostFacts.Count -gt 0) {
            $raw = Invoke-RestMethod -Uri 'https://packages.vmware.com/tools/versions' -Method Get -ErrorAction Stop

            $rows = foreach ($line in (($raw -split "`n") | Where-Object { $_ -and $_ -notmatch '^\s*#' })) {
                $norm = (($line -replace "`t", ' ') -replace '\s+', ' ').Trim()
                if (-not $norm) { continue }

                $parts = $norm -split ' '
                if ($parts.Count -lt 4) { continue }

                $ngcVersion = $parts[0]
                $esxToken   = $parts[1]
                $esxBuild   = $parts[2]
                $toolsVer   = $parts[3]

                if ($esxToken -notmatch '^ESX/(?<ver>.+)$') { continue }
                if ($esxBuild -notmatch '^\d+$') { continue }
                if ($toolsVer -notmatch '^\d+$') { continue }

                [pscustomobject]@{
                    EsxVersion  = $Matches['ver']
                    EsxBuild    = [int64]$esxBuild
                    ToolsVersion = [int]$toolsVer
                    NgcVersion  = $ngcVersion
                }
            }

            if ($rows) {
                $resolvedVersions = foreach ($hf in $hostFacts) {
                    $hostVersion = $hf.Version
                    $hostBuild   = $hf.Build

                    $candidates = @(
                        $rows | Where-Object {
                            $rowVer = [string]$_.EsxVersion
                            $hostVersion -eq $rowVer -or
                            $hostVersion.StartsWith("$rowVer.")
                        } | Sort-Object -Property EsxBuild
                    )

                    if (-not $candidates -or $candidates.Count -eq 0) {
                        $majorMinor = (($hostVersion -split '\.') | Select-Object -First 2) -join '.'
                        $candidates = @(
                            $rows | Where-Object {
                                $rowVer = [string]$_.EsxVersion
                                $rowVer -eq $majorMinor -or
                                $rowVer.StartsWith("$majorMinor.")
                            } | Sort-Object -Property EsxBuild
                        )
                    }

                    if ($candidates -and $candidates.Count -gt 0) {
                        $match = $candidates | Where-Object { $hostBuild -ge $_.EsxBuild } | Select-Object -Last 1
                        if (-not $match) {
                            $match = $candidates | Select-Object -First 1
                        }
                        if ($match -and $match.ToolsVersion) {
                            [int]$match.ToolsVersion
                        }
                    }
                }

                if ($resolvedVersions) {
                    $script:PlatformToolsRequiredVersion = [string](($resolvedVersions | Group-Object | Sort-Object -Property Count, Name -Descending | Select-Object -First 1).Name)
                }
            }
        }
    }
    catch {}

    # Fallback: if the host-build mapping cannot be resolved, use the most common config-side Tools version seen in inventory.
    # This is only a fallback and should be overridden by the host-build lookup above when available.
    if (-not $script:PlatformToolsRequiredVersion) {
        try {
            $fallback = @(
                $VMs | ForEach-Object {
                    try { $_.ExtensionData.Config.Tools.ToolsVersion } catch {}
                } | Where-Object { $_ -ne $null -and $_ -match '^\d+$' }
            )

            if ($fallback) {
                $script:PlatformToolsRequiredVersion = [string](($fallback | Group-Object | Sort-Object -Property Count, Name -Descending | Select-Object -First 1).Name)
            }
        }
        catch {}
    }

    return $script:PlatformToolsRequiredVersion
}

function Get-vInfo {
    param($vms, $about, $GetVMContextFn)
    $total = $vms.Count
    $i = 0
    foreach ($vm in $vms) {
        $i++
        Write-InlineProgress -Activity "Collecting vInfo $i of $total VMs" -PercentComplete ([int](($i / $total) * 100)) -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null

        $ed = $vm.ExtensionData
        $ctx = & $GetVMContextFn $vm
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        $boot = $ed.Config.BootOptions
        $videoCard = $ed.Config.Hardware.Device | Where-Object {
            $typeName = $_.GetType().Name
            $typeName -match 'Video|SVGA' -or
            ($_.PSObject.Properties.Name -contains 'VideoRamSizeInKB') -or
            ($_.DeviceInfo -and $_.DeviceInfo.Label -match '^Video')
        } | Select-Object -First 1
        $nics = @($ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualEthernetCard] })
        $disks = @($ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] })
        $networkNames = @($nics | ForEach-Object {
            if ($_.Backing -is [VMware.Vim.VirtualEthernetCardDistributedVirtualPortBackingInfo]) {
                $_.DeviceInfo.Summary
            }
            elseif ($_.Backing -is [VMware.Vim.VirtualEthernetCardNetworkBackingInfo]) {
                $_.Backing.DeviceName
            }
            else {
                $_.DeviceInfo.Summary
            }
        })
        $primaryNic = $nics | Select-Object -First 1
        $primaryGuestNet = if ($primaryNic) { Get-VMGuestNetworkDetails -VM $vm -Nic $primaryNic } else { $null }
        $primaryIP = if (-not [string]::IsNullOrWhiteSpace($ed.Summary.Guest.IpAddress)) { $ed.Summary.Guest.IpAddress } elseif ($primaryGuestNet) { $primaryGuestNet.IPv4 } else { $null }
        $dnsName = if (-not [string]::IsNullOrWhiteSpace($ed.Guest.HostName)) { $ed.Guest.HostName } elseif ($primaryGuestNet) { $primaryGuestNet.DnsName } else { $null }
        $clusterRules = Get-VMRuleSummary -VM $vm
        $dasVmSettings = $ed.Config.DasVmSettings
        $enableUuid = Get-VMAdvancedSettingValue -VM $vm -Name 'disk.EnableUUID'
        $opNotificationTimeout = Get-VMAdvancedSettingValue -VM $vm -Name 'tools.guestlib.enableHostInfo'
        $resourceConfig = $ed.ResourceConfig
        $storage = $ed.Summary.Storage

        [PSCustomObject]@{
            Name                                     = $vm.Name
            Powerstate                               = $effectivePowerState
            Template                                 = $ed.Config.Template
            "SRM Placeholder"                        = $false
            "Config status"                          = $ed.OverallStatus
            "DNS Name"                               = $dnsName
            "Connection state"                       = $ed.Runtime.ConnectionState
            "Guest state"                            = $ed.Guest.GuestState
            Heartbeat                                = $ed.GuestHeartbeatStatus
            "Consolidation Needed"                   = $ed.Runtime.ConsolidationNeeded
            PowerOn                                  = ($effectivePowerState -match 'PoweredOn')
            "Suspended To Memory"                    = (Get-SafeValue -Value $ed.Runtime.SuspendedToMemory)
            "Suspend time"                           = $ed.Runtime.SuspendTime
            "Suspend Interval"                       = (Get-SafeValue -Value $ed.Runtime.SuspendInterval)
            "Creation date"                          = $ed.Config.CreateDate
            ChangeVersion                            = $ed.Config.ChangeVersion
            CPUs                                     = $ed.Config.Hardware.NumCPU
            "Overall Cpu Readiness"                  = (Get-SafeValue -Value $ed.Summary.QuickStats.OverallCpuReadiness)
            Memory                                   = $ed.Config.Hardware.MemoryMB
            "Active Memory"                          = $ed.Summary.QuickStats.GuestMemoryUsage
            NICs                                     = $nics.Count
            Disks                                    = $disks.Count
            "Total disk capacity MiB"                = [int](($disks | Measure-Object -Property CapacityInKB -Sum).Sum / 1024)
            "Fixed Passthru HotPlug"                 = (Get-SafeValue -Value $ed.Config.FixedPassthruHotPlugEnabled)
            "min Required EVC Mode Key"              = (Get-SafeValue -Value $ed.Runtime.MinRequiredEVCModeKey)
            "Latency Sensitivity"                    = $ed.Config.LatencySensitivity.Level
            "Op Notification Timeout"                = $opNotificationTimeout
            EnableUUID                               = $enableUuid
            CBT                                      = $ed.Config.ChangeTrackingEnabled
            "Primary IP Address"                     = $primaryIP
            "Network #1"                             = (Get-SafeValue -Value $networkNames[0] -Default '')
            "Network #2"                             = (Get-SafeValue -Value $networkNames[1] -Default '')
            "Network #3"                             = (Get-SafeValue -Value $networkNames[2] -Default '')
            "Network #4"                             = (Get-SafeValue -Value $networkNames[3] -Default '')
            "Network #5"                             = (Get-SafeValue -Value $networkNames[4] -Default '')
            "Network #6"                             = (Get-SafeValue -Value $networkNames[5] -Default '')
            "Network #7"                             = (Get-SafeValue -Value $networkNames[6] -Default '')
            "Network #8"                             = (Get-SafeValue -Value $networkNames[7] -Default '')
            "Num Monitors"                           = (Get-SafeValue -Value $videoCard.NumDisplays)
            "Video Ram KiB"                          = if ($videoCard) { [int]($videoCard.VideoRamSizeInKB) } else { $null }
            "Resource pool"                          = $ctx.ResourcePool
            "Folder ID"                              = (Get-SafeValue -Value $vm.Folder.Id)
            Folder                                   = $ctx.Folder
            vApp                                     = $ctx.vApp
            "DAS protection"                         = (Get-SafeValue -Value $ed.Runtime.DasVmProtection.DasProtected)
            "FT State"                               = (Get-SafeValue -Value $ed.Runtime.FaultToleranceState)
            "FT Role"                                = (Get-SafeValue -Value $ed.Runtime.RecordReplayState)
            "FT Latency"                             = (Get-SafeValue -Value $ed.Runtime.FtLatencyStatus)
            "FT Bandwidth"                           = (Get-SafeValue -Value $ed.Runtime.FtSecondaryLatency)
            "FT Sec. Latency"                        = (Get-SafeValue -Value $ed.Runtime.FtSecondaryLatency)
            "Vm Failover In Progress"                = (Get-SafeValue -Value $ed.Runtime.DasVmProtection.DasVmProtectionState)
            "Provisioned MiB"                        = if ($storage) { [math]::Round($storage.Committed / 1MB, 0) + [math]::Round($storage.Uncommitted / 1MB, 0) } else { $null }
            "In Use MiB"                             = if ($storage) { [math]::Round($storage.Committed / 1MB, 0) } else { $null }
            "Unshared MiB"                           = if ($storage) { [math]::Round($storage.Unshared / 1MB, 0) } else { $null }
            "HA Restart Priority"                    = (Get-SafeValue -Value $dasVmSettings.RestartPriority)
            "HA Isolation Response"                  = (Get-SafeValue -Value $dasVmSettings.IsolationResponse)
            "HA VM Monitoring"                       = (Get-SafeValue -Value $dasVmSettings.VmMonitoring)
            "Cluster rule(s)"                        = $clusterRules.EnabledRuleNames
            "Cluster rule name(s)"                   = $clusterRules.RuleNames
            "Boot Required"                          = $boot.EnterBIOSSetup
            "Boot delay"                             = $boot.BootDelay
            "Boot retry delay"                       = $boot.BootRetryDelay
            "Boot retry enabled"                     = $boot.BootRetryEnabled
            "Boot BIOS setup"                        = $boot.EnterBIOSSetup
            "Reboot PowerOff"                        = $boot.BootOrder
            "EFI Secure boot"                        = (Get-SafeValue -Value $ed.Config.BootOptions.EfiSecureBootEnabled)
            Firmware                                 = $ed.Config.Firmware
            "HW version"                             = $ed.Config.Version
            "HW upgrade status"                      = (Get-SafeValue -Value $ed.Config.ScheduledHardwareUpgradeInfo.Status)
            "HW upgrade policy"                      = (Get-SafeValue -Value $ed.Config.ScheduledHardwareUpgradeInfo.UpgradePolicy)
            "HW target"                              = (Get-SafeValue -Value $ed.Config.ScheduledHardwareUpgradeInfo.VersionKey)
            Path                                     = $ed.Config.Files.VmPathName
            "Log directory"                          = $ed.Config.Files.LogDirectory
            "Snapshot directory"                     = $ed.Config.Files.SnapshotDirectory
            "Suspend directory"                      = $ed.Config.Files.SuspendDirectory
            Annotation                               = $ed.Config.Annotation
            "com.vrlcm.snapshot"                     = $null
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $runtimeHostName
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $toolsOsFullName
            "Customization Info"                     = (Get-SafeValue -Value $ed.Guest.CustomizationInfo)
            "Guest Detailed Data"                    = $ed.Guest.GuestDetailedData
            "VM ID"                                  = $ed.MoRef.Value
            "SMBIOS UUID"                            = $ed.Config.InstanceUuid
            "VM UUID"                                = $ed.Config.Uuid
            "VI SDK Server type"                     = $about.ApiType
            "VI SDK API Version"                     = $about.ApiVersion
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        $toolsStatus2 = $ed.Guest.ToolsVersionStatus2
        $requiredVersion = $PlatformToolsRequiredVersion
        if ($toolsStatus2 -eq 'guestToolsUnmanaged') {
            $requiredVersion = 'Unmanaged'
        }
        [PSCustomObject]@{
            Name                                     = $vm.Name
            Powerstate                               = $effectivePowerState
            Template                                 = $ed.Config.Template
            "SRM Placeholder"                        = $false
            Annotation                               = $ed.Config.Annotation
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $runtimeHostName
            Folder                                   = $ctx.Folder
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $toolsOsFullName
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            vCPUs                                    = $ed.Config.Hardware.NumCPU
            Sockets                                  = $ed.Config.Hardware.NumCPU / $ed.Config.Hardware.NumCoresPerSocket
            CoresPerSocket                           = $ed.Config.Hardware.NumCoresPerSocket
            MaxCPU_MHz                               = (Get-SafeValue -Value $ed.Runtime.MaxCpuUsage -Default 0)
            CPU_Usage_MHz                            = $ed.Summary.QuickStats.OverallCpuUsage
            SharesLevel                              = $ed.Config.CpuAllocation.Shares.Level
            Shares                                   = $ed.Config.CpuAllocation.Shares.Shares
            CPU_Reservation                          = $ed.Config.CpuAllocation.Reservation
            CPULimit                                 = if ($ed.Config.CpuAllocation.Limit -eq -1) { 0 } else { $ed.Config.CpuAllocation.Limit }
            EntitlementMHz                           = $ed.Summary.QuickStats.StaticCpuEntitlement
            DRSEntitlementMHz                        = $ed.Summary.QuickStats.DistributedCpuEntitlement
            CPUHotAdd                                = $ed.Config.CpuHotAddEnabled
            CPUHotRemove                             = $ed.Config.CpuHotRemoveEnabled
            NumaHotaddExposed                        = Get-VMAdvancedSettingValue -VM $vm -Name 'numa.autosize.vcpu.maxPerVirtualNode'
            "com.vrlcm.snapshot"                     = Get-VMAdvancedSettingValue -VM $vm -Name 'com.vrlcm.snapshot'
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        $toolsStatus2 = $null
        try { $toolsStatus2 = $ed.Guest.ToolsVersionStatus2 } catch {}
        $requiredVersion = $PlatformToolsRequiredVersion
        if ($toolsStatus2 -eq 'guestToolsUnmanaged') {
            $requiredVersion = 'Unmanaged'
        }
        [PSCustomObject]@{
            Name                                     = $vm.Name
            Powerstate                               = $effectivePowerState
            Template                                 = $ed.Config.Template
            "SRM Placeholder"                        = $false
            Annotation                               = $ed.Config.Annotation
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $runtimeHostName
            Folder                                   = $ctx.Folder
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $toolsOsFullName
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            MemoryMB                                 = $ed.Config.Hardware.MemoryMB
            MemoryReservationLockedToMax             = $ed.Config.MemoryReservationLockedToMax
            Overhead                                 = (Get-SafeValue -Value $ed.Config.MemoryAllocation.OverheadLimit -Default 0)
            MaxMemoryUsageMB                         = (Get-SafeValue -Value $ed.Runtime.MaxMemoryUsage -Default 0)
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
            "com.vrlcm.snapshot"                     = Get-VMAdvancedSettingValue -VM $vm -Name 'com.vrlcm.snapshot'
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        foreach ($disk in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualDisk] }) {
            $ctl = $ed.Config.Hardware.Device | Where-Object { $_.Key -eq $disk.ControllerKey }
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Powerstate                               = $effectivePowerState
                Template                                 = $ed.Config.Template
                "SRM Placeholder"                        = $false
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $runtimeHostName
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $toolsOsFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                Disk                                     = $disk.DeviceInfo.Label
                DiskKey                                  = $disk.Key
                DiskUUID                                 = $disk.Backing.Uuid
                CapacityMB                               = [int]($disk.CapacityInKB / 1024)
                Thin                                     = $disk.Backing.ThinProvisioned
                EagerZero                                = $disk.Backing.EagerlyScrub
                Mode                                     = $disk.Backing.DiskMode
                SharingMode                              = $disk.Sharing
                Split                                    = $disk.Backing.Split
                WriteThrough                             = $disk.Backing.WriteThrough
                Level                                    = $disk.StorageIOAllocation.Shares.Level
                Shares                                   = $disk.StorageIOAllocation.Shares.Shares
                Reservation                              = $disk.StorageIOAllocation.Reservation
                Limit                                    = $disk.StorageIOAllocation.Limit
                Controller                               = $ctl.GetType().Name
                Label                                    = $ctl.DeviceInfo.Label
                ControllerBus                            = $ctl.BusNumber
                Unit                                     = $disk.UnitNumber
                SharedBus                                = $ctl.SharedBus
                Path                                     = $disk.Backing.FileName
                Raw                                      = ($disk.Backing -is [VMware.Vim.VirtualDiskRawDiskMappingVer1BackingInfo])
                LunUuid                                  = $disk.Backing.LunUuid
                RDMMode                                  = $disk.Backing.CompatibilityMode
                InternalSortColumn                       = ('{0:D2}:{1:D2}' -f [int]$ctl.BusNumber,[int]$disk.UnitNumber)
                "com.vrlcm.snapshot"                     = Get-VMAdvancedSettingValue -VM $vm -Name 'com.vrlcm.snapshot'
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm

        $guestDiskMap = @{}
        try {
            if (-not (Test-IsTemplateLikeObject -vm $vm)) {
                foreach ($guestDisk in (Get-VMGuestDisk -VM $vm -ErrorAction SilentlyContinue)) {
                    $matchedHardDisk = Get-HardDisk -VMGuestDisk $guestDisk -ErrorAction SilentlyContinue
                    if ($guestDisk.DiskPath) {
                        $guestDiskMap[$guestDisk.DiskPath] = $matchedHardDisk
                    }
                }
            }
        }
        catch { }

        if ($ed.Guest.Disk) {
            foreach ($gdisk in $ed.Guest.Disk) {
                $matchedDisk = $null
                if ($gdisk.DiskPath -and $guestDiskMap.ContainsKey($gdisk.DiskPath)) {
                    $matchedDisk = $guestDiskMap[$gdisk.DiskPath]
                }

                [PSCustomObject]@{
                    Name                                     = $vm.Name
                    Powerstate                               = $effectivePowerState
                    Template                                 = $ed.Config.Template
                    "SRM Placeholder"                        = $false
                    DiskKey                                  = if ($matchedDisk) { $matchedDisk.ExtensionData.Key } else { $null }
                    Disk                                     = $gdisk.DiskPath
                    CapacityMB                               = [math]::Round($gdisk.Capacity / 1MB, 0)
                    ConsumedMB                               = [math]::Round((($gdisk.Capacity - $gdisk.FreeSpace) / 1MB), 0)
                    FreeMB                                   = [math]::Round($gdisk.FreeSpace / 1MB, 0)
                    FreePct                                  = if ($gdisk.Capacity -gt 0) { [math]::Round(($gdisk.FreeSpace / $gdisk.Capacity * 100), 0) } else { 0 }
                    InternalSortColumn                       = $gdisk.DiskPath
                    Annotation                               = $ed.Config.Annotation
                    "com.vrlcm.snapshot"                     = $null
                    Datacenter                               = $ctx.Datacenter
                    Cluster                                  = $ctx.Cluster
                    Host                                     = $runtimeHostName
                    Folder                                   = $ctx.Folder
                    "OS according to the configuration file" = $ed.Config.GuestFullName
                    "OS according to the VMware Tools"       = $toolsOsFullName
                    VM_ID                                    = $ed.MoRef.Value
                    UUID                                     = $ed.Config.Uuid
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
       if ($ed.guest.guestfamily -like "*windowsGuest") {
            if($($VM | Get-AdvancedSetting -name "disk.EnableUUID") -like "*TRUE*"){
                $diskData = $true
            }else{
                $diskData = $false
            }
       }
       Else{$diskData=$true}
       If($diskData -and -not (Test-IsTemplateLikeObject -vm $vm)){
            foreach ($gDisk in (Get-VMGuestDisk -VM $vm -ErrorAction SilentlyContinue)) {

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
                    Host                                     = $runtimeHostName
                    Folder                                   = $ctx.Folder
                    "OS according to the configuration file" = $ed.Config.GuestFullName
                    "OS according to the VMware Tools"       = $toolsOsFullName
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        foreach ($nic in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualEthernetCard] }) {
            $networkName = $null
            $switchName = $null
            if ($nic.Backing -is [VMware.Vim.VirtualEthernetCardDistributedVirtualPortBackingInfo]) {
                $port = $nic.Backing.Port
                try {
                    $pgView = Get-View -Id $port.PortgroupKey -ErrorAction Stop
                    $networkName = $pgView.Name
                    $switchName = (Get-View -Id $pgView.Config.DistributedVirtualSwitch -ErrorAction Stop).Name
                }
                catch {
                    $networkName = $nic.DeviceInfo.Summary
                    $switchName = $nic.DeviceInfo.Summary
                }
            }
            elseif ($nic.Backing -is [VMware.Vim.VirtualEthernetCardNetworkBackingInfo]) {
                $networkName = $nic.Backing.DeviceName
                $switchName = $nic.Backing.DeviceName
            }
            else {
                $networkName = $nic.DeviceInfo.Summary
                $switchName = $nic.DeviceInfo.Summary
            }

            $guestNet = Get-VMGuestNetworkDetails -VM $vm -Nic $nic

            [PSCustomObject]@{
                Name                                     = $vm.Name
                Powerstate                               = $effectivePowerState
                Template                                 = $ed.Config.Template
                "SRM Placeholder"                        = $false
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $runtimeHostName
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $toolsOsFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                Adapter                                  = $nic.DeviceInfo.Label
                Network                                  = $networkName
                Switch                                   = $switchName
                MAC                                      = $nic.MacAddress
                Type                                     = $nic.GetType().Name
                Connected                                = $nic.Connectable.Connected
                StartConnected                           = $nic.Connectable.StartConnected
                "IPv4 Address"                           = $guestNet.IPv4
                "IPv6 Address"                           = $guestNet.IPv6
                "Direct Path IO"                         = (($nic.GetType().Name -match 'Sriov|Passthrough|DirectPath') -or (($null -ne $nic.Backing) -and ($nic.Backing.GetType().Name -match 'Sriov|Passthrough|DirectPath')) -or (($null -ne $nic.Connectable) -and ($nic.Connectable.GetType().Name -match 'Sriov|Passthrough|DirectPath')))
                "Internal Sort Column"                   = $nic.UnitNumber
                "com.vrlcm.snapshot"                     = $null
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        foreach ($fl in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualFloppy] }) {
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Annotation                               = $ed.Config.Annotation
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $runtimeHostName
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $toolsOsFullName
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        foreach ($cd in $ed.Config.Hardware.Device | Where-Object { $_ -is [VMware.Vim.VirtualCdrom] }) {
            [PSCustomObject]@{
                Name                                     = $vm.Name
                Powerstate                               = $effectivePowerState
                Template                                 = $ed.Config.Template
                "SRM Placeholder"                        = $false
                DeviceNode                               = $cd.DeviceInfo.Label
                Connected                                = $cd.Connectable.Connected
                StartConnected                           = $cd.Connectable.StartConnected
                Type                                     = $cd.Backing.GetType().Name
                Annotation                               = $ed.Config.Annotation
                "com.vrlcm.snapshot"                     = $null
                Datacenter                               = $ctx.Datacenter
                Cluster                                  = $ctx.Cluster
                Host                                     = $runtimeHostName
                Folder                                   = $ctx.Folder
                "OS according to the configuration file" = $ed.Config.GuestFullName
                "OS according to the VMware Tools"       = $toolsOsFullName
                VM_ID                                    = $ed.MoRef.Value
                UUID                                     = $ed.Config.Uuid
                VMRef                                    = $ed.MoRef.Value
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm
        if ($ed.Snapshot) {
            foreach ($snap in Get-Snapshot -vm $vm ) {
                
                [PSCustomObject]@{
                    Name                                     = $vm.Name
                    Powerstate                               = $effectivePowerState
                    SnapshotName                             = $snap.Name
                    Description                              = $snap.Description
                    Created                                  = $snap.Created
                    Filename                                 = if ($snap.ExtensionData.Config.FileName) { $snap.ExtensionData.Config.FileName } elseif ($snap.ExtensionData.ReplaySupported -ne $null) { ([string]$snap.ExtensionData) } else { "$($ed.LayoutEx.Snapshot | Select-Object -ExpandProperty Id -First 1)" }
                    SizeMiBVMSN                              = if ($snap.SizeGB -ne $null) { [math]::Round($snap.SizeGB * 1024,0) } else { $null }
                    SizeMiBTotal                             = if ($snap.SizeGB -ne $null) { [math]::Round($snap.SizeGB * 1024,0) } else { $null }
                    Quiesced                                 = $snap.Quiesced
                    State                                    = if ($snap.IsCurrent) { 'current' } else { 'notCurrent' }
                    Annotation                               = $ed.Config.Annotation
                    "com.vrlcm.snapshot"                     = $null
                    Datacenter                               = $ctx.Datacenter
                    Cluster                                  = $ctx.Cluster
                    Host                                     = $runtimeHostName
                    Folder                                   = $ctx.Folder
                    "OS according to the configuration file" = $ed.Config.GuestFullName
                    "OS according to the VMware Tools"       = $toolsOsFullName
                    VM_ID                                    = $ed.MoRef.Value
                    UUID                                     = $ed.Config.Uuid
                    "VI SDK Server"                          = $about.FullName
                    "VI SDK UUID"                            = $about.InstanceUuid
                }
            }
        }
    }
    Write-InlineProgress -Activity 'vSnapshot Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}

function Get-vTools {
    param($vms, $about, $GetVMContextFn, $PlatformToolsRequiredVersion)
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
        $runtimeHostName = Get-VMRuntimeHostName -VM $vm
        $effectivePowerState = Get-VMEffectivePowerState -VM $vm
        $toolsOsFullName = Get-VMEffectiveToolsOS -VM $vm

        $guestExt = $null
        try { $guestExt = $vm.Guest.ExtensionData } catch {}
        if (-not $guestExt) {
            try { $guestExt = $ed.Guest } catch {}
        }

        $toolsStatus2 = $null
        try { $toolsStatus2 = $guestExt.ToolsVersionStatus2 } catch {}

        $toolsVersionNumber = $null
        try { $toolsVersionNumber = $guestExt.ToolsVersion } catch {}

        $toolsInstallType = $null
        try { $toolsInstallType = $guestExt.ToolsInstallType } catch {}

        $requiredVersion = $PlatformToolsRequiredVersion
        if ($toolsStatus2 -eq 'guestToolsUnmanaged') {
            $requiredVersion = 'Unmanaged'
        }

        [PSCustomObject]@{
            Name                                     = $vm.Name
            Powerstate                               = $effectivePowerState
            Template                                 = $ed.Config.Template
            "SRM Placeholder"                        = $false
            Annotation                               = $ed.Config.Annotation
            Datacenter                               = $ctx.Datacenter
            Cluster                                  = $ctx.Cluster
            Host                                     = $runtimeHostName
            Folder                                   = $ctx.Folder
            "OS according to the configuration file" = $ed.Config.GuestFullName
            "OS according to the VMware Tools"       = $toolsOsFullName
            VM_ID                                    = $ed.MoRef.Value
            UUID                                     = $ed.Config.Uuid
            VMVersion                                = $ed.Config.Version -replace "vmx-", ""
            Tools                                    = $ed.Guest.ToolsStatus
            ToolsVersion                             = $toolsVersionNumber
            ToolsStatus2                             = $toolsStatus2
            ToolsInstallType                         = $toolsInstallType
            RequiredVersion                          = $requiredVersion
            SyncTime                                 = $ed.Config.Tools.SyncTimeWithHost
            AppStatus                                = $ed.Guest.AppState
            AppHeartbeat                             = $ed.Guest.AppHeartbeatStatus
            KernelCrash                              = $ed.Guest.GuestKernelCrashed
            OpsReady                                 = $ed.Guest.GuestOperationsReady
            InteractiveReady                         = $ed.Guest.InteractiveGuestOperationsReady
            StateChangeSupported                     = $ed.Guest.GuestStateChangeSupported
            ToolsUpgradePolicy                       = $ed.Config.Tools.ToolsUpgradePolicy
            "com.vrlcm.snapshot"                     = Get-VMAdvancedSettingValue -VM $vm -Name 'com.vrlcm.snapshot'
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
            ResourcePoolPath                  = $rp.ExtensionData.Parent.Value
            Status                            = $ed.OverallStatus
            NumVMsTotal                       = $vmsInPool.Count
            NumVMs                            = $vmsInPool.Count
            vCPUs                             = ($vmsInPool | Measure-Object -Property NumCPU -Sum).Sum
            CPU_Limit                         = $ed.Config.CpuAllocation.Limit
            CPU_OverheadLimit                 = $null
            CPU_Reservation                   = $ed.Config.CpuAllocation.Reservation
            CPU_SharesLevel                   = $ed.Config.CpuAllocation.Shares.Level
            CPU_Shares                        = $ed.Config.CpuAllocation.Shares.Shares
            CPU_Expandable                    = $ed.Config.CpuAllocation.ExpandableReservation
            CPU_maxUsage                      = $ed.Runtime.Cpu.MaxUsage
            CPU_overallUsage                  = $ed.Runtime.Cpu.OverallUsage
            CPU_reservationUsed               = $ed.Runtime.Cpu.ReservationUsed
            CPU_reservationUsedForVm          = $ed.Runtime.Cpu.ReservationUsedForVm
            CPU_unreservedForPool             = $ed.Runtime.Cpu.UnreservedForPool
            CPU_unreservedForVm               = $ed.Runtime.Cpu.UnreservedForVm
            MemConfigured                     = ($vmsInPool | Measure-Object -Property MemoryMB -Sum).Sum
            Mem_Limit                         = $ed.Config.MemoryAllocation.Limit
            Mem_OverheadLimit                 = $null
            Mem_Reservation                   = $ed.Config.MemoryAllocation.Reservation
            Mem_SharesLevel                   = $ed.Config.MemoryAllocation.Shares.Level
            Mem_Shares                        = $ed.Config.MemoryAllocation.Shares.Shares
            Mem_Expandable                    = $ed.Config.MemoryAllocation.ExpandableReservation
            "Mem maxUsage"                    = $ed.Runtime.Memory.MaxUsage
            "Mem overallUsage"                = $ed.Runtime.Memory.OverallUsage
            "Mem reservationUsed"             = $ed.Runtime.Memory.ReservationUsed
            "Mem reservationUsedForVm"        = $ed.Runtime.Memory.ReservationUsedForVm
            "Mem unreservedForPool"           = $ed.Runtime.Memory.UnreservedForPool
            "Mem unreservedForVm"             = $ed.Runtime.Memory.UnreservedForVm
            ObjectID                          = $ed.MoRef.Value
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
        $vmMonSettings = if ($dasDef -and $dasDef.PSObject.Properties.Name -contains 'VmToolsMonitoringSettings') { $dasDef.VmToolsMonitoringSettings } elseif ($dasDef -and $dasDef.PSObject.Properties.Name -contains 'VmMonitoringSettings') { $dasDef.VmMonitoringSettings } else { $null }
        $drs = $ed.Configuration.DrsConfig
        $dpm = if ($ed.ConfigurationEx -and $ed.ConfigurationEx.PSObject.Properties.Name -contains 'DpmConfigInfo' -and $ed.ConfigurationEx.DpmConfigInfo) { $ed.ConfigurationEx.DpmConfigInfo } else { $ed.Configuration.DpmConfig }
        [PSCustomObject]@{
            Name                            = $cl.Name
            "Config status"                 = $ed.OverallStatus
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
            "Cluster Settings"              = $ed.ConfigurationEx.DasConfig.Enabled
            "Max Failures"                  = if ($vmMonSettings) { $vmMonSettings.MaxFailures } else { $null }
            "Max Failure Window"            = if ($vmMonSettings) { $vmMonSettings.MaxFailureWindow } else { $null }
            "Failure Interval"              = if ($vmMonSettings) { $vmMonSettings.FailureInterval } else { $null }
            "Min Up Time"                   = if ($vmMonSettings) { $vmMonSettings.MinUpTime } else { $null }
            "VM Monitoring"                 = $das.VmMonitoring
            "DRS enabled"                   = $drs.Enabled
            "DRS default VM behavior"       = $drs.DefaultVmBehavior
            "DRS vmotion rate"              = $drs.VmotionRate
            "DPM enabled"                   = if ($dpm.PSObject.Properties.Name -contains 'Enabled') { $dpm.Enabled } else { $null }
            "DPM default behavior"          = if ($dpm.PSObject.Properties.Name -contains 'DefaultDpmBehavior') { $dpm.DefaultDpmBehavior } elseif ($dpm.PSObject.Properties.Name -contains 'DefaultBehavior') { $dpm.DefaultBehavior } else { $null }
            "DPM Host Power Action Rate"    = if ($dpm.PSObject.Properties.Name -contains 'HostPowerActionRate') { $dpm.HostPowerActionRate } else { $null }
            ObjectID                        = $ed.MoRef.Value
            "com.vmware.vcenter.cluster.edrs.upgradeHostAdded" = (Get-AdvancedSetting -Entity $cl -Name 'com.vmware.vcenter.cluster.edrs.upgradeHostAdded' -ErrorAction SilentlyContinue).Value
            "com.vrlcm.snapshot"            = (Get-AdvancedSetting -Entity $cl -Name 'com.vrlcm.snapshot' -ErrorAction SilentlyContinue).Value
            "VI SDK Server"                 = $about.FullName
            "VI SDK UUID"                   = $about.InstanceUuid
        }
    }
    Write-InlineProgress -Activity 'vCluster Processed' -Complete -ProgressCharacter ([char]9632) -ProgressFillCharacter ([char]9632) -ProgressFill ([char]183) -BarBracketStart $null -BarBracketEnd $null
}


function Get-PolicyValue {
    param(
        [Parameter(Mandatory=$false)] $PolicyObject,
        [Parameter(Mandatory=$false)] $DefaultValue = $null
    )
    try {
        if ($null -eq $PolicyObject) { return $DefaultValue }
        if ($PolicyObject.PSObject.Properties.Name -contains 'Inherited') {
            if ($PolicyObject.Inherited -eq $true) { return $DefaultValue }
        }
        if ($PolicyObject.PSObject.Properties.Name -contains 'Value') {
            if ($null -ne $PolicyObject.Value) { return $PolicyObject.Value }
        }
        return $PolicyObject
    }
    catch {
        return $DefaultValue
    }
}

function Get-RemoteCertificateInfo {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [int]$Port = 443,
        [int]$TimeoutMs = 5000
    )

    $client = $null
    $sslStream = $null
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            throw "Connection timeout"
        }
        $client.EndConnect($async)

        $callback = { param($sender,$cert,$chain,$errors) return $true }
        $sslStream = New-Object System.Net.Security.SslStream($client.GetStream(), $false, $callback)
        $sslStream.AuthenticateAsClient($ComputerName)

        $cert2 = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 $sslStream.RemoteCertificate
        [pscustomobject]@{
            Subject    = $cert2.Subject
            Issuer     = $cert2.Issuer
            NotBefore  = $cert2.NotBefore
            NotAfter   = $cert2.NotAfter
            Thumbprint = $cert2.Thumbprint
            Status     = if ($cert2.NotAfter -lt (Get-Date)) { 'Expired' } else { 'Valid' }
        }
    }
    catch {
        $null
    }
    finally {
        if ($sslStream) { try { $sslStream.Dispose() } catch {} }
        if ($client) { try { $client.Close() } catch {} }
    }
}

function Get-HostContext {
    param([Parameter(Mandatory = $true)] $ESXhost)

    if (-not $script:HostContextCache) { $script:HostContextCache = @{} }

    $hostKey = $null
    try { $hostKey = $ESXhost.Id } catch {}
    if ([string]::IsNullOrWhiteSpace($hostKey)) { try { $hostKey = $ESXhost.ExtensionData.MoRef.Value } catch {} }
    if ([string]::IsNullOrWhiteSpace($hostKey)) { $hostKey = [string]$ESXhost.Name }

    if ($script:HostContextCache.ContainsKey($hostKey)) { return $script:HostContextCache[$hostKey] }

    $datacenter = $null
    $cluster = $null
    try { $datacenter = ($ESXhost | Get-Datacenter -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -First 1) } catch {}
    try { $cluster = ($ESXhost | Get-Cluster -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name -First 1) } catch {}

    $ctx = [PSCustomObject]@{ Datacenter = $datacenter; Cluster = $cluster }
    $script:HostContextCache[$hostKey] = $ctx
    return $ctx
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
        $hostVMs = @($ESXhost | Get-VM)
        $hostNameCfg = $ed.Config.Network.DnsConfig.HostName
        $domainCfg = $ed.Config.Network.DnsConfig.DomainName
        $dnsServers = @($ed.Config.Network.DnsConfig.Address)
        $dnsSearchOrder = @($ed.Config.Network.DnsConfig.SearchDomain)
        $ntpServers = @($ed.Config.DateTimeInfo.NtpConfig.Server)
        $timeZoneObj = $ed.Config.DateTimeInfo.TimeZone
        $licenseKey = ($ESXhost | Select-Object -ExpandProperty LicenseKey -ErrorAction SilentlyContinue | Select-Object -First 1)
        $cert = Get-RemoteCertificateInfo -ComputerName $ESXhost.Name
        if (-not $cert) { $cert = $ed.Config.Certificate }
        $ed = $ESXhost.ExtensionData
        $hostNetwork = Get-VMHostNetwork -VMHost $ESXhost -ErrorAction SilentlyContinue
        $mgmtVmk = Get-VMHostNetworkAdapter -VMKernel -VMHost $ESXhost -ErrorAction SilentlyContinue | Where-Object { $_.ManagementTrafficEnabled } | Select-Object -First 1
        $hostCtx = Get-HostContext -ESXhost $ESXhost
        [PSCustomObject]@{
            Name               = $ESXhost.Name
            Datacenter         = $hostCtx.Datacenter
            Cluster            = $hostCtx.Cluster
            Status             = $ed.OverallStatus
            ComplianceCheckState = $ed.Summary.ComplianceCheckState
            InMaintenanceMode  = $ed.Runtime.InMaintenanceMode
            InQuarantineMode   = $ed.Runtime.InQuarantineMode
            CPUModel           = $ed.Summary.Hardware.CpuModel
            SpeedMHz           = $ed.Summary.Hardware.CpuMhz
            HT_Available       = $ed.Config.HyperThread.Available
            HT_Active          = $ed.Config.HyperThread.Active
            CPUPackages        = $ed.Summary.Hardware.NumCpuPkgs
            CoresPerCPU        = $ed.Summary.Hardware.NumCpuCores / $ed.Summary.Hardware.NumCpuPkgs
            TotalCores         = $ed.Summary.Hardware.NumCpuCores
            CPUUsagePct        = [math]::Round(($ed.Summary.QuickStats.OverallCpuUsage / ($ed.Summary.Hardware.CpuMhz * $ed.Summary.Hardware.NumCpuCores) * 100), 1)
            MemoryGB           = [math]::Round($ed.Summary.Hardware.MemorySize / 1MB, 0)
            MemoryTieringType  = $ed.Summary.Hardware.MemoryTieringType
            MemoryUsagePct     = [math]::Round(($ed.Summary.QuickStats.OverallMemoryUsage / ($ed.Summary.Hardware.MemorySize / 1MB) * 100), 1)
            Console            = 0
            NumNICs            = $ed.Config.Network.Pnic.Count
            NumHBAs            = $ed.Config.StorageDevice.HostBusAdapter.Count
            NumVMs             = @($hostVMs | Where-Object { -not $_.ExtensionData.Config.Template }).Count
            NumVMsTotal        = $hostVMs.Count
            VMsPerCore         = [math]::Round((($hostVMs.Count / $ed.Summary.Hardware.NumCpuCores)), 2)
            NumvCPUs           = ($hostVMs | Measure-Object -Property NumCpu -Sum).Sum
            vCPUsPerCore       = if ($ed.Summary.Hardware.NumCpuCores) { [math]::Round((($hostVMs | Measure-Object -Property NumCpu -Sum).Sum / $ed.Summary.Hardware.NumCpuCores),2) } else { $null }
            vRAM               = ($hostVMs | Measure-Object -Property MemoryMB -Sum).Sum
            VMUsedMemory       = $ed.Summary.QuickStats.OverallMemoryUsage
            VMMemorySwapped    = $ed.Summary.QuickStats.OverallMemoryUsage
            VMMemoryBallooned  = 0
            VMotionSupport     = $ed.Capability.VmotionSupported
            StorageVMotionSupport = $ed.Capability.StorageVMotionSupported
            CurrentEVC         = $ed.Summary.CurrentEVCModeKey
            MaxEVC             = $ed.Summary.MaxEVCModeKey
            vSANFaultDomainName = $null
            AssignedLicenses   = $licenseKey
            ATSHeartbeat       = $ed.Config.Option | Where-Object { $_.Key -eq 'VMFS3.UseATSForHBOnVMFS5' } | Select-Object -ExpandProperty Value -First 1
            ATSLocking         = $ed.Config.Option | Where-Object { $_.Key -eq 'VMFS3.HardwareAcceleratedLocking' } | Select-Object -ExpandProperty Value -First 1
            CurrentCPUPowerManPolicy = $ed.Config.PowerSystemInfo.CurrentPolicy.ShortName
            SupportedCPUPowerMan = (@($ed.Config.PowerSystemInfo.AvailablePolicy) | ForEach-Object { $_.ShortName }) -join ', '
            HostPowerPolicy    = $ed.Config.PowerSystemInfo.CurrentPolicy.Key
            ESXiVersion        = $prod.FullName
            BootTime           = $ed.Summary.Runtime.BootTime
            DNSServers         = ($dnsServers -join ', ')
            DHCP               = $mgmtVmk.DhcpEnabled
            Domain             = $domainCfg
            DomainList         = ($dnsSearchOrder -join ', ')
            DNSSearchOrder     = ($dnsSearchOrder -join ', ')
            NTPServers         = ($ntpServers -join ', ')
            NTPDRunning        = ((Get-VMHostService -VMHost $ESXhost -ErrorAction SilentlyContinue | Where-Object { $_.Key -eq 'ntpd' } | Select-Object -ExpandProperty Running -First 1))
            TimeZone           = $timeZoneObj.Key
            TimeZoneName       = $timeZoneObj.Name
            GMTOffset          = $timeZoneObj.GmtOffset
            Vendor             = $sys.Vendor
            Model              = $sys.Model
            Serial             = $sys.SerialNumber
            ServiceTag         = $sys.OtherIdentifyingInfo | Where-Object { $_.IdentifierType.Key -match 'ServiceTag|AssetTag' } | Select-Object -ExpandProperty IdentifierValue -First 1
            OEMSpecificString  = ($sys.OtherIdentifyingInfo | ForEach-Object { $_.IdentifierValue }) -join ', '
            BIOSVendor         = $bios.Vendor
            BIOSVersion        = $bios.BiosVersion
            BIOSDate           = $bios.ReleaseDate
            CertificateIssuer  = $cert.Issuer
            CertificateStartDate = $cert.NotBefore
            CertificateExpiryDate = $cert.NotAfter
            CertificateStatus  = if ($cert.NotAfter -and $cert.NotAfter -lt (Get-Date)) { 'expired' } elseif ($cert.NotAfter) { 'good' } else { $null }
            CertificateSubject = $cert.Subject
            ObjectID           = $ed.MoRef.Value
            "AutoDeploy.MachineIdentity" = (Get-AdvancedSetting -Entity $ESXHost -Name 'AutoDeploy.MachineIdentity' -ErrorAction SilentlyContinue).Value
            UUID               = $ed.Hardware.SystemInfo.Uuid
            "com.vrlcm.snapshot" = (Get-AdvancedSetting -Entity $ESXHost -Name 'com.vrlcm.snapshot' -ErrorAction SilentlyContinue).Value
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
        $hostCtx = Get-HostContext -ESXhost $ESXhost
        $dcName = $hostCtx.Datacenter
        $clusterName = $hostCtx.Cluster
        foreach ($hba in $ed.Config.StorageDevice.HostBusAdapter) {
            [PSCustomObject]@{
                Host              = $ESXhost.Name
                Datacenter        = $dcName
                Cluster           = $clusterName
                HBADevice         = $hba.Device
                Model             = $hba.Model
                Type              = $hba.Model -replace ' Controller$',''
                Status            = $hba.Status
                Bus               = $hba.Bus
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
        $hostCtx = Get-HostContext -ESXhost $ESXhost
        $dcName = $hostCtx.Datacenter
        $clusterName = $hostCtx.Cluster
        foreach ($pnic in $ed.Config.Network.Pnic) {
            $switchName = ((Get-VirtualSwitch -VMHost $ESXhost -ErrorAction SilentlyContinue | Where-Object { $_.Nic -contains $pnic.Device } | Select-Object -ExpandProperty Name -First 1))
            [PSCustomObject]@{
                Host            = $ESXhost.Name
                Datacenter      = $dcName
                Cluster         = $clusterName
                PNICDevice      = $pnic.Device
                MAC             = $pnic.Mac
                LinkSpeed       = $pnic.LinkSpeed.SpeedMb
                Duplex          = $pnic.LinkSpeed.Duplex
                Driver          = $pnic.Driver
                Switch          = $switchName
                UplinkPort      = $null
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

        $hostCtx = Get-HostContext -ESXhost $ESXhost
        $dcName = $hostCtx.Datacenter
        $clusterName = $hostCtx.Cluster
        $hostEd = $ESXhost.ExtensionData
        foreach ($vs in Get-VirtualSwitch -VMHost $ESXhost) {
            $vsExt = $vs.ExtensionData
            $policy = $vsExt.Spec.Policy
            [PSCustomObject]@{
                Host              = $ESXhost.Name
                Datacenter        = $dcName
                Cluster           = $clusterName
                vSwitch           = $vs.Name
                NumPorts          = $vsExt.NumPorts
                NumPortsAvailable = $vsExt.NumPortsAvailable
                MTU               = $vs.Mtu
                Nic               = ($vs.Nic -join ',')
                ActiveNic         = $vsExt.Spec.Policy.NicTeaming.NicOrder.ActiveNic -join ','
                StandbyNic        = $vsExt.Spec.Policy.NicTeaming.NicOrder.StandbyNic -join ','
                AllowPromiscuous  = Get-PolicyValue -PolicyObject $policy.Security.AllowPromiscuous
                ForgedTransmits   = Get-PolicyValue -PolicyObject $policy.Security.ForgedTransmits
                MacChanges        = Get-PolicyValue -PolicyObject $policy.Security.MacChanges
                TrafficShaping    = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.Enabled
                Width             = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.AverageBandwidth -DefaultValue 0
                Peak              = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.PeakBandwidth -DefaultValue 0
                Burst             = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.BurstSize -DefaultValue 0
                Policy            = Get-PolicyValue -PolicyObject $policy.NicTeaming.Policy
                ReversePolicy     = Get-PolicyValue -PolicyObject $policy.NicTeaming.ReversePolicy
                NotifySwitch      = Get-PolicyValue -PolicyObject $policy.NicTeaming.NotifySwitches
                RollingOrder      = Get-PolicyValue -PolicyObject $policy.NicTeaming.RollingOrder
                Offload           = $true
                TSO               = if ($null -ne $hostEd.Capability.TsoSupported) { $hostEd.Capability.TsoSupported } else { $true }
                ZeroCopyXmit      = if ($null -ne $hostEd.Capability.ZeroCopyXmitSupported) { $hostEd.Capability.ZeroCopyXmitSupported } else { $true }
                CheckBeacon       = Get-PolicyValue -PolicyObject $policy.NicTeaming.FailureCriteria.CheckBeacon
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

        $hostCtx = Get-HostContext -ESXhost $ESXhost
        $dcName = $hostCtx.Datacenter
        $clusterName = $hostCtx.Cluster
        $hostEd = $ESXhost.ExtensionData
        $vSwitchIndex = @{}
        foreach ($vs in Get-VirtualSwitch -VMHost $ESXhost -ErrorAction SilentlyContinue) {
            $vSwitchIndex[$vs.Name] = $vs.ExtensionData.Spec.Policy
        }

        foreach ($pg in Get-VirtualPortGroup -VMHost $ESXhost) {
            $pgExt = $pg.ExtensionData
            $policy = $pgExt.Spec.Policy
            $switchPolicy = $null
            if ($vSwitchIndex.ContainsKey($pg.VirtualSwitchName)) { $switchPolicy = $vSwitchIndex[$pg.VirtualSwitchName] }

            [PSCustomObject]@{
                Host             = $ESXhost.Name
                Datacenter       = $dcName
                Cluster          = $clusterName
                PortGroup        = $pg.Name
                vSwitch          = $pg.VirtualSwitchName
                VLANId           = $pg.VlanId
                NumPorts         = $pgExt.NumPorts
                ActivePorts      = $pgExt.NumPortsActive
                AllowPromiscuous = Get-PolicyValue -PolicyObject $policy.Security.AllowPromiscuous -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.Security.AllowPromiscuous)
                ForgedTransmits  = Get-PolicyValue -PolicyObject $policy.Security.ForgedTransmits -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.Security.ForgedTransmits)
                MacChanges       = Get-PolicyValue -PolicyObject $policy.Security.MacChanges -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.Security.MacChanges)
                TrafficShaping   = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.Enabled -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.ShapingPolicy.Enabled)
                Width            = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.AverageBandwidth -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.ShapingPolicy.AverageBandwidth -DefaultValue 0)
                Peak             = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.PeakBandwidth -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.ShapingPolicy.PeakBandwidth -DefaultValue 0)
                Burst            = Get-PolicyValue -PolicyObject $policy.ShapingPolicy.BurstSize -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.ShapingPolicy.BurstSize -DefaultValue 0)
                Policy           = Get-PolicyValue -PolicyObject $policy.NicTeaming.Policy -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.NicTeaming.Policy)
                ReversePolicy    = Get-PolicyValue -PolicyObject $policy.NicTeaming.ReversePolicy -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.NicTeaming.ReversePolicy)
                NotifySwitch     = Get-PolicyValue -PolicyObject $policy.NicTeaming.NotifySwitches -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.NicTeaming.NotifySwitches)
                RollingOrder     = Get-PolicyValue -PolicyObject $policy.NicTeaming.RollingOrder -DefaultValue (Get-PolicyValue -PolicyObject $switchPolicy.NicTeaming.RollingOrder)
                Offload          = $true
                TSO              = if ($null -ne $hostEd.Capability.TsoSupported) { $hostEd.Capability.TsoSupported } else { $true }
                ZeroCopyXmit     = if ($null -ne $hostEd.Capability.ZeroCopyXmitSupported) { $hostEd.Capability.ZeroCopyXmitSupported } else { $true }
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

     
        $hostCtx = Get-HostContext -ESXhost $ESXhost
        $dcName = $hostCtx.Datacenter
        $clusterName = $hostCtx.Cluster
        $ed = $ESXhost.ExtensionData
        $hostNetwork = Get-VMHostNetwork -VMHost $ESXhost -ErrorAction SilentlyContinue
        foreach ($vmk in Get-VMHostNetworkAdapter -VMKernel -VMHost $ESXhost) {
            [PSCustomObject]@{
                Host                  = $ESXhost.Name
                Datacenter            = $dcName
                Cluster               = $clusterName
                VMKernelAdapter       = $vmk.Name
                IPAddress             = $vmk.IP
                SubnetMask            = $vmk.SubnetMask
                MAC                   = $vmk.Mac
                MTU                   = $vmk.Mtu
                vSwitch               = $vmk.VirtualSwitch
                PortGroup             = $vmk.PortGroupName
                DHCPEnabled           = $vmk.DhcpEnabled
                Gateway               = if ($ed.Config.Network.IpRouteConfig) { $ed.Config.Network.IpRouteConfig.DefaultGateway } else { $null }
                IPv6                  = $vmk.IPv6[0].address
                IPv6Gateway           = if ($ed.Config.Network.PSObject.Properties.Name -contains 'IpV6RouteConfig') { $ed.Config.Network.IpV6RouteConfig.DefaultGateway } else { $null }
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
        $hostsOnDs = @($ds | Get-VMHost)
        $clusterNames = @($hostsOnDs | Get-Cluster | Select-Object -ExpandProperty Name -Unique)
        $vmCount = ($ds | Get-VM).Count
        $capMiB = [math]::Round(($ds.CapacityGB * 1024), 0)
        $freeMiB = [math]::Round(($ds.FreeSpaceGB * 1024), 0)
        $provisionedMiB = if ($ds.ExtensionData.Summary.Capacity -and $ds.ExtensionData.Summary.FreeSpace) { [math]::Round((($ds.ExtensionData.Summary.Capacity - $ds.ExtensionData.Summary.FreeSpace) / 1MB), 0) } else { $null }
        [PSCustomObject]@{
            Name            = $ds.Name
            ConfigStatus    = $ds.ExtensionData.OverallStatus
            Address         = $ds.ExtensionData.Info.Vmfs.Extent | ForEach-Object { $_.DiskName } | Select-Object -First 1
            Accessible      = $ds.ExtensionData.Summary.Accessible
            Type            = $ds.Type
            NumVMsTotal     = $vmCount
            NumVMs          = $vmCount
            CapacityMiB     = $capMiB
            ProvisionedMiB  = $provisionedMiB
            InUseMiB        = if ($capMiB -and $freeMiB) { $capMiB - $freeMiB } else { $null }
            FreeMiB         = $freeMiB
            FreePct         = if ($capMiB) { [math]::Round(($freeMiB / $capMiB) * 100, 1) } else { $null }
            SIOCEnabled     = $ds.ExtensionData.IormConfiguration.Enabled
            SIOCThreshold   = if ($ds.ExtensionData.IormConfiguration) { $ds.ExtensionData.IormConfiguration.CongestionThreshold } else { $null }
            NumHosts        = $hostsOnDs.Count
            Hosts           = ($ESXhostsOnDS -join ', ')
            Cluster         = ($clusterNames -join ', ')
            ClusterCapacityMiB = $capMiB
            ClusterFreeSpaceMiB = $freeMiB
            BlockSize       = $ds.ExtensionData.Info.Vmfs.BlockSizeMb
            MaxBlocks       = $ds.ExtensionData.Info.Vmfs.MaxBlocks
            NumExtents      = @($ds.ExtensionData.Info.Vmfs.Extent).Count
            MajorVersion    = $ds.ExtensionData.Info.Vmfs.MajorVersion
            Version         = $ds.ExtensionData.Info.Vmfs.Version
            VMFSUpgradeable = $ds.ExtensionData.Info.Vmfs.VmfsUpgradable
            MHA             = $ds.ExtensionData.Summary.MultipleHostAccess
            Datacenter      = ($hostsOnDs | Get-Cluster | Get-Datacenter | Select-Object -ExpandProperty Name -Unique) -join ', '
            URL             = $dsExt.Info.Url
            ObjectID        = $dsExt.MoRef.Value
            UUID            = $dsExt.Info.MembershipUuid
            "com.vrlcm.snapshot" = ''
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
    $searchSpec.Query = @($generic)
    $searchSpec.Details = $flags
    $searchSpec.MatchPattern = @(
        '*.vmx',
        '*.vmdk',
        '*-flat.vmdk',
        '*-delta.vmdk',
        '*-ctk.vmdk'
    )

    $total = $Datastores.Count
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

        if (($ds.Type -in @('VMFS', 'vsan')) -and $ds.ExtensionData.Summary.MultipleHostAccess) {

            $browser = Get-View $ds.ExtensionData.Browser
            $rootPath = "[{0}]" -f $ds.Name

            # Build hash-set of files referenced by VMs/templates
            $vmFiles = @{}

            foreach ($obj in @(
                (Get-VM -Datastore $ds -ErrorAction SilentlyContinue)
                (Get-Template -Datastore $ds -ErrorAction SilentlyContinue)
            ) | Where-Object { $_ }) {

                try {
                    $view = Get-View $obj.Id -ErrorAction Stop
                    foreach ($file in ($view.LayoutEx.File | Where-Object { $_.Name })) {
                        $name = $file.Name.ToLower()

                        # normalize duplicate slashes
                        $name = $name -replace '(?<!:)/{2,}', '/'

                        $vmFiles[$name] = $true
                    }
                }
                catch {
                }
            }

            # Enumerate datastore contents
            $result = $browser.SearchDatastoreSubFolders($rootPath, $searchSpec)

            if ($result) {
                foreach ($folder in $result) {
                    foreach ($f in ($folder.File | Where-Object { $_ })) {

                        # FolderPath already usually ends with /
                        $full = "{0}{1}" -f $folder.FolderPath, $f.Path
                        $full = $full.ToLower() -replace '(?<!:)/{2,}', '/'

                        if (-not $vmFiles.ContainsKey($full)) {
                            [pscustomobject]@{
                                Datastore       = $ds.Name
                                FilePath        = $full
                                FileSizeGB      = if ($null -ne $f.FileSize) { [math]::Round($f.FileSize / 1GB, 2) } else { $null }
                                Modified        = $f.Modification
                                'VI SDK Server' = $about.FullName
                                'VI SDK UUID'   = $about.InstanceUuid
                            }
                        }
                    }
                }
            }
        }
    }

    Write-InlineProgress -Activity 'vZombie Processed' -Complete `
        -ProgressCharacter ([char]9632) `
        -ProgressFillCharacter ([char]9632) `
        -ProgressFill ([char]183) `
        -BarBracketStart $null `
        -BarBracketEnd $null
}

function Get-vLicense {
    param($about)
    foreach ($lic in $(Get-View LicenseManager -Server $vcc).Licenses) {
        $expiration = ($lic.Properties | Where-Object { $_.Key -match 'expirationDate|expiration' } | Select-Object -ExpandProperty Value -First 1)
        if (-not $expiration -and $lic.Name -ne 'Product Evaluation') { $expiration = 'Never' }

        $featureValues = foreach ($prop in @($lic.Properties | Where-Object { $_.Key -eq 'feature' })) {
            if ($prop.Value -is [System.Array]) {
                foreach ($entry in $prop.Value) {
                    if ($entry -and $entry.PSObject.Properties.Name -contains 'Value') { $entry.Value }
                    elseif ($entry -and $entry.PSObject.Properties.Name -contains 'Key') { $entry.Key }
                    else { [string]$entry }
                }
            }
            elseif ($prop.Value -and $prop.Value.PSObject.Properties.Name -contains 'Value') {
                $prop.Value.Value
            }
            else {
                $prop.Value
            }
        }

        [PSCustomObject]@{
            Name            = $lic.Name
            LicenseKey      = $lic.LicenseKey
            EditionKey      = $lic.EditionKey
            CostUnit        = $lic.costunit
            Total           = if ($lic.Total -eq 0 -and $lic.Name -ne 'Product Evaluation') { 'Unlimited' } else { $lic.Total }
            Used            = $lic.Used
            ExpirationDate  = $expiration
            Features        = (($featureValues | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ', ')
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
$platformToolsRequiredVersion = Get-PlatformToolsRequiredVersion -VMHosts $ESXhosts -VMs $vms
$vtools = Get-vTools     -vms $vms -about $about -GetVMContextFn ${function:Get-VMContext} -PlatformToolsRequiredVersion $platformToolsRequiredVersion
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
$redactionConfig = @{
    RedactVMNames = $doRedactVMNames
    RedactFqdnDomain = $doRedactFqdnDomain
    RedactIPAddresses = $doRedactIPAddresses
}

if ($redactionConfig.RedactVMNames -or $redactionConfig.RedactFqdnDomain -or $redactionConfig.RedactIPAddresses) {
    Write-Host "Applying requested redaction options before Excel export..." -ForegroundColor Yellow

    $vmNameSheets = @('vInfo', 'vCPU', 'vMemory', 'vDisk', 'vPartition', 'vSCSI', 'vNetwork', 'vFloppy', 'vCD', 'vSnapshot', 'vTools')
    $vmNameMap = if ($redactionConfig.RedactVMNames) { New-VMRedactionMap -vms $vms } else { @{} }

    $vinfo = Invoke-Redaction -Data $vinfo -SheetName 'vInfo' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vcpu = Invoke-Redaction -Data $vcpu -SheetName 'vCPU' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vmemory = Invoke-Redaction -Data $vmemory -SheetName 'vMemory' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vdisk = Invoke-Redaction -Data $vdisk -SheetName 'vDisk' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vpartition = Invoke-Redaction -Data $vpartition -SheetName 'vPartition' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vSCSI = Invoke-Redaction -Data $vSCSI -SheetName 'vSCSI' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vnetwork = Invoke-Redaction -Data $vnetwork -SheetName 'vNetwork' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vfloppy = Invoke-Redaction -Data $vfloppy -SheetName 'vFloppy' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vcd = Invoke-Redaction -Data $vcd -SheetName 'vCD' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vsnapshot = Invoke-Redaction -Data $vsnapshot -SheetName 'vSnapshot' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vtools = Invoke-Redaction -Data $vtools -SheetName 'vTools' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets

    $vrp = Invoke-Redaction -Data $vrp -SheetName 'vRP' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vcluster = Invoke-Redaction -Data $vcluster -SheetName 'vCluster' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vhost = Invoke-Redaction -Data $vhost -SheetName 'vHost' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vTLShost = Invoke-Redaction -Data $vTLShost -SheetName 'vTLShost' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vhba = Invoke-Redaction -Data $vhba -SheetName 'vHBA' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vnic = Invoke-Redaction -Data $vnic -SheetName 'vNIC' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vswitch = Invoke-Redaction -Data $vswitch -SheetName 'vSwitch' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vport = Invoke-Redaction -Data $vport -SheetName 'vPort' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $dvswitch = Invoke-Redaction -Data $dvswitch -SheetName 'dvSwitch' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $dvport = Invoke-Redaction -Data $dvport -SheetName 'dvPort' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vsc_vmk = Invoke-Redaction -Data $vsc_vmk -SheetName 'vSC_VMK' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vdatastore = Invoke-Redaction -Data $vdatastore -SheetName 'vDatastore' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vZombie = Invoke-Redaction -Data $vZombie -SheetName 'vZombieFiles' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vlicense = Invoke-Redaction -Data $vlicense -SheetName 'vLicense' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
    $vhealth = Invoke-Redaction -Data $vhealth -SheetName 'vHealth' -Config $redactionConfig -VMNameMap $vmNameMap -VMNameSheets $vmNameSheets
}

Write-Host "Creating worksheets layout..." -ForegroundColor Yellow

$OutputHeaderMap = Get-OutputHeaderMap
    $aliasMap = @{
        'vInfo' = @{ 'VM' = 'Name' }
        'vCPU' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Template' = 'Template'; 'SRM Placeholder' = 'SRM Placeholder'; 'CPUs' = 'vCPUs'; 'Cores p/s' = 'CoresPerSocket'; 'Max' = 'MaxCPU_MHz'; 'Overall' = 'CPU_Usage_MHz'; 'Level' = 'SharesLevel'; 'Reservation' = 'CPU_Reservation'; 'Entitlement' = 'EntitlementMHz'; 'DRS Entitlement' = 'DRSEntitlementMHz'; 'Limit' = 'CPULimit'; 'Hot Add' = 'CPUHotAdd'; 'Hot Remove' = 'CPUHotRemove'; 'Numa Hotadd Exposed' = 'NumaHotaddExposed'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID' }
        'vMemory' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Template' = 'Template'; 'SRM Placeholder' = 'SRM Placeholder'; 'Size MiB' = 'MemoryMB'; 'Memory Reservation Locked To Max' = 'MemoryReservationLockedToMax'; 'Consumed' = 'MemConsumedMB'; 'Consumed Overhead' = 'OverheadMB'; 'Private' = 'PrivateMB'; 'Shared' = 'SharedMB'; 'Swapped' = 'SwappedMB'; 'Ballooned' = 'BalloonedMB'; 'Active' = 'ActiveMB'; 'Entitlement' = 'EntitlementMB'; 'DRS Entitlement' = 'DRSEntitlementMB'; 'Level' = 'MemSharesLevel'; 'Shares' = 'MemShares'; 'Reservation' = 'MemReservation'; 'Limit' = 'MemLimitMB'; 'Max' = 'MaxMemoryUsageMB'; 'Hot Add' = 'MemHotAdd'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID' }
        'vDisk' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Template' = 'Template'; 'SRM Placeholder' = 'SRM Placeholder'; 'Disk Key' = 'DiskKey'; 'Disk UUID' = 'DiskUUID'; 'Capacity MiB' = 'CapacityMB'; 'Disk Path' = 'Path'; 'Disk Mode' = 'Mode'; 'Sharing mode' = 'SharingMode'; 'Eagerly Scrub' = 'EagerZero'; 'Split' = 'Split'; 'Write Through' = 'WriteThrough'; 'Level' = 'Level'; 'Reservation' = 'Reservation'; 'Limit' = 'Limit'; 'Controller' = 'Controller'; 'Label' = 'Label'; 'SCSI Unit #' = 'ControllerBus'; 'Unit #' = 'Unit'; 'Shared Bus' = 'SharedBus'; 'Raw LUN ID' = 'LunUuid'; 'Raw Comp. Mode' = 'RDMMode'; 'Internal Sort Column' = 'InternalSortColumn'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID' }
        'vPartition' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Template' = 'Template'; 'SRM Placeholder' = 'SRM Placeholder'; 'Disk Key' = 'DiskKey'; 'Capacity MiB' = 'CapacityMB'; 'Consumed MiB' = 'ConsumedMB'; 'Free MiB' = 'FreeMB'; 'Free %' = 'FreePct'; 'Internal Sort Column' = 'InternalSortColumn'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID' }
        'vNetwork' = @{ 'VM' = 'Name'; 'NIC label' = 'Adapter'; 'Mac Address' = 'MAC'; 'Starts Connected' = 'StartConnected'; 'Direct Path IO' = 'Direct Path IO'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID' }
        'vCD' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Template' = 'Template'; 'SRM Placeholder' = 'SRM Placeholder'; 'Device Node' = 'DeviceNode'; 'Starts Connected' = 'StartConnected'; 'Device Type' = 'Type'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID'; 'VMRef' = 'VMRef' }
        'vSnapshot' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Name' = 'SnapshotName'; 'Date / time' = 'Created'; 'Filename' = 'Filename'; 'Size MiB (vmsn)' = 'SizeMiBVMSN'; 'Size MiB (total)' = 'SizeMiBTotal'; 'State' = 'State'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID' }
        'vTools' = @{ 'VM' = 'Name'; 'Powerstate' = 'Powerstate'; 'Template' = 'Template'; 'SRM Placeholder' = 'SRM Placeholder'; 'VM Version' = 'VMVersion'; 'Tools Version' = 'ToolsVersion'; 'Required Version' = 'RequiredVersion'; 'Upgradeable' = 'ToolsStatus2'; 'Upgrade Policy' = 'ToolsUpgradePolicy'; 'Sync time' = 'SyncTime'; 'Heartbeat status' = 'AppHeartbeat'; 'Kernel Crash state' = 'KernelCrash'; 'Operation Ready' = 'OpsReady'; 'State change support' = 'StateChangeSupported'; 'Interactive Guest' = 'InteractiveReady'; 'VM UUID' = 'UUID'; 'VM ID' = 'VM_ID'; 'VMRef' = 'VM_ID' }
        'vRP' = @{ 'Resource Pool name' = 'Name'; 'Resource Pool path' = 'ResourcePoolPath'; '# VMs total' = 'NumVMsTotal'; '# VMs' = 'NumVMs'; '# vCPUs' = 'vCPUs'; 'CPU limit' = 'CPU_Limit'; 'CPU overheadLimit' = 'CPU_OverheadLimit'; 'CPU reservation' = 'CPU_Reservation'; 'CPU level' = 'CPU_SharesLevel'; 'CPU shares' = 'CPU_Shares'; 'CPU expandableReservation' = 'CPU_Expandable'; 'CPU maxUsage' = 'CPU_maxUsage'; 'CPU overallUsage' = 'CPU_overallUsage'; 'CPU reservationUsed' = 'CPU_reservationUsed'; 'CPU reservationUsedForVm' = 'CPU_reservationUsedForVm'; 'CPU unreservedForPool' = 'CPU_unreservedForPool'; 'CPU unreservedForVm' = 'CPU_unreservedForVm'; 'Mem Configured' = 'MemConfigured'; 'Mem limit' = 'Mem_Limit'; 'Mem overheadLimit' = 'Mem_OverheadLimit'; 'Mem reservation' = 'Mem_Reservation'; 'Mem level' = 'Mem_SharesLevel'; 'Mem shares' = 'Mem_Shares'; 'Mem expandableReservation' = 'Mem_Expandable'; 'Object ID' = 'ObjectID' }
        'vHost' = @{ 'Host' = 'Name'; 'Config status' = 'Status'; 'CPU usage %' = 'CPUUsagePct'; 'Memory usage %' = 'MemoryUsagePct'; 'Compliance Check State' = 'ComplianceCheckState'; 'in Maintenance Mode' = 'InMaintenanceMode'; 'in Quarantine Mode' = 'InQuarantineMode'; 'vSAN Fault Domain Name' = 'vSANFaultDomainName'; 'Speed' = 'SpeedMHz'; '# CPU' = 'CPUPackages'; '# Cores' = 'TotalCores'; '# Memory' = 'MemoryGB'; 'Memory Tiering Type' = 'MemoryTieringType'; 'Console' = 'Console'; '# NICs' = 'NumNICs'; '# HBAs' = 'NumHBAs'; '# VMs total' = 'NumVMsTotal'; '# VMs' = 'NumVMs'; '# vCPUs' = 'NumvCPUs'; 'vCPUs per Core' = 'vCPUsPerCore'; 'vRAM' = 'vRAM'; 'VM Used memory' = 'VMUsedMemory'; 'VM Memory Swapped' = 'VMMemorySwapped'; 'VM Memory Ballooned' = 'VMMemoryBallooned'; 'VMotion support' = 'VMotionSupport'; 'Storage VMotion support' = 'StorageVMotionSupport'; 'Current EVC' = 'CurrentEVC'; 'Max EVC' = 'MaxEVC'; 'Assigned License(s)' = 'AssignedLicenses'; 'ATS Heartbeat' = 'ATSHeartbeat'; 'ATS Locking' = 'ATSLocking'; 'Current CPU power man. policy' = 'CurrentCPUPowerManPolicy'; 'Supported CPU power man.' = 'SupportedCPUPowerMan'; 'Host Power Policy' = 'HostPowerPolicy'; 'ESX Version' = 'ESXiVersion'; 'DNS Servers' = 'DNSServers'; 'DHCP' = 'DHCP'; 'Domain' = 'Domain'; 'Domain List' = 'DomainList'; 'DNS Search Order' = 'DNSSearchOrder'; 'NTP Server(s)' = 'NTPServers'; 'NTPD running' = 'NTPDRunning'; 'Time Zone' = 'TimeZone'; 'Time Zone Name' = 'TimeZoneName'; 'GMT Offset' = 'GMTOffset'; 'Serial number' = 'Serial'; 'Service tag' = 'ServiceTag'; 'OEM specific string' = 'OEMSpecificString'; 'Certificate Issuer' = 'CertificateIssuer'; 'Certificate Start Date' = 'CertificateStartDate'; 'Certificate Expiry Date' = 'CertificateExpiryDate'; 'Certificate Status' = 'CertificateStatus'; 'Certificate Subject' = 'CertificateSubject'; 'Object ID' = 'ObjectID' }
        'vHBA' = @{ 'Device' = 'HBADevice'; 'WWN' = 'NodeWorldWideName'; 'Bus' = 'Bus'; 'Pci' = 'PCI' }
        'vNIC' = @{ 'Network Device' = 'PNICDevice'; 'Speed' = 'LinkSpeed'; 'Duplex' = 'Duplex'; 'Switch' = 'Switch'; 'Uplink port' = 'UplinkPort'; 'WakeOn' = 'WakeOnLAN' }
        'vSwitch' = @{ 'Switch' = 'vSwitch'; '# Ports' = 'NumPorts'; 'Free Ports' = 'NumPortsAvailable'; 'Promiscuous Mode' = 'AllowPromiscuous'; 'Mac Changes' = 'MacChanges'; 'Forged Transmits' = 'ForgedTransmits'; 'Traffic Shaping' = 'TrafficShaping'; 'Width' = 'Width'; 'Peak' = 'Peak'; 'Burst' = 'Burst'; 'Policy' = 'Policy'; 'Reverse Policy' = 'ReversePolicy'; 'Notify Switch' = 'NotifySwitch'; 'Rolling Order' = 'RollingOrder'; 'Offload' = 'Offload'; 'TSO' = 'TSO'; 'Zero Copy Xmit' = 'ZeroCopyXmit' }
        'vPort' = @{ 'Port Group' = 'PortGroup'; 'Switch' = 'vSwitch'; 'VLAN' = 'VLANId'; 'Promiscuous Mode' = 'AllowPromiscuous'; 'Mac Changes' = 'MacChanges'; 'Forged Transmits' = 'ForgedTransmits'; 'Traffic Shaping' = 'TrafficShaping'; 'Width' = 'Width'; 'Peak' = 'Peak'; 'Burst' = 'Burst'; 'Policy' = 'Policy'; 'Reverse Policy' = 'ReversePolicy'; 'Notify Switch' = 'NotifySwitch'; 'Rolling Order' = 'RollingOrder'; 'Offload' = 'Offload'; 'TSO' = 'TSO'; 'Zero Copy Xmit' = 'ZeroCopyXmit' }
        'dvSwitch' = @{ 'Switch' = 'Name'; '# Ports' = 'NumPorts' }
        'dvPort' = @{ 'Port' = 'Name'; 'Switch' = 'VDSwitch'; 'VLAN' = 'VLANId'; '# Ports' = 'NumPorts' }
        'vSC_VMK' = @{ 'Device' = 'VMKernelAdapter'; 'Mac Address' = 'MAC'; 'DHCP' = 'DHCPEnabled'; 'IP Address' = 'IPAddress'; 'IP 6 Address' = 'IPv6'; 'Subnet mask' = 'SubnetMask'; 'Gateway' = 'Gateway'; 'IP 6 Gateway' = 'IPv6Gateway'; 'Port Group' = 'PortGroup' }
        'vDatastore' = @{ 'Config status' = 'ConfigStatus'; 'Address' = 'Address'; 'Accessible' = 'Accessible'; '# VMs total' = 'NumVMsTotal'; '# VMs' = 'NumVMs'; 'Capacity MiB' = 'CapacityMiB'; 'Provisioned MiB' = 'ProvisionedMiB'; 'In Use MiB' = 'InUseMiB'; 'Free MiB' = 'FreeMiB'; 'Free %' = 'FreePct'; 'SIOC enabled' = 'SIOCEnabled'; 'SIOC Threshold' = 'SIOCThreshold'; '# Hosts' = 'NumHosts'; 'Cluster name' = 'Cluster'; 'Cluster capacity MiB' = 'ClusterCapacityMiB'; 'Cluster free space MiB' = 'ClusterFreeSpaceMiB'; 'Block size' = 'BlockSize'; 'Max Blocks' = 'MaxBlocks'; '# Extents' = 'NumExtents'; 'Major Version' = 'MajorVersion'; 'Version' = 'Version'; 'VMFS Upgradeable' = 'VMFSUpgradeable'; 'MHA' = 'MHA'; 'Object ID' = 'UUID' }
        'vLicense' = @{ 'Key' = 'LicenseKey'; 'Cost Unit' = 'CostUnit'; 'Expiration Date' = 'ExpirationDate'; 'Features' = 'Features' }
        'vHealth' = @{ 'Name' = 'AlarmName'; 'Message' = 'AlarmData'; 'Message type' = 'AlarmStatus' }
    }

    $vinfo = Convert-ToOutputSchema -Data $vinfo -SheetName 'vInfo' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vcpu = Convert-ToOutputSchema -Data $vcpu -SheetName 'vCPU' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vmemory = Convert-ToOutputSchema -Data $vmemory -SheetName 'vMemory' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vdisk = Convert-ToOutputSchema -Data $vdisk -SheetName 'vDisk' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vpartition = Convert-ToOutputSchema -Data $vpartition -SheetName 'vPartition' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vnetwork = Convert-ToOutputSchema -Data $vnetwork -SheetName 'vNetwork' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vcd = Convert-ToOutputSchema -Data $vcd -SheetName 'vCD' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vsnapshot = Convert-ToOutputSchema -Data $vsnapshot -SheetName 'vSnapshot' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vtools = Convert-ToOutputSchema -Data $vtools -SheetName 'vTools' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vrp = Convert-ToOutputSchema -Data $vrp -SheetName 'vRP' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vcluster = Convert-ToOutputSchema -Data $vcluster -SheetName 'vCluster' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vhost = Convert-ToOutputSchema -Data $vhost -SheetName 'vHost' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vhba = Convert-ToOutputSchema -Data $vhba -SheetName 'vHBA' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vnic = Convert-ToOutputSchema -Data $vnic -SheetName 'vNIC' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vswitch = Convert-ToOutputSchema -Data $vswitch -SheetName 'vSwitch' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vport = Convert-ToOutputSchema -Data $vport -SheetName 'vPort' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $dvswitch = Convert-ToOutputSchema -Data $dvswitch -SheetName 'dvSwitch' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $dvport = Convert-ToOutputSchema -Data $dvport -SheetName 'dvPort' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vsc_vmk = Convert-ToOutputSchema -Data $vsc_vmk -SheetName 'vSC_VMK' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vdatastore = Convert-ToOutputSchema -Data $vdatastore -SheetName 'vDatastore' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vlicense = Convert-ToOutputSchema -Data $vlicense -SheetName 'vLicense' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
    $vhealth = Convert-ToOutputSchema -Data $vhealth -SheetName 'vHealth' -OutputHeaderMap $OutputHeaderMap -AliasMap $aliasMap
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

$DMexecend = Get-Date
$duration = New-TimeSpan -Start $DMexecstart -End $DMexecend
Write-Host "DMTools execution time: $($duration.ToString("hh\:mm\:ss"))" -ForegroundColor Green





