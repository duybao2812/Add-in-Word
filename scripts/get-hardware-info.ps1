# Script PowerShell thu thap day du thong so he thong va tra ve chuoi JSON
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    $os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
    $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $cpus = Get-CimInstance Win32_Processor -ErrorAction SilentlyContinue
    $cpu = $cpus | Select-Object -First 1
    $ramModules = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
    $gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue
    $disks = Get-CimInstance Win32_DiskDrive -ErrorAction SilentlyContinue
    $volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction SilentlyContinue
    $bios = Get-CimInstance Win32_BIOS -ErrorAction SilentlyContinue
    $bb = Get-CimInstance Win32_BaseBoard -ErrorAction SilentlyContinue
    $csp = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
    $netAdapters = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled = TRUE" -ErrorAction SilentlyContinue

    $ramList = @()
    if ($ramModules) {
        foreach ($m in $ramModules) {
            $ramList += [PSCustomObject]@{
                Manufacturer            = if ($m.Manufacturer) { $m.Manufacturer.Trim() } else { "Unknown" }
                CapacityGB              = [math]::Round($m.Capacity / 1GB, 2)
                SpeedMHz                = $m.Speed
                ConfiguredClockSpeedMHz = $m.ConfiguredClockSpeed
                PartNumber              = if ($m.PartNumber) { ($m.PartNumber -replace '\s+', ' ').Trim() } else { "Unknown" }
                BankLabel               = $m.BankLabel
                DeviceLocator           = $m.DeviceLocator
            }
        }
    }

    $gpuList = @()
    if ($gpus) {
        foreach ($g in $gpus) {
            $vramGB = if ($g.AdapterRAM -and $g.AdapterRAM -gt 0) { [math]::Round($g.AdapterRAM / 1GB, 2) } else { 0 }
            $res = if ($g.CurrentHorizontalResolution) { "{0}x{1} @ {2}Hz" -f $g.CurrentHorizontalResolution, $g.CurrentVerticalResolution, $g.CurrentRefreshRate } else { "N/A" }
            $gpuList += [PSCustomObject]@{
                Name              = $g.Name
                DriverVersion     = $g.DriverVersion
                VideoProcessor    = $g.VideoProcessor
                AdapterRAM_GB     = $vramGB
                CurrentResolution = $res
            }
        }
    }

    $diskList = @()
    if ($disks) {
        foreach ($d in $disks) {
            $diskList += [PSCustomObject]@{
                Model         = if ($d.Model) { $d.Model.Trim() } else { "Unknown" }
                InterfaceType = $d.InterfaceType
                SizeGB        = [math]::Round($d.Size / 1GB, 2)
                Partitions    = $d.Partitions
                SerialNumber  = if ($d.SerialNumber) { ($d.SerialNumber -replace '\s+', ' ').Trim() } else { "Unknown" }
            }
        }
    }

    $volumeList = @()
    if ($volumes) {
        foreach ($v in $volumes) {
            $totalGB = [math]::Round($v.Size / 1GB, 2)
            $freeGB = [math]::Round($v.FreeSpace / 1GB, 2)
            $pct = if ($v.Size -gt 0) { [math]::Round(($v.FreeSpace / $v.Size) * 100, 1) } else { 0 }
            $volumeList += [PSCustomObject]@{
                DriveLetter = $v.DeviceID
                VolumeName  = $v.VolumeName
                FileSystem  = $v.FileSystem
                TotalSizeGB = $totalGB
                FreeSpaceGB = $freeGB
                FreePercent = $pct
            }
        }
    }

    $netList = @()
    if ($netAdapters) {
        foreach ($n in $netAdapters) {
            $netList += [PSCustomObject]@{
                Description = $n.Description
                MACAddress  = $n.MACAddress
                IPAddress   = ($n.IPAddress -join ', ')
            }
        }
    }

    # Tao ID duy nhat cho may tinh: uu tien System UUID > Motherboard Serial > BIOS Serial > ComputerName
    $rawUUID = if ($csp -and $csp.UUID) { $csp.UUID.Trim() } else { "" }
    $biosSerial = if ($bios -and $bios.SerialNumber) { $bios.SerialNumber.Trim() } else { "" }
    $mbSerial = if ($bb -and $bb.SerialNumber) { $bb.SerialNumber.Trim() } else { "" }
    $primaryMac = ($netList | Where-Object { $_.MACAddress } | Select-Object -First 1).MACAddress

    $uniqueMachineId = if ($rawUUID -and $rawUUID -ne "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF") {
        $rawUUID
    } elseif ($biosSerial) {
        "{0}_{1}" -f $env:COMPUTERNAME, $biosSerial
    } else {
        "{0}_{1}" -f $env:COMPUTERNAME, ($primaryMac -replace ':', '')
    }

    $profile = [PSCustomObject]@{
        MachineId         = $uniqueMachineId
        ComputerName      = $env:COMPUTERNAME
        UserName          = $env:USERNAME
        SystemUUID        = $rawUUID
        Timestamp         = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        OperatingSystem   = [PSCustomObject]@{
            Caption       = $os.Caption
            Version       = $os.Version
            BuildNumber   = $os.BuildNumber
            Architecture  = $os.OSArchitecture
        }
        Motherboard       = [PSCustomObject]@{
            Manufacturer  = $bb.Manufacturer
            Product       = $bb.Product
            SerialNumber  = $mbSerial
        }
        BIOS              = [PSCustomObject]@{
            Manufacturer  = $bios.Manufacturer
            Version       = $bios.SMBIOSBIOSVersion
            SerialNumber  = $biosSerial
        }
        Processor         = [PSCustomObject]@{
            Name                      = $cpu.Name
            Manufacturer              = $cpu.Manufacturer
            Architecture              = ("{0}-bit" -f $cpu.AddressWidth)
            NumberOfCores             = $cpu.NumberOfCores
            NumberOfLogicalProcessors = $cpu.NumberOfLogicalProcessors
            MaxClockSpeedMHz          = $cpu.MaxClockSpeed
            SocketDesignation         = $cpu.SocketDesignation
        }
        Memory            = [PSCustomObject]@{
            TotalPhysicalRAM_GB       = [math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
            SlotCount                 = $ramList.Count
            Modules                   = $ramList
        }
        Graphics          = $gpuList
        PhysicalDisks     = $diskList
        LogicalVolumes    = $volumeList
        NetworkAdapters   = $netList
    }

    $profile | ConvertTo-Json -Depth 6
} catch {
    Write-Error $_
    exit 1
}
