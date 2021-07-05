<#
 Ajouter 
 - get-nasystemversion
 - get-nasysteminfo
 - Get-NaShelf
 - Get-NaLun | Get-NaLunMap (liste des LUNs et du mapping)


#>
# Import des modules Netapp
param([Parameter(Mandatory=$True)]
    [string]$Controller# = "iaast02-stor02-cbv-val"
)

#Import-Module DataONTAP

Function Connect-Netapp {
    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$True)]
         [ValidateNotNullorEmpty()]
         [string]$NaController
    )
    Process {
        if ((Get-NaCredential).name -contains $NaController){
            try {
                $Connected = Connect-NaController $NaController -ErrorAction Stop | Out-Null
                write-host "Connecté à $NaController"
            }
            catch [ Exception ]{
                if ($_.CategoryInfo.Reason -eq "InvalidLogin"){
                    write-host "Le login ou password est incorrecte."
                    Remove-NaCredential -Controller $NaController
                    Connect-Netapp $NaController
                }
                write-host $_#.CategoryInfo.Reason

            }
        }else{
            write-host "Veuillez entrer les informations de connection :"
            $Credential = Get-Credential
            Add-NaCredential -Name $NaController -Credential $Credential  | Out-Null

            Connect-Netapp $NaController
        }
        return $Connected
    }
}

$NaControllers = @()
$NaControllers += $Controller

Connect-Netapp $Controller

# Répertoires
$CurrentPath = Split-Path $script:MyInvocation.MyCommand.Path
$DataPath = "$CurrentPath\..\data\$Controller"
if( -Not (Test-Path -Path $DataPath) ) { New-Item -ItemType directory -Path $DataPath | Out-Null}

write-host -ForegroundColor Blue "Get-NaCluster"
$Partner = (Get-NaCluster).Partner

#Connect-Netapp $Partner
if ($Partner){
    $NaControllers += $Partner
}

$AuditDisks = @()
$AuditAggrs = @()
$AuditVols = @()
$AuditLuns = @()
$AuditLunMaps = @()
$AuditIgroups = @()
$AuditShelves = @()
$AuditNfsExport = @()


foreach ($NaController in $NaControllers){
    Connect-Netapp $NaController
    write-host -ForegroundColor Blue "  Get-NaDisk"
    $Disks = Get-NaDisk #-Controller $NaControllers
# A développer
    foreach ($Disk in $Disks){
        $objDisk = "" | Select-Object  Name, Node, Aggregate, Status,RaidGroup, VendorId, DiskType, DiskModel, FirmwareRevision, PhysicalSpaceGB, Rpm, SerialNumber, `
        Shelf, Bay
        $objDisk.Name = $Disk.Name
        $objDisk.Node = $Disk.Node
        $objDisk.Aggregate = $Disk.Aggregate
        $objDisk.Status = $Disk.Status
        $objDisk.RaidGroup = $Disk.RaidGroup
        $objDisk.VendorId = $Disk.VendorId
        $objDisk.DiskType = $Disk.DiskType
        $objDisk.DiskModel = $Disk.DiskModel
        $objDisk.Rpm = $Disk.Rpm
        $objDisk.SerialNumber = $Disk.SerialNumber
        $objDisk.FirmwareRevision = $Disk.FirmwareRevision
        $objDisk.PhysicalSpaceGB = $Disk.PhysicalSpace /1024 /1024 /1024

        $objDisk.Shelf = $Disk.Shelf
        $objDisk.Bay = $Disk.Bay

        $AuditDisks += $objDisk
    }

    write-host -ForegroundColor Blue "  Get-NaAggr"
    $Aggrs = Get-NaAggr #-Controller $NaController
    foreach ($Aggr in $Aggrs){
        $objAggr = "" | Select-Object  Name, HomeName, HaPolicy, IsMirrored , IsSnaplock, MirrorStatus,  `
        DiskCount, RaidSize, RaidStatus,`
        SizeAvailable, SizePercentageUsed,  SizeTotal, SizeUsed, `
        MountState, State, PlexCount, `
        VolumeCount
        
        $objAggr.Name = $Aggr.Name
        $objAggr.HomeName = $Aggr.HomeName
        $objAggr.HaPolicy = $Aggr.HaPolicy
        $objAggr.IsMirrored = $Aggr.IsMirrored
        $objAggr.IsSnaplock = $Aggr.IsSnaplock
        $objAggr.MirrorStatus = $Aggr.MirrorStatus

        $objAggr.DiskCount = $Aggr.DiskCount
        $objAggr.RaidSize = $Aggr.RaidSize
        $objAggr.RaidStatus = $Aggr.RaidStatus

        $objAggr.SizeAvailable = $Aggr.SizeAvailable /1024 /1024 /1024
        $objAggr.SizePercentageUsed = $Aggr.SizePercentageUsed
        $objAggr.SizeTotal = $Aggr.SizeTotal/1024 /1024 /1024
        $objAggr.SizeUsed = $Aggr.SizeUsed /1024 /1024 /1024

        $objAggr.MountState = $Aggr.MountState
        $objAggr.State = $Aggr.State
        $objAggr.PlexCount = $Aggr.PlexCount

        $objAggr.VolumeCount = $Aggr.VolumeCount


        $AuditAggrs += $objAggr
    }

    write-host -ForegroundColor Blue "  Get-NaVol"
    $Vols = Get-NaVol #-Controller $NaController
    foreach ($Vol in $Vols){
        $objVol = "" | Select-Object  Name, Node, Aggregate, Dedupe, `
                State, TotalSize, Used, Available

        $objVol.Name = $Vol.Name
        $objVol.Node = $NaController
        $objVol.Aggregate = $Vol.Aggregate
        $objVol.Dedupe = $Vol.Dedupe
        $objVol.State = $Vol.State
        $objVol.TotalSize = $Vol.TotalSize /1024 /1024 /1024
        $objVol.Used = $Vol.Used
        $objVol.Available = $Vol.Available /1024 /1024 /1024


        $AuditVols += $objVol
    }

    write-host -ForegroundColor Blue "  Get-NaLun et get-MapLun"
    $Luns = Get-NaLun #-Controller $NaController
    foreach ($Lun in $Luns){
        $objLun = "" | Select-Object  Path, Protocol, Thin, Mapped, Online, SizeGB, SizeUsedGB
        $objLun.Path = $Lun.Path
        $objLun.Protocol = $Lun.Protocol
        $objLun.Thin = $Lun.Thin
        $objLun.Mapped = $Lun.Mapped
        $objLun.Online = $Lun.Online
        $objLun.SizeGB = $Lun.Size / 1024 /1024 /1024
        $objLun.SizeUsedGB = $Lun.SizeUsed /1024 /1024/1024

        $AuditLuns += $objLun

        $LunMaps = Get-NaLunMap  -Path $Lun.Path #-Controller $NaController
        $objLunMap = "" | Select-Object  Name, Type, Protocol, Partner, Initiators
        $objLunMap.Name = $LunMaps.Name
        $objLunMap.Type = $LunMaps.Type
        $objLunMap.Protocol = $LunMaps.Protocol
        $objLunMap.Partner = $LunMaps.Partner
        $objLunMap.Initiators = $LunMaps.Initiators -join (", ")

        $AuditLunMaps += $objLunMap
    }

    write-host -ForegroundColor Blue "  Get-NaIgroup"
    $Igroups = Get-NaIgroup
    foreach ($Igroup in $Igroups){
        $objIgroup = "" | Select-Object  Name,Node, Type, Protocol, ALUA, Initiators
        $objIgroup.Name =  $Igroup.Name
        $objIgroup.Node =  $NaController
        $objIgroup.Type =  $Igroup.Type
        $objIgroup.Protocol =  $Igroup.Protocol
        $objIgroup.ALUA =  $Igroup.ALUA
        $objIgroup.Initiators =  $Igroup.Initiators -join (", ")

        $AuditIgroups += $objIgroup | Get-Unique

    }

    write-host -ForegroundColor Blue "  Get-NaShelf"
    $Shelves = Get-NaShelf
    foreach ($Shelve in $Shelves){
        $objShelf = "" | Select-Object  Name, Node, Status, Channel, ShelfId, ShelfType, FirmwareRevA, FirmwareRevB, Module, ModuleState, ShelfUid
        $objShelf.Name =  $Shelve.Name
        $objShelf.Node =  $NaController
        $objShelf.Status =  $Shelve.Status
        $objShelf.Channel =  $Shelve.Channel
        $objShelf.ShelfId =  $Shelve.ShelfId
        $objShelf.ShelfType =  $Shelve.ShelfType
        $objShelf.FirmwareRevA =  $Shelve.FirmwareRevA
        $objShelf.FirmwareRevB =  $Shelve.FirmwareRevB
        $objShelf.Module =  $Shelve.Module
        $objShelf.ModuleState =  $Shelve.ModuleState
        $objShelf.ShelfUid =  $Shelve.ShelfUid

        $AuditShelves += $objShelf |  get-unique

    }

    write-host -ForegroundColor Blue "  Get-NaNfsExport"
    $NfsExports = Get-NaNfsExport
    foreach ($NfsExport in $NfsExports){
        $objNfsEsxport = "" | Select-Object  PathName, ReadOnly, ReadWrite
        $objNfsEsxport.PathName =  $NfsExport.PathName
        $objNfsEsxport.ReadOnly =  $NfsExport.SecurityRules.ReadOnly
        $objNfsEsxport.ReadWrite = $NfsExport.SecurityRules.ReadWrite
        
        $AuditNfsExport += $objNfsEsxport | Get-Unique

    }
}
   
$Time = "{0:yyyyMMdd_HHmm}" -f (get-date)
$XlsFile = "$DataPath\Export_Netapp_$Time.xlsx"

$AuditAggrs     |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname Aggrs
$AuditDisks     |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname Disks
$AuditVols      |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname Vols
$AuditLuns      |  Sort-Object Path | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname Luns
$AuditLunMaps   |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname LunMaps
$AuditIgroups   |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname Igroups
$AuditShelves   |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname Shelves
$AuditNfsExport |  Sort-Object Name | Export-Excel -AutoFilter -AutoSize -FreezeTopRowFirstColumn -Path $XlsFile -WorkSheetname NfsExport

invoke-item -Path $XlsFile
