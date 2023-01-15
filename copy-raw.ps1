#
# USE AT YOUR OWN RISK! NO LIABILITY!
#
# Synopsis: 
# Powershell Script to copy the raw version of a foto from a connected Nikon camera for all
# jpeg version fotos in the current directory that have a minimum rating of 3 stars
# raw files are stored in a "raw" sub folder that is created if needed
#
# Helpful, if you use Lightroom Classic that does not support your newer camera raw files and
# if you first do the selection and rating with the jpg files 
# and then get the raw version for the high rated fotos - this speeds up your workflow and saves 
# storage
# 
# Adjustment:
# Adjust and localize parameter $source_path. (currently Nikon D780 and german)
# Adjust parameter $rating. (currently 3,4 and 5 stars)
#
# Usage:
# Open Powershell Terminal in VS Code
# Change directoy (cd) into the directory with the rated fotos jpeg fotos
# Call this script (like ..\..\..\copy-raw.ps1 if it is stored some parent folders up)
# PS C:\Users\Me\Pictures\2020-2029\2023\2023-01-07 Winter> ..\..\..\copy-raw.ps1
#
# Sources:
# https://devblogs.microsoft.com/scripting/use-powershell-to-find-metadata-from-photograph-files/
# https://stackoverflow.com/questions/55628092/how-to-reliably-copy-items-with-powershell-to-an-mtp-device
# https://gist.github.com/woehrl01/5f50cb311f3ec711f6c776b2cb09c34e
# https://gallery.technet.microsoft.com/scriptcenter/Get-FileMetaData-3a7ddea7
# 
# 
param($rating = '[3-5]',  # as reg_ex!
    [byte]$source_folder_index_start = 100,
    [string]$source_path = 'D780/Wechselmedien 10001/DCIM',
    [string]$dest_path = "$(Get-Location)\raw",
    [string]$raw ='.NEF'
)

function Get-FileMetaData 
{ 
    <# 
    .SYNOPSIS 
        Get-FileMetaData returns metadata information about a single file. 
    .DESCRIPTION 
        This function will return all metadata information about a specific file. It can be used to access the information stored in the filesystem. 
    .EXAMPLE 
        Get-FileMetaData -File "c:\temp\image.jpg" 
        Get information about an image file. 
    .EXAMPLE 
        Get-FileMetaData -File "c:\temp\image.jpg" | Select Dimensions 
        Show the dimensions of the image. 
    .EXAMPLE 
        Get-ChildItem -Path .\ -Filter *.exe | foreach {Get-FileMetaData -File $_.Name | Select Name,"File version"} 
        Show the file version of all binary files in the current folder. 
    #> 
 
    param([Parameter(Mandatory=$True)][string]$File) 
 
    if(!(Test-Path -Path $File)) 
    { 
        throw "File does not exist: $File" 
        Exit 1 
    } 
 
    $tmp = Get-ChildItem $File 
    $pathname = $tmp.DirectoryName 
    $filename = $tmp.Name 
 
    $hash = @{}
    try{
        $shellobj = New-Object -ComObject Shell.Application 
        $folderobj = $shellobj.namespace($pathname) 
        $fileobj = $folderobj.parsename($filename) 
        
        for($i=0; $i -le 294; $i++) 
        { 
            $name = $folderobj.getDetailsOf($null, $i);
            if($name){
                $value = $folderobj.getDetailsOf($fileobj, $i);
                if($value){
                    $hash[$($name)] = $($value)
                }
            }
        } 
    }finally{
        if($shellobj){
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$shellobj) | out-null
        }
    }

    return New-Object PSObject -Property $hash
} 

function Copy-Item {
    param (
        [Parameter(Mandatory=$true)] $item, 
        [Parameter(Mandatory=$true)] [String] $path)

    $dir_exists = $false
    if (-not $dir_exists) {
        if ( -not (Test-Path -PathType Container $path )) {
            try {
                New-Item -Path $path -ItemType Directory -ErrorAction Stop | Out-Null #-Force
            }
            catch {
                Write-Error -Message "Unable to create directory '$path'. Error was: $_" -ErrorAction Stop
            }
            "Successfully created directory '$path'."
        }
        $dir_exists = $true
    }
    $Shell = New-Object -ComObject Shell.Application
    $DestFolder = $Shell.NameSpace($path).self.GetFolder()
    $DestFolder.CopyHere($item)

    Do {
        Start-Sleep -Milliseconds 100
        $CopiedFile = $DestFolder.Items() | Where-Object{$_.Name -eq $Item.Name}
    }While( ($null -eq $CopiedFile) )#skip sleeping if it's already copied

    Write-Host "Copied $($item.Name)"

}

$Shell = New-Object -ComObject Shell.Application
$source_folder = $Shell.NameSpace(17).self
$source_path -split '/' | ForEach-Object {
    $folder = $_
    $source_folder = $source_folder.GetFolder.items() | Where-Object { $_.name -eq $folder }
}

$source_folder.GetFolder.items() | ForEach-Object {
    $source_subfolder = $_
    Get-ChildItem -File -Filter *.jpg | ForEach-Object {
        if ( ( Get-FileMetaData -File $_ | Select-Object Bewertung) -match $rating ) {
            $file = "$(($_.BaseName -Split "-")[1])$raw"
            #foreach($Item in $source_folder.GetFolder.items() | Where-Object{$_.Name -match $file }) {
            #    Copy-Item $item $dest_path
            #}
            $source_subfolder.GetFolder.items() | Where-Object{$_.Name -eq $file} | ForEach-Object {
            Copy-Item $_ $dest_path
            }
        }
    }
}





