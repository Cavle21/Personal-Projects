<#

.PARAMETER $mode
PrepForEmail, FixAfterEmail, QuickChange

#>

param (
    [Parameter(Mandatory=$true,ParameterSetName='Mode')]
    [ValidateSet("PrepForEmail", "FixAfterEmail", "QuickChange")]
    [ValidateNotNullOrEmpty()]
        [string]$mode = "QuickChange"
)

Function New-FolderToPlaceFiles {
    param()

    [string]$folderPathandName = "$PSScriptRoot\FilesToCopy"

    if((Test-Path $folderPathandName) -ne $true){

        New-Item -ItemType directory -Path $folderPathandName | Out-Null
        Write-Host "Folder Created" -BackgroundColor Blue
    }else {
        Write-host "Folder already exists"
    }

    $folderContents = Get-ChildItem $folderPathandName

    if ([string]::IsNullOrEmpty($folderContents) -eq $false){
        Write-host "Folder has contents. Deleting files" -BackgroundColor Blue
        ForEach ($file in $folderContents){
            [io.path]::Combine($folderPathandName,$file) | Remove-Item
        }
    }else {
        Write-host "Folder empty."
    }
    return $folderPathandName
}
Function Get-FilesFromBrowse {
    param(
        [String]$title = "Choose The Files You Want To Process",
        [String]$initialDirectory = "$PSScriptRoot",
        [string]$startingFile = "",
        [boolean]$multiSelectFlag = $true
    )

    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = Split-Path $initialDirectory -Parent 
    $OpenFileDialog.Multiselect = $multiSelectFlag
    #$OpenFileDialog.CheckFileExists
    #$OpenFileDialog.CheckPathExists
    $OpenFileDialog.Title = $title
    if ([string]::IsNullOrWhiteSpace($startingFile -ne $true)){
        $OpenFileDialog.FileName = Split-path $startingFile -leaf
    }
    $OpenFileDialog.ShowDialog() | Out-Null

    [System.Collections.ArrayList]$fileNames = $OpenFileDialog.SafeFileNames
    [System.Collections.ArrayList]$fileLocations = $OpenFileDialog.FileNames
    [System.Collections.Hashtable]$fileHash = @{}

    For ($i = 0; $i -lt $fileNames.Count; $i++){
        $fileHash.Add($fileNames[$i], $fileLocations[$i])
    }
    #cannot pass an empty value and cannot pass a folder.
    if ([string]::IsNullOrEmpty($fileHash) -ne $true -and $fileHash.GetType() -ne [System.IO.FileSystemInfo]){
        return $fileHash
    }else{
        Write-host "Please choose at least one file to prepare for email" -BackgroundColor Red
        Get-Files
    }
}

Function Get-FilesToConvertBack {
    param(

        [String]$path
    )

    Write-host "Grabbing all files from folder"



    [System.Collections.ArrayList]$myFiles = @(Get-ChildItem -Path $path -Exclude "*.csv" | Select-Object BaseName, FullName)
    [System.Collections.Hashtable]$fileSpecifics = @{}

    ForEach ($file in $myFiles){
        $fileSpecifics.add($file.BaseName, $file.FullName)
    }

    Return $fileSpecifics
}

Function Copy-Files {
    param (
        [System.Collections.HashTable]$filesToCopy,

        [string]$filePath
    )

    ForEach ($key in $filesToCopy.keys){
        try{
            Copy-Item -Path $filesToCopy[$key] -Destination $filePath -ErrorAction stop
            Write-host "$key copied to $filePath"
        }catch{
            Write-host "$key failed to copy to $filePath"
            Write-host $error[0].exception.Message
            return $false
        }
    }
    return $true   
}

Function Get-AllFileExtensions {
    param(
        [String]$path
    )

    [System.Collections.Hashtable]$fileExtensions = @{}

    Write-host "Getting all file extensions"
    
    Get-ChildItem -Path $path -Exclude "*.csv" | ForEach-Object { 
        [string]$name = $_.BaseName
        [string]$extension = $_.Extension
        Write-host ":::::Name: $name - Extension: $extension"

        $fileExtensions.add($name, $extension)

    }

    Write-host "Complete"

    return $fileExtensions
}

Function Save-OriginalFileExtensionsAndNamesToCSV {

    param (
    
        [string]$Path
    
    )

    [System.Collections.Hashtable]$fileExtensions = Get-AllFileExtensions -path $path

    $csvStoredPath = [io.path]::Combine($path,"FileExtensions.csv")

    Write-host "Saving original file extensions to $csvStoredPath"

    Try {

        $fileExtensions.GetEnumerator() | Select-Object Key,Value | Export-Csv -path $csvStoredPath -NoTypeInformation
    }catch {
        Write-host $error[0].exception.message
        Write-host "Export failed!"
        Return $false
    }

    return $true
}

Function Set-FileExtensionToTxt {

    param(
        [System.Collections.HashTable]$myFiles,
        [string]$folderPath
    )

    ForEach ($file in $myFiles.keys){
        $newFileName = [io.path]::ChangeExtension("$file", "txt")
        $pathTofile = [io.path]::Combine($folderPath, $file)

        try {
            rename-item -Path "$pathToFile" -newname $newFileName 
        }catch{
            Write-host "File rename failed"
            Write-host $error[0].exception.message
            return $false
        }
    }

    return $true
}

Function Get-FileNamePairedOriginalFileExtensionsFromCSV{
    param()

    [System.Collections.Hashtable]$csvFileInformation = Get-FilesFromBrowse -title "Get CSV File That Was Generated" -multiSelectFlag $false

    [string]$csvStoredPath = $csvFileInformation.Values

    [System.Collections.Hashtable]$FileNamePairedOriginalFileExtensions = @{}

    Import-Csv -Path $csvStoredPath | ForEach-Object { $FileNamePairedOriginalFileExtensions.add($_.key, $_.value)}

    return $FileNamePairedOriginalFileExtensions

}

Function Set-FileExtensionToOriginal {
    
    param(
    
        [System.Collections.hashtable]$FilesPairedFileLocations,
        [System.Collections.Hashtable]$FileNamePairedOriginalFileExtensions,
        [string]$Path
    )

    [System.collections.hashtable]$FileNamePairedOriginalFileExtensions = @{}
    [System.Collections.ArrayList]$newFileNames = @()

    Write-host "Changing extensions back to original"

    Try {

        ForEach ($item in $FilesPairedFileLocations.Keys){
            $file = $item.BaseName
            ForEach ($key in $FileNamePairedOriginalFileExtensions.keys){
                if($key -ne $file){
                    continue                    
                }else{
                    $newFileName = [io.path]::ChangeExtension($file, $FileNamePairedOriginalFileExtensions[$key])
                    $pathTofile = [io.path]::Combine($Path, $File + $item.extension)
                    rename-item -Path $pathToFile -newname $newFileName 
                    Write-host "::::$file changed to $newFileName"
                    $newFileNames.add($newFileName)
                    break
                }
            }
        }
    }catch{
        Write-host $error[0].exception.message
    }

    return $newFileNames
}

Function Clear-Folder {
    param(
        [string]$path
    )

    Write-host "Cleaning up folder and CSV"

    Get-ChildItem -path $path -Exclude "*.csv" | Remove-Item -Recurse

    $csvStoredPath = [io.path]::Combine($path, "fileExtensions.csv")

    Clear-Content -path $csvStoredPath -Force

}

Function Get-ExtensionFromUser {
    param()

    [string]$extension = Read-Host -Prompt "What would you like to change these files to? ex. .ps1"

    if ([string]::IsNullOrWhiteSpace($extension) -or $extension.GetType() -ne [string]){
        
        Get-ExtensionFromUser
    } elseif ($extension -eq "Exit"){
        exit
    } elseif ($extension.StartsWith(".") -ne $true){
        Write-host "Please put the dot before the extension name. Ex. .ps1"
        Get-ExtensionFromUser
    }elseif ($extension.Length -gt 4) {
        Write-Host "Please input an extension or type EXIT to get out of the script"
        Get-ExtensionFromUser
    }

    return $extension
}

Function Set-ExtensionFromUser {
    param (
        [string]$extension,
        [System.Collections.Hashtable]$fileNamesPairedFileLocation
    )

    try{
        ForEach ($FileName in $fileNamesPairedFileLocation.Keys){
                $newFileName = [io.path]::ChangeExtension($fileName, $extension)
                $pathTofile = $fileNamesPairedFileLocation[$FileName]
                rename-item -Path $pathToFile -newname $newFileName 
                Write-host "::::$fileBaseName changed to $newFileName"
        }
    }catch{
        Write-host $error[0].exception.message
    return $false
}

return $true
}

Function Main {

    [string]$filePath
    [System.Collections.HashTable]$filesToCopy
    [System.Collections.Hashtable]$originalFileExtensions
    [System.Collections.Hashtable]$filesToFix
    [System.Collections.Hashtable]$fileNamesPairedFileLocation

    $resultFlag = $false

    if ($mode -eq "PrepForEmail"){

        $filesToCopy = Get-FilesFromBrowse

        $filePath = New-FolderToPlaceFiles

        Copy-Files $filesToCopy $filePath

        Save-OriginalFileExtensionsAndNamesToCSV -Path $filePath

        $resultFlag = Set-FileExtensionToTxt $filesToCopy $filePath

        if ($resultFlag -eq $true){

            Write-host "Process completed succesfully"

        }else {
            Write-host "Process failed!"
        }

    }elseif ($mode -eq "FixAfterEmail"){

        [string]$savePath = $PSScriptRoot
        [boolean]$copyFlag = $false
        $filePath = "$PSScriptRoot\FilesToCopy"

        $filesToConvert = Get-FilesToConvertBack -path $filePath

        $originalFileExtensions = Get-FileNamePairedOriginalFileExtensionsFromCSV

        $filesToCopy = Set-FileExtensionToOriginal -myFiles $filesToConvert -Path $filePath -originalFileExtensions $originalFileExtensions

        $copyFlag = Copy-Files $filesToCopy $savePath

        if ($copyFlag -eq $true){

            Clear-Folder -path $filePath
            Write-host "Process Complete!"
        }else{
            Write-host "Copy Failed, delete stopped"
            $resultFlag = $false
        }

    }elseif ($mode -eq "QuickChange"){

        $fileNamesPairedFileLocation = Get-FilesFromBrowse

        $extension = Get-ExtensionFromUser

        $successFlag = Set-ExtensionFromUser -extension $extension -fileNamesPairedFileLocation $fileNamesPairedFileLocation

        if($successFlag -eq $true){
            Write-host "Process completed successfully"
        }else {
            Write-host "Process failed!"
        }

    }else {
        Write-host "Invalid Argument"
        Write-host "*----------------------*"
        Write-host "To Prep for Email Delivery : PrepForEmail"
        Write-host "To fix after Email Delivery: FixAfterEmail"
    }
}

Main

