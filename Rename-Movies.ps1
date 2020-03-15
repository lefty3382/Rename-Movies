$VerbosePreference = "Continue"
$Server = "192.168.0.64"
$DownloadsDirectory = "\\$Server\storage\Film\_New\_Movies\"
$YearRegex = "[1|2][9|0][0-9][0-9]"
$ParenYearRegex = "^[(][1|2][9|0][0-9][0-9][)]$"
$OriginalNameRegex = "[ ][(][1|2][9|0][0-9][0-9][)]$"

#Get all subfolders from target folder
foreach ($MovieFolder in (gci $DownloadsDirectory))
{
    Write-Output "------------------"
    Write-Output "Folder: $($MovieFolder.name)"
    
    #Set variables to empty
    $TempFileName = ""
    $TempFolderName = ""
    $FolderSplit = @()
    $FileSplit = @()
    $NewName = ""
    $Element = ""
    $Year = ""
    $SRTDone = $false
    $ok = $false

    #Validate movie folder is not already renamed
    if (($MovieFolder.Name -match $OriginalNameRegex) -and
        ($MovieFolder.Name -notlike "*1080p*") -and
        ($MovieFolder.Name -notlike "*720p*") -and
        ($MovieFolder.Name -notlike "*BluRay*") -and
        ($MovieFolder.Name -notlike "*WEBrip*") -and
        ($MovieFolder.Name -notlike "*H264*") -and
        ($MovieFolder.Name -notlike "*RARBG*") -and
        ($MovieFolder.Name -notlike "*AMZN*"))
    {
        Write-Output "Folder already appears to be renamed"
        Write-Output "Validating file names..."

        #If movie title starts with "The", move to end of new title name
        if ($FolderSplit[0] -eq "The")
        {
            Write-Output "Movie folder name starts with `"The`""
            $ok = $true
            for ($i=1; $i -lt ($FolderSplit.Count-1); $i++)
            {
                $NewName += $FolderSplit[$i] + " "
            }
            $NewName = $NewName.TrimEnd(" ")
            $NewName += ", The "
            $NewName += $FolderSplit[-1]
            Write-Output "NewName: $NewName"
        }
        else
        {
            Write-Warning "Folder appears to already be renamed, moving on..."
            continue
        }
    }
    else
    {
        #Determine new name from existing parent folder name
        $TempFolderName = $MovieFolder.name
        $TempFolderName = $TempFolderName.replace('.',' ')
        $TempFolderName = $TempFolderName.replace('[','')
        $TempFolderName = $TempFolderName.replace(']','')
        $FolderSplit = $TempFolderName.split(' ')

        foreach ($string in $FolderSplit)
        { 
            if ($ok -eq $false)
            {
                Write-Verbose "Parsing string: $string"
                Write-Verbose "Matching `"$string`" to year regex expression `"$YearRegex`""
                if ($string -match $YearRegex)
                {
                    Write-Verbose "String `"$string`" matches regex `"$YearRegex`""
                    if ($string -match $ParenYearRegex)
                    {
                        Write-Verbose "String `"$string`" matches regex `"$ParenYearRegex`""
                        $Year = $string
                    }
                    else
                    {
                        #Remove all non-digit characters from string, add parentheses
                        $Year = $string -replace '[\D]', ''
                        $Year = $Year -replace "$Year", "($Year)"
                    }

                    Write-Verbose "Final year string: $year"
                    $ok = $true
                    $Element = [array]::indexof($FolderSplit,$string)

                    #If movie title starts with "The", move to end of new title name
                    if ($FolderSplit[0] -eq "The")
                    {
                        for ($i=1; $i -lt $element; $i++)
                        {
                            $NewName += $FolderSplit[$i] + " "
                        }
                        $NewName = $NewName.TrimEnd(" ")
                        $NewName += ", The "
                        $NewName += $Year
                    }
                    else
                    {
                        for ($i=0; $i -lt $element; $i++)
                        {
                            $NewName += $FolderSplit[$i] + " "
                        }
                        $NewName += $Year
                    }
                    Write-Verbose "NewName: $NewName"
                }
                else
                {
                    Write-Verbose "String `"$string`" does NOT match regex `"$YearRegex`""
                }
            }
        }
    }

    #Continue if new name exists, else skip to next loop
    If ($ok -eq $false)
    {
        Write-Warning "Year could not be parsed from folder name, skipping"
        continue
    }

    #Get child items in target folder
    $MovieChildItems = Get-ChildItem -LiteralPath $MovieFolder.FullName
    foreach ($MovieChildItem in $MovieChildItems)
    {
        Write-Output "Parsing item `"$($MovieChildItem.name)`""
        if ($MovieChildItem.PSIsContainer -eq $false)
        {
            #Delete .TXT .EXE .NFO files
            if (($MovieChildItem.name -like "*.nfo") -or
                ($MovieChildItem.name -like "*.exe") -or
                ($MovieChildItem.name -like "*.txt"))
            {
                Write-Output "Removing file: `"$($MovieChildItem.name)`""
                try
                {
                    Remove-Item -LiteralPath $MovieChildItem.fullname -Force -ErrorAction Stop
                    Write-Output "Successfully removed file: `"$($MovieChildItem.name)`""
                }
                catch
                {
                    Write-Warning "Failed to remove: `"$($MovieChildItem.name)`""
                }
            }

            #if video or subtitle file, rename
            elseif (($MovieChildItem.name -like "*.mp4") -or
                    ($MovieChildItem.name -like "*.mkv") -or
                    ($MovieChildItem.name -like "*.avi") -or
                    ($MovieChildItem.name -like "*.srt") -or
                    ($MovieChildItem.name -like "*.idx") -or
                    ($MovieChildItem.name -like "*.com") -or
                    ($MovieChildItem.name -like "*.sub"))
            {
                Write-Output "Renaming file: `"$($MovieChildItem.name)`""
                $FileExt = $MovieChildItem.name.split('.')[-1]
                $NewSubItemName = $NewName + "." + $FileExt
                Write-Output "New file name: $NewSubItemName"
                try
                {
                    Rename-Item -LiteralPath $MovieChildItem.fullname -NewName $NewSubItemName -ErrorAction Stop
                    Write-Output "Successfully renamed file: `"$($MovieChildItem.name)`""
                }
                catch
                {
                    Write-Warning "Failure renaming file: `"$($MovieChildItem.name)`""
                }
                #Verify file rename
                finally
                {
                    Write-Output "Verifying file rename: `"$($MovieChildItem.name)`"..."
                    if (Test-Path -LiteralPath $MovieChildItem.FullName)
                    {
                        Write-Warning "Old file name detected"
                        Write-Warning "Re-attempting file rename"
                        $NewFullPath = Join-Path -Path $MovieChildItem.DirectoryName -ChildPath $NewSubItemName
                        try
                        {
                            Move-Item -LiteralPath $MovieChildItem.fullname -Destination $NewFullPath -ErrorAction Stop
                            Write-Output "Successfully renamed file: `"$($MovieChildItem.name)`""
                        }
                        catch
                        {
                            Write-Warning "Failure renaming file: `"$($MovieChildItem.name)`""
                        }
                    }
                    elseif (Test-Path -LiteralPath "$($MovieFolder.FullName)\$NewSubItemName")
                    {
                        Write-Output "File rename verified: `"$NewSubItemName`""
                    }
                }
            }
        }
                
        #Subs folder: rename and move subtitle files to parent folder
        elseif ($MovieChildItem.PSIsContainer -eq $true)
        {
            if ($MovieChildItem.Name -eq "Subs")
            {
                Write-Output "Subtitle folder found!"
                $SubsFolder = Get-ChildItem $MovieChildItem.fullname
                foreach ($SubtitleFile in $SubsFolder)
                {
                    $Ext = $SubtitleFile.name.Split(".")[-1]
                    if ($SubtitleFile.name -like "*.srt")
                    {
                        if ($SubtitleFile.name -like "*Eng*")
                        {
                            if ($SRTDone -eq $false)
                            {
                                $NewSubtitleFileName = $SubtitleFile.Name.Replace($SubtitleFile.Name,$NewName)
                                Write-Output "Moving sub file to movie folder: `"$($SubtitleFile.Name)`""
                                Write-Output "New path: $($MovieFolder.fullname)\$NewSubtitleFileName.$Ext"
                                Move-Item -LiteralPath $SubtitleFile.fullname -Destination "$($MovieFolder.fullname)\$NewSubtitleFileName.$Ext"
                                $SRTDone = $true
                            }
                            else
                            {
                                Write-Output "Duplicate English .SRT sub file found"
                                Write-Output "Removing subtitle file: `"$($SubtitleFile.Name)`""
                                Remove-Item -LiteralPath $SubtitleFile.FullName -Force
                            }
                        }
                        else
                        {
                            Write-Output "Sub file: `"$($SubtitleFile.FullName)`""
                            Write-Output "Sub file name does not contain ENG string, so we delete it"
                            Remove-Item -LiteralPath $SubtitleFile.FullName -Force
                        }
                    }
                    else
                    {
                        $NewSubtitleFileName = $SubtitleFile.Name.Replace($SubtitleFile.Name,$NewName)
                        Write-Output "Moving subtitle file: `"$($SubtitleFile.FullName)`""
                        Write-Output "New path: $($MovieFolder.fullname)\$NewSubtitleFileName.$Ext"
                        Move-Item -LiteralPath $SubtitleFile.fullname -Destination "$($MovieFolder.fullname)\$NewSubtitleFileName.$Ext"
                    }
                }
                #Delete Subs folder
                Write-Output "Removing folder: `"$($MovieChildItem.FullName)`" folder"
                Remove-Item -LiteralPath $MovieChildItem.FullName -Recurse -Force
            }
            else
            {
                Write-Output "Removing folder: `"$($MovieChildItem.FullName)`" folder"
                Remove-Item -LiteralPath $MovieChildItem.FullName -Recurse -Force
            }
        }
    }

    #Rename target folder
    Write-Output "Renaming folder: `"$($MovieFolder.Name)`""
    Write-Output "New folder name: `"$NewName`""
    try
    {
        Rename-Item -LiteralPath $MovieFolder.FullName -NewName $NewName -ErrorAction Stop
        Write-Output "Successfully renamed movie folder"
    }
    catch
    {
        Write-Warning "Failed to rename movie folder"
        $error[0]
    }
}