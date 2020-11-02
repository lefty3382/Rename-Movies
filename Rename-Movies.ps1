<#	
	.NOTES
	===========================================================================
	 Created with: 	Visual Studio Code 1.47.2
	 Created on:   	12/08/2019 11:09 AM
	 Created by:   	Jason Witters
	 Organization: 	Witters Inc.
	 Filename:     	Rename-Movies.ps1
	===========================================================================
	.DESCRIPTION
		Rename downloaded movie files and folders.
#>

[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $false,
        Position = 0
    )]
    [string]$DownloadsDirectory,
    # Test switch, no changes enforced but all console output is displayed
    [Parameter(
        Mandatory = $false,
        Position = 1
    )]
    [switch]$Test = $false
)

#ScriptVersion = "1.0.3.0"

###################################
# Script Variables
###################################

#$VerbosePreference = "Continue"
$Server = "192.168.0.64"
$YearRegex = "^[1|2][9|0][0-9][0-9]$"
#$ParenYearRegex = "^[(][1|2][9|0][0-9][0-9][)]$"
$OriginalNameRegex = "[ ][(][1|2][9|0][0-9][0-9][)]$"
$AfterYearRegex = "^1080p|2160p|REMASTERED|UNRATED|EXTENDED|DC|SHOUT|UNCUT|Colorized|DUBBED|FS|WS$"

if (!$DownloadsDirectory)
{
    $DownloadsDirectory = "\\$Server\storage\Film\_New\_Movies\"
}

###################################
# Script Functions
###################################

function Get-NewMovieName {
    [CmdletBinding()]
    param (
        [Parameter(
        Mandatory = $true,
        Position = 0
    )]
    [string]$OriginalMovieString
    )
    
    begin
    {
        # Replace dots and brackets with spaces, then split into array
        $TempFolderName = $OriginalMovieString
        $TempFolderName = $TempFolderName.replace('.',' ')
        $TempFolderName = $TempFolderName.replace('[','')
        $TempFolderName = $TempFolderName.replace(']','')
        $FolderSplit = $TempFolderName.split(' ')
    }
    
    process
    {
        # Find movie year, keep only what's before
        for ($a=0; $a -lt $FolderSplit.count; $a++)
        { 
            $CurrentString = $FolderSplit[$a]
            Write-Verbose "Parsing string: $CurrentString"
            Write-Verbose "Matching `"$CurrentString`" to year regex expression `"$YearRegex`""
            if ($CurrentString -match $YearRegex)
            {
                Write-Verbose "String `"$CurrentString`" matches regex `"$YearRegex`""
                
                # Check for expected string after year
                if ($FolderSplit[$a+1] -match $AfterYearRegex)
                {
                    Write-Verbose "Found expected string after identified year: $($FolderSplit[$a+1])"
                    
                    #Remove all non-digit characters from string, add parentheses
                    $Year = $CurrentString -replace '[\D]', ''
                    $Year = $Year -replace "$Year", "($Year)"

                    Write-Verbose "Final year string: $year"
                    $Element = [array]::indexof($FolderSplit,$CurrentString)

                    #If movie title starts with "The", move to end of new title name
                    if ($FolderSplit[0] -eq "The")
                    {
                        for ($i=1; $i -lt $element; $i++)
                        {
                            $NewName += $FolderSplit[$i] + " "
                        }
                        $NewName = $NewName.TrimEnd(" ")
                        $NewName += ", The "
                    }
                    else
                    {
                        for ($i=0; $i -lt $element; $i++)
                        {
                            $NewName += $FolderSplit[$i] + " "
                        }
                    }
                    $NewName += $Year
                    $NewName = $NewName -creplace ("Of","of")
                    Write-Verbose "NewName: $NewName"
                    $a=100
                }
                else
                {
                    Write-Verbose "Did NOT find expected string after identified year: $($FolderSplit[$a+1])"
                    Write-Verbose "Continuing to parse through movie folder name"
                }
            }
            else
            {
                Write-Verbose "String `"$CurrentString`" does NOT match regex `"$YearRegex`""
            }
        }
    }
    
    end
    {
        return $NewName
    }
}

###################################
# Main
###################################

#Get all subfolders from target folder
foreach ($MovieFolder in (Get-ChildItem $DownloadsDirectory))
{
    Write-Output "------------------"
    Write-Output "Folder: $($MovieFolder.name)"
    
    #Set variables to empty
    $TempFolderName = ""
    $FolderSplit = @()
    $NewName = ""
    $Element = ""
    $Year = ""
    $SRTDone = $false

    #Validate movie folder is not already renamed
    if (($MovieFolder.Name -match $OriginalNameRegex) -and
        ($MovieFolder.Name -notlike "*2160p*") -and
        ($MovieFolder.Name -notlike "*1080p*") -and
        ($MovieFolder.Name -notlike "*720p*") -and
        ($MovieFolder.Name -notlike "*BluRay*") -and
        ($MovieFolder.Name -notlike "*WEBrip*") -and
        ($MovieFolder.Name -notlike "*H264*") -and
        ($MovieFolder.Name -notlike "*x265*") -and
        ($MovieFolder.Name -notlike "*RARBG*") -and
        ($MovieFolder.Name -notlike "*HULU*") -and
        ($MovieFolder.Name -notlike "*DSNP*") -and
        ($MovieFolder.Name -notlike "*ATVP*") -and
        ($MovieFolder.Name -notlike "*AMZN*"))
    {
        Write-Warning "Folder appears to already be renamed, moving on..."
        continue
    }
    else
    {
        try
        {
            $NewName = Get-NewMovieName -OriginalMovieString $MovieFolder.Name -ErrorAction Stop
            Write-Output "Final movie name: $NewName"
        }
        catch
        {
            Write-Warning "New movie name could not be determined"
            $Answer = Read-Host "Continue?"
            if ($Answer -match [Yy]) { Continue }
            else { exit }
        }
    }

    #Get child items in target folder
    $MovieChildItems = Get-ChildItem -LiteralPath $MovieFolder.FullName
    foreach ($MovieChildItem in $MovieChildItems)
    {
        Write-Output "Parsing item: `"$($MovieChildItem.name)`""
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
                    Remove-Item -LiteralPath $MovieChildItem.fullname -Force -ErrorAction Stop -WhatIf:$Test
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
                    Rename-Item -LiteralPath $MovieChildItem.fullname -NewName $NewSubItemName -ErrorAction Stop -WhatIf:$Test
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
                            Move-Item -LiteralPath $MovieChildItem.fullname -Destination $NewFullPath -ErrorAction Stop -WhatIf:$Test
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
                                Move-Item -LiteralPath $SubtitleFile.fullname -Destination "$($MovieFolder.fullname)\$NewSubtitleFileName.$Ext" -WhatIf:$Test
                                $SRTDone = $true
                            }
                            else
                            {
                                Write-Output "Duplicate English .SRT sub file found"
                                Write-Output "Removing subtitle file: `"$($SubtitleFile.Name)`""
                                Remove-Item -LiteralPath $SubtitleFile.FullName -Force -WhatIf:$Test
                            }
                        }
                        else
                        {
                            Write-Output "Sub file: `"$($SubtitleFile.FullName)`""
                            Write-Output "Sub file name does not contain ENG string, so we delete it"
                            Remove-Item -LiteralPath $SubtitleFile.FullName -Force -WhatIf:$Test
                        }
                    }
                    else
                    {
                        $NewSubtitleFileName = $SubtitleFile.Name.Replace($SubtitleFile.Name,$NewName)
                        Write-Output "Moving subtitle file: `"$($SubtitleFile.FullName)`""
                        Write-Output "New path: $($MovieFolder.fullname)\$NewSubtitleFileName.$Ext"
                        Move-Item -LiteralPath $SubtitleFile.fullname -Destination "$($MovieFolder.fullname)\$NewSubtitleFileName.$Ext" -WhatIf:$Test
                    }
                }
                #Delete Subs folder
                Write-Output "Removing folder: `"$($MovieChildItem.FullName)`" folder"
                Remove-Item -LiteralPath $MovieChildItem.FullName -Recurse -Force -WhatIf:$Test
            }
            else
            {
                Write-Output "Removing folder: `"$($MovieChildItem.FullName)`" folder"
                Remove-Item -LiteralPath $MovieChildItem.FullName -Recurse -Force -WhatIf:$Test
            }
        }
    }

    #Rename target folder
    Write-Output "Renaming folder: `"$($MovieFolder.Name)`""
    Write-Output "New folder name: `"$NewName`""
    try
    {
        Rename-Item -LiteralPath $MovieFolder.FullName -NewName $NewName -ErrorAction Stop -WhatIf:$Test
        Write-Output "Successfully renamed movie folder"
    }
    catch
    {
        Write-Warning "Failed to rename movie folder"
        $error[0]
    }
}