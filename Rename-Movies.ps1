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
        Position = 0,
        HelpMessage = "Path must contain '\\192.168.10.2\storage\Film\_New'")]
    #[ValidatePattern("(?i)^(Z:|\\\\192\.168\.0\.64\\storage)\\Film\\_New\\.{1,200}$")]
    [ValidateScript({(Get-ChildItem "\\192.168.10.2\storage\Film\_New" | Select-Object -ExpandProperty FullName) -contains $_})]
    [string]$DownloadsDirectory = "\\192.168.10.2\storage\Film\_New\HD",

    # Test switch, no changes enforced but all console output is displayed
    [Parameter(
        Mandatory = $false,
        Position = 1)]
    [switch]$Test = $false,

    # If movie title starts with "The", move it to the end of the title after a comma
    [Parameter(
        Mandatory = $false,
        Position = 2)]
    [switch]$PutTheAtTheEnd = $false
)

#ScriptVersion = "1.0.7.1"

###################################
# Script Variables
###################################

$YearRegex = "^[1|2][9|0][0-9][0-9]$"
$OriginalNameRegex = "[ ][(][1|2][9|0][0-9][0-9][)]$"
$AfterYearRegex = "^1080p|2160p|IMAX|REMASTERED|UNRATED|EXTENDED|DC|SHOUT|UNCUT|Colorized|DUBBED|FS|WS$"

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
        # Find release year in movie title string
        for ($a=0; $a -lt $FolderSplit.count; $a++)
        { 
            $CurrentString = $FolderSplit[$a]
            Write-Host "Parsing string: $CurrentString" -ForegroundColor Yellow
            Write-Host "Matching `"$CurrentString`" to year regex expression `"$YearRegex`"" -ForegroundColor Yellow
            if ($CurrentString -match $YearRegex)
            {
                Write-Host "String `"$CurrentString`" matches regex `"$YearRegex`"" -ForegroundColor Green
                $ProvisionalYearMatch = $CurrentString
                # Check for expected string after year
                if ($FolderSplit[$a+1] -match $AfterYearRegex)
                {
                    Write-Host "Found expected string after identified year: $($FolderSplit[$a+1])" -ForegroundColor Green
                    $ConfirmedYearMatch = $CurrentString
                    $a=100
                }
                else
                {
                    Write-Host "Did NOT find expected string after identified year: $($FolderSplit[$a+1])" -ForegroundColor Yellow
                    Write-Host "Continuing to parse through movie folder name" -ForegroundColor Yellow
                }
            }
            else
            {
                Write-Host "String `"$CurrentString`" does NOT match regex `"$YearRegex`"" -ForegroundColor Yellow
            }
        }
        
        # Finalize correct year
        if ($ConfirmedYearMatch)
        {
            $FinalYearMatch = $ConfirmedYearMatch
            Write-Host "Using confirmed year match: ($ConfirmedYearMatch)" -ForegroundColor Yellow
        }
        elseif ($ProvisionalYearMatch)
        {
            $FinalYearMatch = $ProvisionalYearMatch
            Write-Host "Using provisional year match: ($ProvisionalYearMatch)" -ForegroundColor Yellow
        }

        ### Get new name using identified year string

        #Remove all non-digit characters from string, add parentheses
        $Year = $FinalYearMatch -replace '[\D]', ''
        $Year = $Year -replace "$Year", "($Year)"

        Write-Host "Final year string: $year" -ForegroundColor Blue
        $Element = [array]::indexof($FolderSplit,$FinalYearMatch)

        #If movie title starts with "The", move to end of new title name
        if (($FolderSplit[0] -eq "The") -and ($PutTheAtTheEnd))
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

        # Add year to end of new string
        $NewName += $Year

        # Fix common issues
        $NewName = $NewName -creplace (" Of "," of ")
        $NewName = $NewName -replace (" Mr "," Mr. ")
        $NewName = $NewName -replace (" Mrs "," Mrs. ")
        $NewName = $NewName -replace ("Mr ","Mr. ")
        $NewName = $NewName -replace ("Mrs ","Mrs. ")
        Write-Host "NewName: $NewName" -ForegroundColor Blue
    }
    
    end
    {
        return $NewName
    }
}

###################################
# Main
###################################

Write-Host "Test mode: $Test" -ForegroundColor Yellow

#Get all subfolders from target folder
foreach ($MovieFolder in (Get-ChildItem $DownloadsDirectory))
{
    Write-Host "------------------"
    Write-Host "Folder: $($MovieFolder.name)"
    
    #Set variables to empty
    $TempFolderName = ""
    $FolderSplit = @()
    $NewName = ""
    $Element = ""
    $Year = ""
    $SRTDone = $false

    # Get new movie folder name if not already renamed
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
        Write-Host "Folder appears to already be renamed, moving on..." -ForegroundColor Yellow
        continue
    }
    else
    {
        try
        {
            $NewName = Get-NewMovieName -OriginalMovieString $MovieFolder.Name -ErrorAction Stop
            Write-Host "Final movie name: $NewName" -ForegroundColor Blue
        }
        catch
        {
            Write-Host "New movie name could not be determined" -ForegroundColor Red
            $Answer = Read-Host "Continue?"
            if ($Answer -match [Yy]) { Continue }
            else { exit }
        }
    }

    $NewFolderName = $NewName
    # If 4K movie, add 4K to end of new name for files
    if ($MovieFolder.Name -like "*.2160p.*")
    {
        $NewName = $NewName + " - UHD"
    }

    #Get child items in target folder
    $MovieChildItems = Get-ChildItem -LiteralPath $MovieFolder.FullName
    foreach ($MovieChildItem in $MovieChildItems)
    {
        Write-Host "Parsing item: `"$($MovieChildItem.name)`"" -ForegroundColor Yellow
        if ($MovieChildItem.PSIsContainer -eq $false)
        {
            #Delete .TXT .EXE .NFO files
            if (($MovieChildItem.name -like "*.nfo") -or
                ($MovieChildItem.name -like "*.exe") -or
                ($MovieChildItem.name -like "*.txt"))
            {
                Write-Host "Removing file: `"$($MovieChildItem.name)`"" -ForegroundColor Yellow
                try
                {
                    Remove-Item -LiteralPath $MovieChildItem.fullname -Force -ErrorAction Stop -WhatIf:$Test
                    Write-Host "Successfully removed file: `"$($MovieChildItem.name)`"" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "Failed to remove: `"$($MovieChildItem.name)`"" -ForegroundColor Red
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
                Write-Host "Renaming file: `"$($MovieChildItem.name)`"" -ForegroundColor Yellow
                $FileExt = $MovieChildItem.name.split('.')[-1]
                $NewSubItemName = $NewName + "." + $FileExt
                Write-Host "New file name: $NewSubItemName"
                try
                {
                    Rename-Item -LiteralPath $MovieChildItem.fullname -NewName $NewSubItemName -ErrorAction Stop -WhatIf:$Test
                    Write-Host "Successfully renamed file: `"$($MovieChildItem.name)`"" -ForegroundColor Green
                }
                catch
                {
                    Write-Host "Failure renaming file: `"$($MovieChildItem.name)`"" -ForegroundColor Red
                }
                #Verify file rename
                finally
                {
                    Write-Host "Verifying file rename: `"$($MovieChildItem.name)`"..."
                    if (Test-Path -LiteralPath $MovieChildItem.FullName)
                    {
                        Write-Host "Old file name detected" -ForegroundColor Red
                        Write-Host "Re-attempting file rename" -ForegroundColor Red
                        $NewFullPath = Join-Path -Path $MovieChildItem.DirectoryName -ChildPath $NewSubItemName
                        try
                        {
                            Move-Item -LiteralPath $MovieChildItem.fullname -Destination $NewFullPath -ErrorAction Stop -WhatIf:$Test
                            Write-Host "Successfully renamed file: `"$($MovieChildItem.name)`"" -ForegroundColor Green
                        }
                        catch
                        {
                            Write-Host "Failure renaming file: `"$($MovieChildItem.name)`"" -ForegroundColor Red
                        }
                    }
                    elseif (Test-Path -LiteralPath "$($MovieFolder.FullName)\$NewSubItemName")
                    {
                        Write-Host "File rename verified: `"$NewSubItemName`"" -ForegroundColor Green
                    }
                }
            }
        }
                
        #Subs folder: rename and move subtitle files to parent folder
        elseif ($MovieChildItem.PSIsContainer -eq $true)
        {
            if ($MovieChildItem.Name -eq "Subs")
            {
                Write-Host "Subtitle folder found!" -ForegroundColor Blue
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
                                Write-Host "Moving sub file to movie folder: `"$($SubtitleFile.Name)`"" -ForegroundColor Yellow
                                Write-Host "New path: $($MovieFolder.fullname)\$NewSubtitleFileName.$Ext" -ForegroundColor Yellow
                                Move-Item -LiteralPath $SubtitleFile.fullname -Destination "$($MovieFolder.fullname)\$NewSubtitleFileName.$Ext" -WhatIf:$Test
                                $SRTDone = $true
                            }
                            else
                            {
                                Write-Host "Duplicate English .SRT sub file found" -ForegroundColor Yellow
                                Write-Host "Removing subtitle file: `"$($SubtitleFile.Name)`"" -ForegroundColor Yellow
                                Remove-Item -LiteralPath $SubtitleFile.FullName -Force -WhatIf:$Test
                            }
                        }
                        else
                        {
                            Write-Host "Sub file: `"$($SubtitleFile.FullName)`"" -ForegroundColor Yellow
                            Write-Host "Sub file name does not contain ENG string, so we delete it" -ForegroundColor Yellow
                            Remove-Item -LiteralPath $SubtitleFile.FullName -Force -WhatIf:$Test
                        }
                    }
                    else
                    {
                        $NewSubtitleFileName = $SubtitleFile.Name.Replace($SubtitleFile.Name,$NewName)
                        Write-Host "Moving subtitle file: `"$($SubtitleFile.FullName)`"" -ForegroundColor Yellow
                        Write-Host "New path: $($MovieFolder.fullname)\$NewSubtitleFileName.$Ext" -ForegroundColor Yellow
                        Move-Item -LiteralPath $SubtitleFile.fullname -Destination "$($MovieFolder.fullname)\$NewSubtitleFileName.$Ext" -WhatIf:$Test
                    }
                }
                #Delete Subs folder
                Write-Host "Removing folder: `"$($MovieChildItem.FullName)`" folder" -ForegroundColor Yellow
                Remove-Item -LiteralPath $MovieChildItem.FullName -Recurse -Force -WhatIf:$Test
            }
            else
            {
                Write-Host "Removing folder: `"$($MovieChildItem.FullName)`" folder" -ForegroundColor Yellow
                Remove-Item -LiteralPath $MovieChildItem.FullName -Recurse -Force -WhatIf:$Test
            }
        }
    }
    #Rename target folder
    Write-Host "Renaming folder: `"$($MovieFolder.Name)`"" -ForegroundColor Yellow
    Write-Host "New folder name: `"$NewFolderName`""
    try
    {
        Rename-Item -LiteralPath $MovieFolder.FullName -NewName $NewFolderName -ErrorAction Stop -WhatIf:$Test
        Write-Host "Successfully renamed movie folder" -ForegroundColor Green
    }
    catch
    {
        Write-Host "Failed to rename movie folder" -ForegroundColor Red
        $error[0]
    }
}