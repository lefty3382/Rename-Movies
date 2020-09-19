# Rename-Movies
Rename downloaded movie files and folders
## Detailed Description
This script will rename downloaded movie files/folders to a format compatible with home media player software like Plex & Emby using regular expressions.  It assumes each movie file(s) under the main directory pointed to in the script parameter is in its own subdirectory.  It checks for the standard video file container formats (.mp4, .mkv, etc) and subtitle file formats (.srt, .sub, .idx) and renames those files to match the renaming of the movie folder.  It will remove any other files types including .EXE and .TXT files.

### Process
The renaming process is based primarily on identifying the year (1900-2099) of the movie from the file name.  The script will look for a four digit string matching a year and assume all text prior to that string is part of the movie name and all text after that string is removed.

#### Example
* American.Pastime.2007.1080p.WEBRip.x265-EXAMPLE

The script will take the above string from the folder containing the movie files and identify the string (2007) as the string matching a valid year.  All characters after the year, in this case ".1080p.WEBRip.x265-EXAMPLE" are removed.  The script then replaces all dots (".") with spaces and places the year in parenthesis.  The final movie string result is "American Pastime (2007)".
