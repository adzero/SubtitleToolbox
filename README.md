# SubtitleToolbox
---
#### A PowerShell module with a set of functions to find and manage your subtitle files.
---

## Prerequisites : 

- PowerShell (V5 or above)
- Git (optional, but recommended to ease updates) 

## Module installation

Open a **PowerShell** console and go to your Windows user profile Documents folder : 

```PowerShell
PS C:\Users\adzero>  cd "$([environment]::getfolderpath(“mydocuments”))"
PS D:\Users\adzero\Documents> 
``` 

Create a directory called *WindowsPowerShell* and its children *Modules* change current directory to it : 

```PowerShell
PS D:\Users\adzero\Documents> New-Item -Name WindowsPowerShell\Modules -ItemType Container
PS D:\Users\adzero\Documents> cd .\WindowsPowerShell\Modules
PS D:\Users\adzero\Documents\WindowsPowerShell\Modules> 
```

Clone the git repository from **GitHub**
```PowerShell
PS D:\Users\adzero\Documents\WindowsPowerShell\Modules> git clone https://github.com/adzero/SubtitleToolbox.git
Cloning into 'SubtitleToolbox'...
remote: Counting objects: 52, done.
remote: Compressing objects: 100% (50/50), done.
remote: Total 52 (delta 29), reused 0 (delta 0)
Unpacking objects: 100% (52/52), done.
PS D:\Users\adzero\Documents\WindowsPowerShell\Modules>
```
Module is now ready to be used. 

## Module update

To get fixes and updates from **GitHub**, open a command prompt or a PowerShell console inside the module folder : 

```PowerShell
PS C:\Users\adzero> cd "$([environment]::getfolderpath(“mydocuments”))"
PS D:\Users\adzero\Documents> cd .\WindowsPowerShell\Modules\SubtitleToolbox\
PS D:\Users\adzero\Documents\WindowsPowerShell\Modules\SubtitleToolbox\>
``` 

Then type the following command to get any update : 

```PowerShell
PS D:\Users\adzero\Documents\WindowsPowerShell\Modules\SubtitleToolbox\> git pull origin -f
Already up-to-date.
PS D:\Users\adzero\Documents\WindowsPowerShell\Modules\SubtitleToolbox\> 
``` 


## Module usage
#### Module loading

If module autoloading (for **PowerShell 3.0** or above) is off, open a **PowerShell** console or use one already opened and type the following command : 

```PowerShell
PS C:\Users\adzero> Import-Module SubtitleToolbox
PS C:\Users\adzero>
```
Now module functions described below can be used. 

#### Get-VideoFiles

```Get-VideoFiles``` function gets video files in or under in a specified directory.

If you search only videos with (or without) external subtitles you can use the ```-Subtitles``` (or ```-NoSubtitles```) option. With these two options you can also specify the language code of the existing (or missing) subtitles with the ```-LanguageCode``` parameter. 
To search all files under the given path, you must use the ```-Recurse``` option. 
Supported files are : MPG, AVI, MP4, MKV, TS

#### Examples : 

##### Example 1: Search video files in a specific directory.

```PowerShell
PS C:\> Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01"

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv
```

##### Example 2: Search video files in a specific directory, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01" | Get-VideoFiles -Path 

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv
```

##### Example 3: Search all video files under a specific path.

```PowerShell
PS C:\> Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S02


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       10/05/2017     16:19      210497922 My.Show.S02E01.HDTV.x264-FLEET.mkv
```

##### Example 4: Search all video files without subtitles under a specific path.

```PowerShell
PS C:\> Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -NoSubtitles -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S02


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       10/05/2017     16:19      210497922 My.Show.S02E01.HDTV.x264-FLEET.mkv
```

##### Example 5: Search all video files with subtitles (default language) under a specific path.

```PowerShell
PS C:\> Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -Subtitles -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv
```

##### Example 6: Search all video files with english subtitles under a specific path.

```PowerShell
PS C:\> Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -Subtitles -LanguageCode "en" -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv
```

#### Get-SubtitleFiles

```Get-SubtitleFiles``` function gets subtitle files in or under in a specified directory.
To search subtitles in a specific language, use the ```-LanguageCode``` parameter to specify the language code. 
To search all subtitles under the given path, you must use the ```-Recurse``` option.
Only SRT and ASS subtitles will match.

#### Examples : 

##### Example 1: Search subtitle files in a specific directory.

```PowerShell
PS C:\> Get-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01"

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt
```

##### Example 2: Search subtitle files in a specific directory, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01" | Get-SubtitleFiles

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt
```

##### Example 3: Search all subtitles located under a specific path.

```PowerShell
PS C:\> Get-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt
```

##### Example 4: Search english subtitle files in a specific directory.

```PowerShell
PS C:\> Get-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01" -LanguageCode "en"

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt
```

#### Remove-SubtitleFiles

```Remove-SubtitleFiles``` function removes subtitle files in or under in a specified directory.
To remove subtitles in a specific language, use the ```-LanguageCode``` parameter to specify the language code. 
To remove all subtitles under the given path, you must use the ```-Recurse``` option.
With the ```-Confirm``` option a prompt ask for confirmation before actual file deletion. 
The ```-WhatIf``` option displays all files to delete without actual deletion. 
Only SRT and ASS subtitles will be removed.

#### Examples : 

##### Example 1: Remove subtitle files in a specific directory.

```PowerShell
PS C:\> Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01"
```

##### Example 2: Remove subtitle files in a specific directory after confirmation.

```PowerShell
PS C:\> Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01" -Confirm

Confirm
Are you sure you want to perform this action?
Performing the operation « Remove file » on target  « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt ».
[Y] Yes  [A] Yes to All  [N] No  [L] No to All [S] Suspend  [?] Help (default is « Y ») : L
```

##### Example 3: List subtitle files to remove in a specific directory.

```PowerShell
PS C:\> Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01" -WhatIf
What if : Performing the operation « Remove file » on target « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt ».
What if : Performing the operation « Remove file » on target « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E02.HDTV.x264-FLEET.en.srt ».
What if : Performing the operation « Remove file » on target « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E03.HDTV.x264-FLEET.en.srt ».
```

##### Example 4: Remove subtitle files in a specific directory, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01" | Remove-SubtitleFiles
```

##### Example 5: Remove subtitle files located under a specific path.

```PowerShell
PS C:\> Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows" -Recurse
```

##### Example 6: Remove english subtitle files in a specific directory.

```PowerShell
PS C:\> Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\S01" -LanguageCode "en"
```

#### Set-SubtitleFileLanguage

```Set-SubtitleFileLanguage``` function adds (or removes) an ISO language code at the end of a subtitle file. 
Only valid ISO codes are allowed. Use of an invalid/unknown code will raise an error. 
Only SRT and ASS files are supported.

#### Examples : 

##### Example 1: Set default language for a subtitle file.

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt"
```

##### Example 2: Set default language for a subtitle file, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt"
```

##### Example 3: Set english language for a subtitle file.

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt" -LanguageCode "en"
```

##### Example 4: Force default language for a subtitle file.

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -LanguageCode "en"
```

##### Example 5: Remove english language for a subtitle file.

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -Remove
```

##### Example 6: Remove english language for a subtitle file after confirmation.

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -Remove -Confirm

Confirm
Are you sure you want to perform this action?
Performing the operation « Rename file » on target  « Element: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt Destination: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt».
[Y] Yes  [A] Yes to All  [N] No  [L] No to All [S] Suspend  [?] Help (default is « Y ») : L
```

##### Example 7: Display changes when setting default language for a subtitle file without actually changing the file name. 

```PowerShell
PS C:\> Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -WhatIf
What if : Performing the operation « Rename file » on target « Element: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt Destination: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt ».
```

#### Get-EpisodeFileInfo

```Get-EpisodeFileInfo``` function parses a video file name and extracts information for subtitles search.
Result is a custom object with the following properties (when found) : 

 - Path
 - FileName
 - ShowTitle
 - Season
 - Episode
 - Resolution
 - ProperRepack
 - Source
 - ReleaseGroup
 
The ValidForSearch property is set to True when enough information have been found for a subtitle search.  

#### Examples : 

##### Example 1: Get the episode information from a file.

```PowerShell
PS C:\> Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv"


Path           : C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv
FileName       : My.Show.S01E01.HDTV.x264-FLEET.mkv
ShowTitle      : My Show
Season         : 1
Episode        : 1
Source         : HDTV
ReleaseGroup   : FLEET
ValidForSearch : True 

```

##### Example 2: Get the episode information from a file, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" | Get-EpisodeFileInfo


Path           : C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv
FileName       : My.Show.S01E01.HDTV.x264-FLEET.mkv
ShowTitle      : My Show
Season         : 1
Episode        : 1
Source         : HDTV
ReleaseGroup   : FLEET
ValidForSearch : True 

```

#### Import-EpisodeSubtitles

```Get-EpisodeFileInfo``` function downloads subtitles matching an episode from Sous-Titres.eu website.
Search is based on the show title, season and episode numbers, language code.
Result is a FileInfo object collection of the downloaded subtitles files.

#### Examples : 

##### Example 1: Download all english subtitles for an episode.

```PowerShell
PS C:\> Import-EpisodeSubtitles -ShowTitle "My Show" -Season 1 -Episode 1 -LanguageCode "en"

    Directory: C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt

```

##### Example 2: Download all english subtitles for an episode, using the pipeline to pass the episode information parameters values.

```PowerShell
PS C:\> Get-EpisodeFileInfo -Path "My.Show.S01E01.HDTV.x264-FLEET.mkv" | Import-EpisodeSubtitles -LanguageCode "en"

    Directory: C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt

```

#### Get-SubtitleFileScore

```Get-SubtitleFileScore``` function gets the matching information and score of a subtitle file.
Score is computed with episode information provided by the Get-EpisodeFileInfo function.
Maximum score is 100.
Result is a custom object with the following properties : 

 - Path
 - FileName
 - MatchShowTitle
 - MatchSeason
 - MatchEpisode
 - MatchResolution
 - MatchSource
 - MatchReleaseGroup
 - MatchLanguage
 - MatchTag
 - MatchProperRepack

#### Examples : 

##### Example 1: Get a subtitle file score for a video file and english language.

```PowerShell
PS C:\> Get-SubtitleFileScore -Path "C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET\My.Show.S01E01.HDTV.x264-FLEET.fr.srt" -EpisodeInformation (Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv") -LanguageCode "en"


Path              : C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET\My.Show.S01E01.HDTV.x264-FLEET.fr.srt
FileName          : My.Show.S01E01.HDTV.x264-FLEET.fr.srt
MatchShowTitle    : True
MatchSeason       : True
MatchEpisode      : True
MatchResolution   : False
MatchSource       : True
MatchReleaseGroup : True
MatchLanguage     : True
MatchTag          : False
MatchProperRepack : False
Score             : 97
SubtitleType      : srt

```

##### Example 2: Get a subtitle file score for a video file and english language, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> Get-Item -Path "C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET\My.Show.S01E01.HDTV.x264-FLEET.fr.srt" | Get-SubtitleFileScore -EpisodeInformation (Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv") -LanguageCode "en"


Path              : C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET\My.Show.S01E01.HDTV.x264-FLEET.fr.srt
FileName          : My.Show.S01E01.HDTV.x264-FLEET.fr.srt
MatchShowTitle    : True
MatchSeason       : True
MatchEpisode      : True
MatchResolution   : False
MatchSource       : True
MatchReleaseGroup : True
MatchLanguage     : True
MatchTag          : False
MatchProperRepack : False
Score             : 97
SubtitleType      : srt

```

#### Find-EpisodeSubtitles

```Find-EpisodeSubtitles``` function finds the best subtitles in a given language for an episode.

#### Examples : 

##### Example 1: Get the best subtitles in default language for a video file.

```PowerShell
PS C:\> $subtitles = Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv"
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.FR.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'
```

##### Example 2: Get the best subtitles in default language for a video file, using the pipeline to pass the Path parameter value.

```PowerShell
PS C:\> $subtitles = Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" | Find-EpisodeSubtitles 
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.FR.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'
```

##### Example 3: Get the best english subtitles for a video file.

```PowerShell
PS C:\> $subtitles = Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" -LanguageCode "en"
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.EN.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.EN.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt'
```

##### Example 4: Get the best subtitles in default language for a video file and force overwrite of existing file. 

```PowerShell
PS C:\> Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" -Force | Out-Null
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.EN.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'
```

##### Example 5: Get the best subtitles in default language for a video file and displays its destination name without actually creating the file.  

```PowerShell
PS C:\> Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" -WhatIf | Out-Null
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.EN.srt'
WhatIf : Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'
```

#### Clear-DownloadCache

```Clear-DownloadCache``` function resets cache and deletes downloaded files

#### Examples : 

##### Example 1: Get the best subtitles in default language for a video file.

```PowerShell
PS C:\> Clear-DownloadCache
```