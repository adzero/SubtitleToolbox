$Script:DefaultLanguageCode = "fr"
$Script:SeriesRootUri = "https://www.sous-titres.eu/series"
$Script:TempDirectory = $ENV:TEMP
if(!$Global:SubtitleToolboxDownloads)
{
	$Global:SubtitleToolboxDownloads = @{}
}

<#
.SYNOPSIS
Gets subtitle files contained in a directory.

.DESCRIPTION
Gets subtitle files contained in a directory.
Only SRT and ASS subtitles will match.

.PARAMETER Path
The root directory for the subtitle search.

.PARAMETER LanguageCode
The language of the subtitles to search.

.PARAMETER Recurse
Search subtitles in subfolders. 

.INPUTS
System.String
System.IO.DirectoryInfo

.OUTPUTS
System.IO.FileInfo[]

.EXAMPLE 
Get-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows"

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt


Search subtitle files in a specific directory.

.EXAMPLE 
Get-Item "C:\Users\adzero\Videos\TV Shows" | Get-SubtitleFiles

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt


Search subtitle files in a specific directory, using the pipeline to pass the Path parameter value.

.EXAMPLE 
Get-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows" -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt


Search all subtitles located under a specific path.

.EXAMPLE 
Get-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01" -LanguageCode "en"

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:56            898 My.Show.S01E02.HDTV.x264-FLEET.en.srt
-a----       11/11/2016     17:58            742 My.Show.S01E03.HDTV.x264-FLEET.en.srt


Search english subtitle files in a specific directory.
#>
function Get-SubtitleFiles
{
	[CmdletBinding()]
	Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='The root directory for the subtitle search')]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Container))
		{
			throw "Provided path is not a valid directory path: '$value'" 
		}
		
		return $true
	})]
	$Path,
	[ValidateScript({ 
		$exists = Test-ISOLanguageName -Name $_
	
		if(!($_ -imatch "^[a-z]{2,3}$"))
		{
			throw "Language code must be a two or three letters string." 
		}
		elseif(!$exists)
		{
			throw "The provided language code is invalid or unknown : '$_'"
		}
		
		return $true
	})]
	[string]$LanguageCode,
	[Parameter(HelpMessage='Search subtitles in subfolders')]
	[switch]$Recurse)
	
	BEGIN
	{
	}
	PROCESS
	{	
		#Getting the root directory  
        if($Path -is [System.IO.DirectoryInfo])
        {
            $rootDirectory = $Path
        }
        else
        {
            $rootDirectory = Get-Item -LiteralPath $Path
        }
	
		#Setting subtitles match pattern
		if($LanguageCode)
		{
			$matchPattern = "\.$LanguageCode\.(srt|ass)$"
		}
		else
		{
			$matchPattern = "\.(srt|ass)$"
		}
		
		return ($rootDirectory | Get-ChildItem -File -Recurse:$Recurse | where {$_.Name -imatch $matchPattern})
	}
	END
	{
	}
}

<#
.SYNOPSIS
Removes subtitle files contained in a directory.

.DESCRIPTION
Removes subtitle files contained in a directory.
Only SRT and ASS subtitles will be removed.

.PARAMETER Path
The root directory with the subtitles to remove.

.PARAMETER LanguageCode
The language of the subtitles to remove.

.PARAMETER Recurse
Remove subtitles in subfolders. 

.PARAMETER Confirm
Asks for confirmation before removing a file.

.PARAMETER WhatIf
Lists actions without applying them. 

.INPUTS
System.String
System.IO.DirectoryInfo

.EXAMPLE 
Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01"


Remove subtitle files in a specific directory.

.EXAMPLE 
Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01" -Confirm

Confirm
Are you sure you want to perform this action?
Performing the operation « Remove file » on target  « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt ».
[Y] Yes  [A] Yes to All  [N] No  [L] No to All [S] Suspend  [?] Help (default is « Y ») : L


Remove subtitle files in a specific directory after confirmation.

.EXAMPLE 
Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows" -WhatIf

What if : Performing the operation « Remove file » on target « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt ».
What if : Performing the operation « Remove file » on target « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E02.HDTV.x264-FLEET.en.srt ».
What if : Performing the operation « Remove file » on target « C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E03.HDTV.x264-FLEET.en.srt ».


List subtitle files to remove in a specific directory.

.EXAMPLE 
Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01" | Remove-SubtitleFiles


Remove subtitle files in a specific directory, using the pipeline to pass the Path parameter value.

.EXAMPLE 
Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows" -Recurse


Remove subtitle files located under a specific path.

.EXAMPLE 
Remove-SubtitleFiles -Path "C:\Users\adzero\Videos\TV Shows\S01" -LanguageCode "en"


Remove english subtitle files in a specific directory.
#>
function Remove-SubtitleFiles
{
	[CmdletBinding()]
	Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='The root directory with the subtitles to remove')]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Container))
		{
			throw "Provided path is not a valid directory path: '$value'" 
		}
		
		return $true
	})]
	$Path,
	[Parameter(Position=1, Mandatory=$false, HelpMessage='The language of the subtitles to remove')]
	[ValidateScript({ 
		$exists = Test-ISOLanguageName -Name $_
	
		if(!($_ -imatch "^[a-z]{2,3}$"))
		{
			throw "Language code must be a two or three letters string." 
		}
		elseif(!$exists)
		{
			throw "The provided language code is invalid or unknown : '$_'"
		}
		
		return $true
	})]
	[string]$LanguageCode,
	[Parameter(HelpMessage='Remove subtitles in subfolders')]
	[switch]$Recurse,
	[Parameter(HelpMessage='Asks for confirmation before removing a file')]
	[switch]$Confirm,
	[Parameter(HelpMessage='Lists actions without applying them')]
	[switch]$WhatIf)
	
	BEGIN
	{
	}
	PROCESS
	{
		#Getting the root directory  
        if($Path -is [System.IO.DirectoryInfo])
        {
            $rootDirectory = $Path
        }
        else
        {
            $rootDirectory = Get-Item -LiteralPath $Path
        }
		
		#Setting parameters for Get-Subtitles function
		$parameters = @{}
		if($LanguageCode)
		{
			$parameters.Add("LanguageCode",$LanguageCode)
		}
		$parameters.Add("Recurse",$Recurse)
		
		#Search subtitles matching parameters and remove them
		$rootDirectory | Get-SubtitleFiles @parameters | Remove-Item -WhatIf:$WhatIf -Confirm:$Confirm
	}
	END
	{
	}
}

<#
.SYNOPSIS
Add or remove an ISO language code at the end of a subtitle file. 

.DESCRIPTION
This function adds or removes an ISO language code at the end of a subtitle file.
Only valid ISO codes are allowed. Use of an invalid/unknown code will raise an error. 
Only SRT and ASS files are supported. 

.PARAMETER Path
The path to the subtitle file to process. 
Pipeline values are supported, either with a string or a System.IO.FileInfo object. 

.PARAMETER LanguageCode
The ISO language code of the subtitles.
Default value is fr (French).

.PARAMETER Force
Force replacement of existing code at the end of the file name or replace existing file.

.PARAMETER Remove
Remove existing language code at the end of the file name.

.PARAMETER Confirm
Ask for confirmation before renaming a file.

.PARAMETER WhatIf
Lists actions without applying them. 

.INPUTS
System.String
System.IO.FileInfo

.OUTPUTS
System.IO.FileInfo[]

.EXAMPLE 
Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt"


Set default language for a subtitle file.

.EXAMPLE 
Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt" | Set-SubtitleFileLanguage


Set default language for the subtitle file, using the pipeline to pass the Path parameter value.

.EXAMPLE 
Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt" -LanguageCode "en"


Set english language for a subtitle file.

.EXAMPLE 
Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -Force


Force default language for a subtitle file.

.EXAMPLE 
Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -Remove


Remove english language for a subtitle file.

.EXAMPLE 
Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -Remove -Confirm

Confirm
Are you sure you want to perform this action?
Performing the operation « Rename file » on target  « Element: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt Destination: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.srt».
[Y] Yes  [A] Yes to All  [N] No  [L] No to All [S] Suspend  [?] Help (default is « Y ») : L


Remove english language for a subtitle file after confirmation.

.EXAMPLE 
Set-SubtitleFileLanguage -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt" -WhatIf
What if : Performing the operation « Rename file » on target « Element: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt Destination: C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt ».


Display changes when setting default language for a subtitle file without actually changing the file name.
#>
function Set-SubtitleFileLanguage
{
	[CmdletBinding(DefaultParameterSetName="SetLanguageCode")]
	Param(
	[Parameter(ParameterSetName="SetLanguageCode",Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The path to the subtitle file to process')]
	[Parameter(ParameterSetName="RemoveLanguageCode",Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The path to the subtitle file to process')]
	[Alias("FullName")]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Leaf))
		{
			throw "Provided path is a valid file : '$value'" 
		}
		
		return $true
	})]
	$Path,
	[Parameter(ParameterSetName="SetLanguageCode",Position=1, Mandatory=$false, HelpMessage='The ISO language code of the subtitles')]
	[ValidateScript({ 
		$exists = Test-ISOLanguageName -Name $_
	
		if(!($_ -imatch "^[a-z]{2,3}$"))
		{
			throw "Language code must be a two or three letters string." 
		}
		elseif(!$exists)
		{
			throw "The provided language code is invalid or unknown : '$_'"
		}
		
		return $true
	})]
	[string]$LanguageCode = $Script:DefaultLanguageCode,
	[Parameter(ParameterSetName="SetLanguageCode",HelpMessage='Force replacement of existing code at the end of the file name')]
	[switch]$Force,
	[Parameter(ParameterSetName="RemoveLanguageCode", Mandatory=$true, HelpMessage='Force replacement of existing code at the end of the file name or replace existing file')]
	[switch]$Remove,
	[Parameter(ParameterSetName="SetLanguageCode",HelpMessage='Ask for confirmation before renaming a file')]
	[Parameter(ParameterSetName="RemoveLanguageCode",HelpMessage='Ask for confirmation before renaming a file')]
	[switch]$Confirm,	
	[Parameter(ParameterSetName="SetLanguageCode",HelpMessage='Lists actions without applying them')]
	[Parameter(ParameterSetName="RemoveLanguageCode",HelpMessage='Lists actions without applying them')]
	[switch]$WhatIf	
	)
	
	BEGIN
	{		
	}
	PROCESS
	{
		if(!($Path -is [System.IO.FileInfo]))
		{
			$file = Get-Item -LiteralPath (Get-FileSystemFullPath -Path $Path)
		}
		else
		{
			$file = $Path
		}
		
		Write-Verbose "Processing '$($file.FullName)' file"
		$newFileName = $file.Name

		if(!$file.Extension -imatch "^\.(srt|ass)$")
		{
			Write-Warning "'$($file.Extension)' is not supported. Only SRT and ASS files can be processed."
			return
		}
		elseif($file.BaseName -imatch "^(?<BaseName>.+)\.(?<LanguageCode>[a-z]{2,3})$")
		{
			$currentCode = $matches["LanguageCode"]
			
			if(Test-ISOLanguageName -Name $currentCode)
			{
				if($PsCmdlet.ParameterSetName -eq "SetLanguageCode" -and $Force)
				{
					$newFileName = "$($matches['BaseName']).$($LanguageCode)$($file.Extension)"
				}
				elseif($PsCmdlet.ParameterSetName -eq "SetLanguageCode")
				{
					Write-Warning "File name already have a language code : '$currentCode'"
				}
				elseif($PsCmdlet.ParameterSetName -eq "RemoveLanguageCode")
				{
					$newFileName = "$($matches['BaseName'])$($file.Extension)"
				}
			}
			else
			{
				Write-Warning "The code detected in the file name is not a valid ISO code '$currentCode'."
				Write-Warning "No action on the file."
			}
		}
		else
		{
			$newFileName = "$($file.BaseName).$($LanguageCode)$($file.Extension)"
		}
		
		if($file -ne $null -and $newFileName -ne $file.Name)
		{
			#Removing any file with the same name
			if($Force)
			{
				Remove-Item -LiteralPath (Join-Path -Path $file.DirectoryName -ChildPath $newFileName) -ErrorAction SilentlyContinue
			}
			if(Test-Path (Join-Path -Path $file.DirectoryName -ChildPath $newFileName) -PathType Leaf)
			{
				Write-Warning "File '$newFileName' already exists ! Use Force option to overwite it."
			}
			else
			{
				Rename-Item -LiteralPath $file.FullName -NewName $newFileName -Force -Confirm:$Confirm -WhatIf:$WhatIf -ErrorAction SilentlyContinue
			
				Write-Verbose "New file name : '$($newFileName)'"
				if(!$WhatIf)
				{
					return Get-Item -LiteralPath (Join-Path -Path $file.DirectoryName -ChildPath $newFileName)
				}
			}
		}		
	}
	END
	{	
	}
}

<#
.SYNOPSIS
Finds video files in a directory.

.DESCRIPTION
Finds video files in a directory.
Supported files are : MPG, AVI, MP4, MKV, TS

.PARAMETER Path
The root directory containing video files.

.PARAMETER NoSubtitles
Search only videos without subtitles

.PARAMETER Subtitles
Search only videos with subtitles

.PARAMETER LanguageCode
The ISO language code of the subtitles.
Requires NoSubtitles or Subtitles option.

.PARAMETER Recurse
Search in subfolders

.INPUTS
System.IO.DirectoryInfo

.OUTPUTS
System.IO.FileInfo[]

.EXAMPLE
Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01"

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01

Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv


Search video files in a specific directory.

.EXAMPLE 
Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01" | Get-VideoFiles

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv


Search video files in a specific directory, using the pipeline to pass the Path parameter value.

.EXAMPLE 
Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\" -Recurse

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


Search all video files under a specific path.

.EXAMPLE 
Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\" -NoSubtitles -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S02


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       10/05/2017     16:19      210497922 My.Show.S02E01.HDTV.x264-FLEET.mkv


Search all video files without subtitles under a specific path.

.EXAMPLE 
Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -Subtitles -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv

Search all video files with subtitles (default language) under a specific path.

.EXAMPLE 
Get-VideoFiles -Path "C:\Users\adzero\Videos\TV Shows\My Show\" -Subtitles -LanguageCode "en" -Recurse

    Directory: C:\Users\adzero\Videos\TV Shows\My Show\S01


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39      200498902 My.Show.S01E01.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:56      187969898 My.Show.S01E02.HDTV.x264-FLEET.mkv
-a----       11/11/2016     17:58      191706742 My.Show.S01E03.HDTV.x264-FLEET.mkv


Search all video files with english subtitles under a specific path.
#>
function Get-VideoFiles
{
    [CmdletBinding(DefaultParameterSetName="All")]
    Param(
    [Parameter(ParameterSetName="All", Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The root directory containing video files')]
    [Parameter(ParameterSetName="NoSubtitles", Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The root directory containing video files')]
    [Parameter(ParameterSetName="Subtitles", Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The root directory containing video files')]
	[Alias("FullName")]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Container))
		{
			throw "Provided path is not a valid directory path: '$value'" 
		}
		
		return $true
	})]
	$Path,
	[Parameter(ParameterSetName="NoSubtitles", Mandatory=$true, HelpMessage='Search only videos without subtitles')]
	[switch]$NoSubtitles,
	[Parameter(ParameterSetName="Subtitles", Mandatory=$true, HelpMessage='Search only videos with subtitles')]
	[switch]$Subtitles,	
	[Parameter(ParameterSetName="NoSubtitles",Position=1, Mandatory=$false, HelpMessage='The ISO language code of the missing subtitles')]
	[Parameter(ParameterSetName="Subtitles",Position=1, Mandatory=$false, HelpMessage='The ISO language code of the subtitles')]
	[ValidateScript({ 
		$exists = Test-ISOLanguageName -Name $_
	
		if(!($_ -imatch "^[a-z]{2,3}$"))
		{
			throw "Language code must be a two or three letters string." 
		}
		elseif(!$exists)
		{
			throw "The provided language code is invalid or unknown : '$_'"
		}
		
		return $true
	})]
	[string]$LanguageCode,
	[Parameter(ParameterSetName="All", HelpMessage='Search in subfolders')]
	[Parameter(ParameterSetName="NoSubtitles", HelpMessage='Search in subfolders')]
	[Parameter(ParameterSetName="Subtitles", HelpMessage='Search in subfolders')]
	[switch]$Recurse)
	
	BEGIN
	{
	}
	PROCESS
	{
        if($Path -is [System.IO.DirectoryInfo])
        {
            $rootDirectory = $Path
        }
        else
        {
            $rootDirectory = Get-Item -LiteralPath $Path
        }

		$videos = $rootDirectory | Get-ChildItem -File -Recurse:$Recurse | where {$_.Extension -imatch "^\.(mpg|avi|mp4|mkv|ts)$"}
		
        $results = @()

        foreach($video in $videos)
        {
			
			$matchPattern = "$($video.BaseName)(\.[a-z]{2,3})?\.(srt|ass)$"
		
			if($LanguageCode)
			{
				$matchPattern = "$($video.BaseName)\.$LanguageCode\.(srt|ass)$"
			}

			if($Subtitles -or $NoSubtitles)
			{
				[array]$videoSubtitles = Get-ChildItem -Path $video.Directory.FullName -File | where{ $_.Name -imatch $matchPattern}
            }

            if(($videoSubtitles.Count -ge 1 -and $Subtitles) -or ($videoSubtitles.Count -eq 0 -and $NoSubtitles) -or ($PsCmdlet.ParameterSetName -eq "All"))
            {
                $results += $video
            }
        }

        return $results		
	}
	END
	{
	}
}

<#
.SYNOPSIS
Parses a video file name and extracts information for subtitles search.

.DESCRIPTION
Parses a video file name and extracts information for subtitles search.
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

.PARAMETER Path
The path of the file to parse.

.INPUTS
System.IO.FileInfo
System.String

.OUTPUTS
System.Object

.EXAMPLE 
Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv"

Path           : C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv
FileName       : My.Show.S01E01.HDTV.x264-FLEET.mkv
ShowTitle      : My Show
Season         : 1
Episode        : 1
Source         : HDTV
ReleaseGroup   : FLEET
ValidForSearch : True 

Get the episode information from a file.

.EXAMPLE 
Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" | Get-EpisodeFileInfo

Path           : C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv
FileName       : My.Show.S01E01.HDTV.x264-FLEET.mkv
ShowTitle      : My Show
Season         : 1
Episode        : 1
Source         : HDTV
ReleaseGroup   : FLEET
ValidForSearch : True 


Get the episode information from a file, using the pipeline to pass the Path parameter value.
#>
function Get-EpisodeFileInfo
{
    [CmdletBinding()]
    Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The path of the file to parse')]
	[Alias("FullName")]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Leaf))
		{
			throw "Provided path is a valid file : '$value'" 
		}
		
		return $true
	})]
	$Path)

    BEGIN
    {
    }
    PROCESS
    {
        if($Path -is [System.IO.FileInfo])
        {
            $File = $Path
        }
        else
        {
            $File = Get-Item -LiteralPath $Path
        }

        $result = New-Object -TypeName System.Object

        $result | Add-Member -MemberType NoteProperty -Name "Path" -Value $File.FullName
        $result | Add-Member -MemberType NoteProperty -Name "FileName" -Value $File.Name
    
        #Searching for Show title, Season and Episode
        if($File.BaseName -imatch "^(?<ShowTitle>.+)(\.|_|-)S?(?<Season>\d+)(E|x)(?<Episode>\d+)")
        {
            $result | Add-Member -MemberType NoteProperty -Name "ShowTitle" -Value $Matches["ShowTitle"].Replace("."," ")
            $result | Add-Member -MemberType NoteProperty -Name "Season" -Value ([int]$Matches["Season"])
            $result | Add-Member -MemberType NoteProperty -Name "Episode" -Value ([int]$Matches["Episode"])
        }

        #Searching for Resolution
        if($File.BaseName -imatch "(?<Resolution>(480|720|1080|2160)(p|i))")
        {
            $result | Add-Member -MemberType NoteProperty -Name "Resolution" -Value $Matches["Resolution"]
        }

        #Searching for Source
        if($File.BaseName -imatch "\.(?<Source>(hdtv|pdtv|web-dl|web\.dl|webrip|web|dvdrip|bdrip|bluray))\.")
        {
            $result | Add-Member -MemberType NoteProperty -Name "Source" -Value $Matches["Source"]
        }

        #Searching for ReleaseGroup
        if($File.BaseName -imatch "-(?<ReleaseGroup>([^-]+))$")
        {
            $result | Add-Member -MemberType NoteProperty -Name "ReleaseGroup" -Value $Matches["ReleaseGroup"]
        }

		#Searching for Proper or Repack tag
        if($File.BaseName -imatch "\.(?<ProperRepack>(proper|repack))\.")
        {
            $result | Add-Member -MemberType NoteProperty -Name "ProperRepack" -Value $Matches["ProperRepack"]
        }

        $result | Add-Member -MemberType ScriptProperty -Name "ValidForSearch" -Value {![String]::IsNullOrWhiteSpace($this.ShowTitle) -and $this.Season -ge 0 -and $this.Episode -gt 0}
    
        return $result
    }
    END
    {
    }
}

<#
.SYNOPSIS
Downloads subtitles matching an episode from Sous-Titres.eu website.

.DESCRIPTION
Downloads subtitles matching an episode from Sous-Titres.eu website.
Search is based on the show title, season and episode numbers, language code.
Result is a FileInfo collection of the downloaded subtitles files.

.PARAMETER ShowTitle
The name of the TV show.

.PARAMETER Season
The season number of the TV show. 

.PARAMETER Episode
The episode number(s) of the TV show episode(s).

.PARAMETER LanguageCode
The ISO code of the subtitles language. 

.PARAMETER SeriesRootUri
The root URI of the https://www.sous-titres.eu TV shows section. 

.INPUTS
System.String
System.Int32

.OUTPUTS
System.IO.FileInfo[]

.EXAMPLE 
Import-EpisodeSubtitles -ShowTitle "My Show" -Season 1 -Episode 1

Download all the subtitles in the default language for the episode of the show.

.EXAMPLE 
Import-EpisodeSubtitles -ShowTitle "My Show" -Season 1 -Episode 1 -LanguageCode "en"

    Directory: C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt


Download all english subtitles for an episode.

.EXAMPLE 
Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\My.Show.S01E01.HDTV.x264-RG.mp4" | Import-EpisodeSubtitles -LanguageCode "en"

    Directory: C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.1x01.FR.FLEET


Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----       07/09/2016     06:39            902 My.Show.S01E01.HDTV.x264-FLEET.en.srt


Download all english subtitles for an episode, using the pipeline to pass the episode information parameters values.
#>
function Import-EpisodeSubtitles
{
    [CmdletBinding()]
    Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='The name of the TV show')]
	[string]$ShowTitle,
	[Parameter(Position=1, Mandatory=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='The season number of the TV show')]
	[int]$Season,
	[Parameter(Position=2, Mandatory=$true, ValueFromPipelineByPropertyName=$true, HelpMessage='The episode number(s) of the TV show episode(s)')]
	[int]$Episode,
	[Parameter(Position=3, ValueFromPipelineByPropertyName=$true, HelpMessage='The ISO code of the subtitles language')]
	[string]$LanguageCode=$Script:DefaultLanguageCode,
    [Parameter(Position=4, HelpMessage='The root URI of the https://www.sous-titres.eu TV shows section')]
    [string]$SeriesRootUri = $Script:SeriesRootUri)

	BEGIN
	{
		if(!$Global:SubtitleToolboxDownloads)
		{
			$Global:SubtitleToolboxDownloads = @{}
		}
	}
	PROCESS
	{
		$uri = "$($SeriesRootUri)/$($ShowTitle.ToLower().Replace(" ","_")).html"
		$response = Invoke-WebRequest -Uri $uri

		if($response.StatusCode -ne 200)
		{
			Write-Error "The URI '$uri' doesn't exist."
		}
		else
		{
			#Searching links matching episode only
			[array]$links = $response.Links | where {$_.href -ne $null -and $_.href -imatch "$($ShowTitle -ireplace " ",".")\..*$($Season)x$($Episode.ToString().PadLeft(2,'0'))\.(.*\.)?$LanguageCode.*\.zip"} 
			#Searching links matching whole season
			[array]$links += $response.Links | where {$_.href -ne $null -and $_.href -imatch "$($ShowTitle -ireplace " ",".")\.S0?$($Season)\.(.*\.)?$LanguageCode.*\.zip"} 
			
			if($links -eq $null -or $links.Count -eq 0)
			{
				Write-Verbose "No matching subtitles download found !"        
			}
			else
			{
				if($links.Count -gt 1)
				{
					Write-Verbose "Several matching subtitles downloads found !"
				}

				#Creating root zip expand path
				$expandPath = Join-Path -Path $Script:TempDirectory -ChildPath $([Guid]::NewGuid().ToString())

				#Processing found links
				foreach($link in $links)
				{
					#Checking if file has already been downloaded
					$zipUri = [System.Uri]"$($SeriesRootUri)/$($link.href)"
					if($Global:SubtitleToolboxDownloads.ContainsKey($zipUri) -and (Test-Path -LiteralPath $Global:SubtitleToolboxDownloads[$zipUri]))
					{
						Write-Verbose "'$zipUri' file already downloaded..."
						$zipPath = $Global:SubtitleToolboxDownloads[$zipUri]
					}
					else
					{
						Write-Verbose "Downloading '$zipUri' file"
						
						$zipPath = Join-Path -Path $Script:TempDirectory -ChildPath ([Guid]::NewGuid().ToString() + "_" +(([System.Uri]"$SeriesRootUri/$($link.href)").Segments | select -Last 1))
						Invoke-WebRequest -Uri $zipUri -OutFile $zipPath -ErrorAction Stop
						
						if(!$Global:SubtitleToolboxDownloads.ContainsKey($zipUri))
						{
							$Global:SubtitleToolboxDownloads.Add($zipUri,$zipPath)
						}
						else
						{
							$Global:SubtitleToolboxDownloads[$zipUri] = $zipPath
						}
					}
					
					#Expanding archive and delete it
					$zipExpandPath = Join-Path -Path $expandPath -ChildPath (([System.Uri]"$SeriesRootUri/$($link.href)").Segments | select -Last 1).Replace(".zip","")
					$zipFile = Get-Item -LiteralPath $zipPath
					Expand-Archive -LiteralPath $zipPath -DestinationPath $zipExpandPath -Force
				}

				return Get-ChildItem -LiteralPath $expandPath -File -Recurse
			}
		}
	}
	END
	{
	}
}

<#
.SYNOPSIS
Gets the matching information and score of a subtitle file.

.DESCRIPTION
Gets the matching information and score of a subtitle file.
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

.PARAMETER Path
The path of the file to analyse.

.PARAMETER EpisodeInformation
Episode file information provided by the Get-EpisodeFileInfo function.

.PARAMETER LanguageCode
The desired language code of the subtitles.

.INPUTS
System.IO.FileInfo
System.String
System.Object

.OUTPUTS
System.Object

.EXAMPLE 
Get-SubtitleFileScore -Path "C:\Users\adzero\AppData\Local\Temp\My.Show.S01E01.HDTV.x264-RG.srt" -EpisodeInformation (Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\My.Show.S01E01.HDTV.x264-RG.mp4") -LanguageCode "en"


Get a subtitle file score for a video file and english language.

.EXAMPLE 
Get-Item -Path "C:\Users\adzero\AppData\Local\Temp\My.Show.S01E01.HDTV.x264-RG.srt" | Get-SubtitleFileScore -EpisodeInformation (Get-EpisodeFileInfo -Path "C:\Users\adzero\Videos\TV Shows\My Show\My.Show.S01E01.HDTV.x264-RG.mp4") -LanguageCode "en"


Get a subtitle file score for a video file and english language, using the pipeline to pass the Path parameter value.
#>
function Get-SubtitleFileScore
{
    [CmdletBinding()]
    Param(
    [Parameter(Position=0,Mandatory=$true,ValueFromPipelineByPropertyName=$true,ValueFromPipeline=$true,HelpMessage='The path of the file to analyse')]
	[Alias("FullName")]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Leaf))
		{
			throw "Provided path is a valid file : '$value'" 
		}
		
		return $true
	})]
    $Path,
    [Parameter(Position=1,Mandatory=$true,HelpMessage='Episode file information provided by the Get-EpisodeFileInfo function')]
    $EpisodeInformation,
    [Parameter(Position=2,Mandatory=$false,HelpMessage='The desired language code of the subtitles')]
	[ValidateScript({ 
		$exists = Test-ISOLanguageName -Name $_
	
		if(!($_ -imatch "^[a-z]{2,3}$"))
		{
			throw "Language code must be a two or three letters string." 
		}
		elseif(!$exists)
		{
			throw "The provided language code is invalid or unknown : '$_'"
		}
		
		return $true
	})]
	[string]$LanguageCode)

    BEGIN
    {
    }
    PROCESS
    {
        if($Path -is [System.IO.FileInfo])
        {
            $file = $Path
        }
        else
        {
            $file = Get-Item -LiteralPath $Path
        }

		$result = New-Object -TypeName System.Object

        $result | Add-Member -MemberType NoteProperty -Name "Path" -Value $file.FullName
        $result | Add-Member -MemberType NoteProperty -Name "FileName" -Value $file.Name
        $result | Add-Member -MemberType NoteProperty -Name "MatchShowTitle" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchSeason" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchEpisode" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchResolution" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchSource" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchReleaseGroup" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchLanguage" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchTag" -Value $false
        $result | Add-Member -MemberType NoteProperty -Name "MatchProperRepack" -Value $false

        $result | Add-Member -MemberType ScriptProperty -Name "Score" -Value {([int]$this.MatchShowTitle)*40+([int]$this.MatchSeason)*20+([int]$this.MatchEpisode)*22+([int]$this.MatchResolution)*1.0+([int]$this.MatchSource)*5+([int]$this.MatchProperRepack)*1.5+([int]$this.MatchReleaseGroup)*10+([int]$this.MatchTag)*0.5}

        if($file.Extension -imatch "^\.(?<Extension>[^\.]+)$")
        {
            $result | Add-Member -MemberType NoteProperty -Name "SubtitleType" -Value $Matches["Extension"]
        }

        if($EpisodeInformation -ne $null)
        {
            if($EpisodeInformation.ShowTitle -and $file.Name -imatch "$($EpisodeInformation.ShowTitle -ireplace " ",".")")
            {
                $result.MatchShowTitle = $true
            }
            if(($EpisodeInformation.Season -and $file.Name -imatch "S?0*$($EpisodeInformation.Season)(E|x)") `
				-or ($EpisodeInformation.Season -and $file.Name -imatch "\.$($EpisodeInformation.Season)\d{2,}\."))
            {
                $result.MatchSeason = $true
            }
            if(($EpisodeInformation.Episode -and $file.Name -imatch "(E|x)$($EpisodeInformation.Episode.ToString().PadLeft(2,'0'))") `
				-or ($EpisodeInformation.Season -and $EpisodeInformation.Episode -and $file.Name -imatch "\.$($EpisodeInformation.Season)$($EpisodeInformation.Episode.ToString().PadLeft(2,'0'))\."))
            {
                $result.MatchEpisode = $true
            }
            if(($EpisodeInformation.Resolution -and $file.Name -imatch "$($EpisodeInformation.Resolution)") `
				-or ($EpisodeInformation.Resolution -and $file.DirectoryName.Replace($Script:TempDirectory,"") -imatch "$($EpisodeInformation.Resolution)"))
            {
                $result.MatchResolution = $true
            }
            if(($EpisodeInformation.Source -and $file.Name -imatch "$($EpisodeInformation.Source)") `
				-or ($EpisodeInformation.Source -and $file.DirectoryName.Replace($Script:TempDirectory,"") -imatch "$($EpisodeInformation.Source)"))
            {
                $result.MatchSource = $true
            }
            if(($EpisodeInformation.ReleaseGroup -and $file.Name -imatch "$($EpisodeInformation.ReleaseGroup)") `
				-or ($EpisodeInformation.ReleaseGroup -and $file.DirectoryName.Replace($Script:TempDirectory,"") -imatch "$($EpisodeInformation.ReleaseGroup)"))
            {
                $result.MatchReleaseGroup = $true
            }
            if($EpisodeInformation.ProperRepack -and $file.Name -imatch "\.$($EpisodeInformation.ProperRepack)\.")
            {
                $result.MatchProperRepack = $true
            }
            if($LanguageCode -and (($file.Name -imatch "\.$($LanguageCode)\.") `
				-or ($file.DirectoryName.Replace($Script:TempDirectory,"") -imatch "\.$($LanguageCode)(\.|\\|$)")))
            {
                $result.MatchLanguage = $true
            }
            if($File.Name -imatch "\.TAG\.")
            {
                $result.MatchTag = $true
            }
        }

        return $result
    }
    END
    {
    }
}

<#
.SYNOPSIS
Finds the best subtitles in a given language for an episode.

.DESCRIPTION
Finds the best subtitles in a given language for an episode.

.PARAMETER Path
The path of the episode video file.

.PARAMETER LanguageCode
The ISO code of the subtitles language. 

.PARAMETER Force
Allow to overwrite existing subtitles. 

.PARAMETER WhatIf
Lists actions without applying them. 

.INPUTS
System.IO.FileInfo
System.String

.OUTPUTS
System.IO.FileInfo

.EXAMPLE 
$subtitles = Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv"
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.FR.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'


Get the best subtitles in default language for a video file.

.EXAMPLE 
$subtitles = Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" | Find-EpisodeSubtitles 
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.FR.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'


Get the best subtitles in default language for a video file, using the pipeline to pass the Path parameter value.

.EXAMPLE 
$subtitles = Get-Item "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" | Find-EpisodeSubtitles 
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.FR.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'


Get the best subtitle file in default language for the video file, using the pipeline to pass the Path parameter value.

.EXAMPLE 
$subtitles = Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" -LanguageCode "en"
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.EN.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.EN.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.en.srt'


Get the best subtitle file in english for the video file.

.EXAMPLE 
Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" -Force | Out-Null
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.EN.srt'
Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'


Get the best subtitles in default language for a video file and force overwrite of existing file. 

.EXAMPLE 
Find-EpisodeSubtitles -Path "C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.mkv" -WhatIf | Out-Null
Searching subtitles for : 'My.Show.S01E01.HDTV.x264-FLEET.mkv' ...
Best subtitle found : 'My.Show.S01E01.HDTV.x264-FLEET.FR.mkv'
Path : 'C:\Users\adzero\AppData\Local\Temp\d79c8ce8-ac5b-4868-8995-9ada7ae7b55b\My.Show.S01E01.HDTV.x264-FLEET.EN.srt'
WhatIf : Saved as : 'C:\Users\adzero\Videos\TV Shows\My Show\S01\My.Show.S01E01.HDTV.x264-FLEET.fr.srt'


Get the best subtitles in default language for a video file and displays its destination name without actually creating the file.  
#>
function Find-EpisodeSubtitles
{
    [CmdletBinding()]
    Param(
    [Parameter(Position=0, Mandatory=$true, ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, HelpMessage='The path of the episode video file')]
    [Alias("FullName")]
	[ValidateScript({ 
		$value = Get-FileSystemFullPath -Path $_
	
		if(!(Test-Path -LiteralPath $value -PathType Leaf))
		{
			throw "Provided path is a valid file : '$value'" 
		}
		
		return $true
	})]
    $Path,
    [Parameter(Position=1, Mandatory=$false, HelpMessage='The ISO code of the subtitles language')]
	[ValidateScript({ 
		$exists = Test-ISOLanguageName -Name $_
	
		if(!($_ -imatch "^[a-z]{2,3}$"))
		{
			throw "Language code must be a two or three letters string." 
		}
		elseif(!$exists)
		{
			throw "The provided language code is invalid or unknown : '$_'"
		}
		
		return $true
	})]
    [string]$LanguageCode=$Script:DefaultLanguageCode,
    [Parameter(Position=2,Mandatory=$false,HelpMessage='Allow to overwrite existing subtitles')]
    [switch]$Force,
	[Parameter(HelpMessage='Lists actions without applying them')]
	[switch]$WhatIf)

    BEGIN
    {
    }
    PROCESS
    {
        if($Path -is [System.IO.FileInfo])
        {
            $videoFile = $Path
        }
        else
        {
            $videoFile = Get-Item -LiteralPath $Path
        }

		Write-Host "Searching subtitles for : '$($videoFile.Name)' ..."
		
        $episodeInfo = $videoFile | Get-EpisodeFileInfo

        if($episodeInfo -ne $null -and $episodeInfo.ValidForSearch)
        {
			$subtitles = $episodeInfo | Import-EpisodeSubtitles | Get-SubtitleFileScore -EpisodeInformation $episodeInfo -LanguageCode $LanguageCode | Sort Score -Descending
			
			if($Verbose)
			{
				Write-Verbose "Subtitles scores : `n"
				foreach($subtitle in $subtitles)
				{
					Write-Verbose "$($subtitle.FileName)"
					Write-Verbose "Score : $($subtitle.Score)"
				}
			}
			
            $bestSubtitle = $subtitles | Select-Object -First 1
            
            if($bestSubtitle)
            {
                $subtitleFile = Get-Item -LiteralPath $bestSubtitle.Path
				Write-Host "Best subtitle found : '$($subtitleFile.Name)'" -ForegroundColor Green
				Write-Host "Path : '$($subtitleFile.DirectoryName)'"
                $destinationPath = Join-Path -Path $videoFile.Directory.FullName -ChildPath "$($videoFile.BaseName)$($subtitleFile.Extension)"
				
				if($WhatIf)
				{
					Write-Host "WhatIf : Saved as : '$(Join-Path -Path $videoFile.Directory.FullName -ChildPath $videoFile.BaseName).$LanguageCode$($subtitleFile.Extension)'"					
				}
				else
				{
					Move-Item -LiteralPath $subtitleFile.FullName -Destination $destinationPath -Force:$Force -ErrorAction SilentlyContinue 
					$destinationFile = $destinationPath | Set-SubtitleFileLanguage -LanguageCode $LanguageCode -Force:$Force -WhatIf:$WhatIf
					Write-Host "Saved as : '$($destinationFile.FullName)'"
				}
				
				#Remove temporary folders
				$tempDirectory = $subtitleFile.Directory
				while($tempDirectory.Parent.FullName -ne (Get-Item $Script:TempDirectory).FullName)
				{	
					$tempDirectory = $tempDirectory.Parent
				}
				Remove-Item -LiteralPath $tempDirectory.FullName -Force -Recurse -ErrorAction SilentlyContinue
				
				Write-Host 
                return $destinationFile
            }
            else
            {
                Write-Warning "No subtitles found !"
            }
        }
    }
    END
    {
    }
}

<#
.SYNOPSIS
Resets cache and deletes downloaded files. 

.EXAMPLE 
Clear-DownloadCache
#>
function Clear-DownloadCache
{
	foreach($download in $Global:SubtitleToolboxDownloads.Keys)
	{
		Remove-Item -LiteralPath $Global:SubtitleToolboxDownloads[$download] -Force -ErrorAction SilentlyContinue
	}
	$Global:SubtitleToolboxDownloads = @{}
}

<#
.SYNOPSIS
Gets all available cultures.

.OUTPUTS
System.Globalization.CultureInfo[]

.EXAMPLE 
Get-Cultures

Get all available cultures.
#>
function Get-Cultures
{
	return [System.Globalization.CultureInfo]::GetCultures([System.Globalization.CultureTypes]::AllCultures)
}

<#
.SYNOPSIS
Checks the validity of an ISO language code. 

.INPUTS
System.String

.PARAMETER Name
The language code to check.

.INPUTS
System.String

.OUTPUTS
System.Boolean

.EXAMPLE 
Test-ISOLanguageName -Name "en"

Test the "en" language code.
#>
function Test-ISOLanguageName
{
    [CmdletBinding()] 
	Param(
    [Parameter(Position=0, Mandatory=$true, HelpMessage='The language code to check')]	
	$Name)
	
	[array]$culture = Get-Cultures | where {$_.TwoLetterISOLanguageName -eq $Name -or $_.ThreeLetterISOLanguageName -eq $Name}
	
	return ($culture -ne $null -and $culture.Count -gt 0)	
}

<#
.SYNOPSIS
Gets the full path of a filesystem item.

.DESCRIPTION
Gets the full path of a filesystem item.
Relative paths are appended to $PSScriptRoot value.

.PARAMETER Path
Path to the file system item.

.INPUTS
System.String
System.IO.FileSystemInfo

.OUTPUTS
System.String

.EXAMPLE 
Get-FileSystemFullPath -Path "C:\Users\adzero\Sample.txt"

Get the full path of the file.

.EXAMPLE 
Get-FileSystemFullPath -Path "Sample.txt"

Get the full path of the file, relative to $PSScriptRoot directory.
#>
function Get-FileSystemFullPath
{
    [CmdletBinding()] 
    Param(
    [Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, HelpMessage='Path to the file system item')]	
	$Path)

    if($Path -is [System.IO.FileSystemInfo])
    {
        return $Path.FullName
    }
    elseif([System.IO.Path]::IsPathRooted($Path))
    {
        return [string]$Path
    }
	else
	{
		return (Join-Path -Path $PSScriptRoot -ChildPath [string]$Path)
	}
}

# SIG # Begin signature block
# MIIIqQYJKoZIhvcNAQcCoIIImjCCCJYCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUXaPcjpecnbneET1DqA7iOaGI
# f56gggUwMIIFLDCCAxSgAwIBAgIQTf/o8FZ5EbFG4TtSC8EkYDANBgkqhkiG9w0B
# AQ0FADAuMSwwKgYDVQQDDCNBZFplcm8gUG93ZXJTaGVsbCBMb2NhbCBDZXJ0aWZp
# Y2F0ZTAeFw0xNjEwMTQxNzA5MjFaFw0xNzEwMTQxNzE4MzhaMC4xLDAqBgNVBAMM
# I0FkWmVybyBQb3dlclNoZWxsIExvY2FsIENlcnRpZmljYXRlMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEA5iKQbshRfeOjuEIYDx8WidA/iprQcOq4EDRF
# /tmiyGgz5RZEqzMA6hN5ASmyuyP+xPyd96l4RIp7wv7oER5bEXsmtL4tiGwwPnVj
# xIlGY6hvhvtORb1wMnq6fpDn1/zLORh+qpGNL179FEBjU4jQYpJNyaNBM//v5iYL
# Hhsb+knFvhz+4wUb9lrvqAFmS6pTqgeSEFlHsUX56+syPBNMc9HMN+50UWJFbS4h
# tmrVd2nqVIk4FDS47X0v+D4M1Yc1Sk7IM5A1Bwd7qMttsQuWGC5sSWOX1Cp25N3n
# 8iFz+fzsrV36nwMh4eogdw687BfEthsoMRD4GAOI+3u7nXB8IXjtW1JdsUNtHRno
# U0Hos2VMWvt3YeBPZu0Pr2OR3Nq4CeeUtWtNGbXACL2pbaCpYaagR5ekxwqBDo5s
# H/gIcY9+sT3jIIz7rh5IUw+HscIgtgialLeafBCa9AVebV4c7w1nioOceMH81FD6
# kLf6InemwTLdNgg/EUDfCdnfTStnlnEyiEY4MwpXGO58HMS9ziValx9u+Q2rjNqW
# uwcPX0OpMaMufzZB4eWgltFZ9I0KaT9wRMbtnx6u6k34aB58ErRpm8T/Si/PVNuh
# 3K14Tb4qUQFTOnHRk+VYbtqog49m0Ie8OfOy4LIP3MByXh4MIOcMrAmY+ZeFUJQT
# 4BlQ31UCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUF
# BwMDMB0GA1UdDgQWBBQJT61K0yqM0R8742VXS/ir+x2pOjANBgkqhkiG9w0BAQ0F
# AAOCAgEA1f/btoE8rV/jBAaXzN0cIa920+rBYvqPclVThoIYlxZJf7XJKe1Glk09
# vdECUUJLoLNfHAKsoGBLULcx25kufDBFuNNb7a2hJrJhO6QhHh4y37LyPZhhn3CB
# 13qYSEraKS9EXFtZx5khEaE7iMpXkLLhBo340J+4Ted0JQvKPlfxrJ10/aBcPhx4
# Oy2j17LzDbIN6JDc1yKXq073Pg+TtpSV6yHEmgMiAaEFyypm9a3geo6P/EvYKK7s
# hHSqEeg4ggh5CsbFMscEO+eANfcyCZxztzF18vbQrEn/FJoBnvJXmWvtMCkNaXCE
# y5opQedVYS1uqzlz1xHFmStQ+OlbtCnyeXEwJt2PYp/9J8iPVzRExYOXT7IJikry
# 0hxaEwk8PH7wqCdnXfoiUJtIqzq0Pu8aFxa6lWGGKo2HmswvN6DxXN486KnVYEdF
# lZ6Ra+1In8g+qkCqR0zzlBQWre56MMHx7kR2zt2kE5hIm7fouzLsJ+jR/trrvyc8
# k9o1oSuuTu7U+11BNKsacTpUY9C8q1nMEvYr1oI02L3dhbI4irg7lHtIQeNpxXR5
# ebM/laNLgu/T//+h7oTTGlaFBK1RHrl6MWKn3iqIWqrzQ3+eKzH4OuA+xNjKCG61
# jF8xXG0qXDnkWeCzCEma5/rRO8rdGEyQtKnDBa6RogceS5Yj3JwxggLjMIIC3wIB
# ATBCMC4xLDAqBgNVBAMMI0FkWmVybyBQb3dlclNoZWxsIExvY2FsIENlcnRpZmlj
# YXRlAhBN/+jwVnkRsUbhO1ILwSRgMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEM
# MQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQB
# gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBTXESaYOLnBaFQE
# InNdbB+MFvXDrTANBgkqhkiG9w0BAQEFAASCAgAEyPaqOQUQSHfZtp2bWHRaeyJq
# EvI2pZ7UrB3hLxVt1F3t6AZQ5lwWbI+eXs+PHbSvOYP6gzd4AQ/zh7qQeXBTNXnP
# gIyvurLcnnE8JXOSGU58MOa5Xg2oXVTFARv7GC4UYgSwqckFeDNvVWDvb9UVDJjF
# +G4ffzcgUXjG4p4g/NM/iRLo1GkaGeG596bSG7o5/FjBgAwxOgwKXOTsABocakWI
# tK00D46t3I/gInQSglm6AYmWRiWzkUa2KVJ0qYLtbsPQozIvrXv5txyqK0Wgg5vP
# wWr0Oz5QAJTPtfBjlrvW9xaON0r1S1hD/g5/cWfIeau5wZoy4kGu4iXbJMs+Jl/P
# 2E7nDwx/0Bypgrz0Z6Xry6OWSy5lzqgsGhLQM48Rds/qQLLgn6zqgMde5g2514SF
# AKGhfi+6h8g3xwitPDbwlKCek2pD962g72F3STegykjnt02pQBT9savnsFGbgJwO
# 8tOi9qCQOcePEfwCIG0iihGq5LJNsXCFzh84Jg9MwScCsYwgmnZB3GmMiUwNE6em
# /HjXTT/63/mZSfi/nGk5oiPntIoGI7jnc5U/uGnx7V3iHG+63tyn4RWeZ+kMYRad
# revJOpEVPwgY0nDoS3mamaQbUvuYFVjIABAxU2+4m5zlaeyZMNKSgg8SIZIhcnGC
# rR+GKBu2eS9yzw06KA==
# SIG # End signature block
