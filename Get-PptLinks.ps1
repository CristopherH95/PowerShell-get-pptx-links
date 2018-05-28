function Get-PptLinks {
    <#
    .SYNOPSIS
    This function unpacks a pptx file and parses all links mentioned in it to a file
    .DESCRIPTION
    Extracts a pptx file to a temp directory and looks for the _rels path inside, where
    it can then parse the xml files containing links (will only grab external links)
    .EXAMPLE
    Get-PptLinks -Path C:\myPowerPoint.pptx -Destination mylinks.txt
    .PARAMETER Path
    The path to the pptx file to process
    .PARAMETER Destination
    The file to save output to (optional)
    .PARAMETER html
    Switch to output links to a formatted html file instead of a text file
    #>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [System.IO.FileInfo]$Path,
        [string]$Destination = "links.txt",
        [switch]$html
    )
    $ErrorActionPreference = "Stop"

    $currentPath = Convert-Path .
    $tempExtract = $currentPath + '\tmpPptExtract' #unpack ppt file to temp directory
    Write-Verbose -Message "Unpacking pptx file $($Path) to temp path: $($tempExtract)"
    Add-Type -assembly "system.io.compression.filesystem"
    [io.compression.zipfile]::ExtractToDirectory($Path, $tempExtract)
    $linksDir = $tempExtract + '\ppt\slides\_rels'  #construct path to ppt relationships
    $links = New-Object System.Collections.ArrayList
    Write-Verbose -Message "Checking slides xml data"
    foreach($file in Get-ChildItem -Path $linksDir) {
        [xml]$linksXml = Get-Content -Path $file.FullName   #get xml of each relationships file
        $linksInXmlObjs = $linksXml.Relationships.Relationship | Where-Object { $_.TargetMode -eq "External" }
        foreach($xmlObj in $linksInXmlObjs) {   #get link from each relationship in xml file
            Write-Verbose -Message "Found link: $($xmlObj.Target)"
            $links.Add($xmlObj.Target) > $null 
        }
    }
    [System.GC]::Collect()
    Write-Verbose -Message "Filtering out any duplicate links"
    $links = $links | Select-Object -Unique #only keep unique links
    $linksFileString = ""
    Write-Verbose -Message "Creating output file"
    if ($html) {
        $Destination = $Destination.Substring(0, $Destination.IndexOf('.'))
        $Destination = $Destination + ".html"
        $linksFileString = $linksFileString + "<html><body><table><caption>Links found in $($Path)</caption>"
        $linksFileString = $linksFileString + "<tr><th>Link</th></tr>"
        foreach($lnk in $links) {   #write output file content in html
            $linksFileString = $linksFileString + "<tr><td>"
            $linksFileString = $linksFileString + "<a href='" + $lnk + "'>$($lnk)</a></td></tr>"
        }
        $linksFileString = $linksFileString + "</table></body></html>"
    }
    else {
        foreach($lnk in $links) {   #write output file content plain text
            $linksFileString = $linksFileString + "- " + $lnk + "`r`n"
        }
    }
    try {
        $writer = New-Object System.IO.StreamWriter($Destination)
        $writer.Write($linksFileString)
        $writer.Close()
    }
    catch {
        Write-Output "Unexpected exception encountered while trying to write to: $($Destination)"
    }
    finally {
        if ($writer -ne $null) {
            $writer.Dispose()
        }
    }
    Write-Verbose -Message "Cleaning up temp files at $($tempExtract)"
    Remove-Item -Recurse -Force $tempExtract    #clean up
    Write-Verbose -Message "Output saved to: $($Destination)"
    [System.GC]::Collect()
}