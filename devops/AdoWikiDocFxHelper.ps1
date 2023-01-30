import-module Powershell-Yaml

$baseUrl = "http://home.local"
$baseUri = [Uri]::new($baseUrl)

function script:Get-AdoWikiFolders
{
  param($Path, [string[]]$Exclude)

  $workingDirectory = (Get-Location).Path
  $folders = [System.Collections.ArrayList]::new()

  $folders.Add((Get-Item $Path).FullName) | Out-null

  $subFolders = Get-ChildItem -path $Path -Recurse -Directory

  foreach($subFolder in $subFolders)
  {
    <#
      $subFolder = $subFolders | select-object -first 1
    #>
    $relative = $subFolder.FullName.Substring($workingDirectory.Length)
    
    $segments = $relative.Split("$([IO.Path]::DirectorySeparatorChar)", [System.StringSplitOptions]::RemoveEmptyEntries)

    if (!$segments.Where({$_ -in $Exclude}))
    {
      $folders.Add($subFolder.FullName) | out-null
    }
  }
  
  return $folders
}

function script:Get-WikiMarkdowns
{
  param($Folder)

  return Get-ChildItem -path $Folder -File -Filter "*.md"

}

function script:Set-MdYamlHeader
{
  param($file, $key, $value)
  
  $content = get-content -path $file.Fullname

  $yamlHeaderMarkers = $content | select-string -pattern '^---\s*$'

  $data = [ordered]@{}
  $conceptual = $content

  if ($yamlHeaderMarkers.count -ge 2 -and $yamlHeaderMarkers[0].LineNumber -eq 1)
  {
    $yaml = $content[1 .. ($yamlHeaderMarkers[1].LineNumber - 2)]
    $data = ConvertFrom-Yaml -Yaml ($yaml -join "`n") -Ordered
    $conceptual = $content | select-object -skip $yamlHeaderMarkers[1].LineNumber
  }

  $data[$key] = $value

  $content = "---`n$(ConvertTo-Yaml -Data $data  )---`n$($conceptual -join "`n")"

  $content | set-content -path $file.Fullname

}

function script:Get-DocfxItemMetadata
{
  param($mdFile)

  $workingDirectory = (Get-Location)

  $item = [ordered]@{
    AdoWiki = [ordered]@{
      File            = $mdFile                 # A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-).md [FileInfo]
      FileName        = $mdFile.Name            # A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-).md
      FileAbsolute    = $mdFile.FullName        # c:\x\y\A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-).md
      FileRelative    = $null                   # .\A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-).md
      FileRelativeUri = $null                   # ./A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-).md
      LinkOrderItem   = $null                   # A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-)
      LinkRelative    = $null                   # ./A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-)
      LinkAbsolute    = $null                   # /A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-)
      LinkMarkdown    = $null                   # /A-%2D-b%2Dc\(d\)-\(e\)-%2D-\(f\)-%2D-\(-h-\)
      LinkDisplay     = $null                   # A - b-c(d) (e) - (f) - ( h )
      FolderName      = $null                   # A-%2D-b%2Dc(d)-(e)-%2D-(f)-%2D-(-h-)
      Folder          = $null
      WikiPath        = $null                   # /A - b-c(d) (e) - (f) - ( h )
    }
    DocFxSafe = [ordered]@{
      File            = $null
      FileName        = $null
      FileAbsolute    = $null
      FileRelative    = $null
      FileRelativeUri = $null
      LinkRelative    = $null
      LinkAbsolute    = $null
      LinkMarkdown    = $null
      LinkDisplay     = $null
      FolderName      = $null
      RenameRequired  = $false
      FileIsRenamed   = $false
    }
  }
  $item.AdoWiki.FileRelative    = ".$($item.AdoWiki.FileAbsolute.Substring($workingDirectory.Path.Length))"
  $item.AdoWiki.FileRelativeUri = ".$($item.AdoWiki.FileAbsolute.Substring($workingDirectory.Path.Length))".Replace("$([IO.Path]::DirectorySeparatorChar)", "/")
  $item.AdoWiki.LinkOrderItem   = $item.AdoWiki.FileName.Replace(".md", "")
  $item.AdoWiki.LinkRelative    = $item.AdoWiki.FileRelativeUri.Replace(".md", "")
  $item.AdoWiki.LinkAbsolute    = $item.AdoWiki.LinkRelative.Substring(1)
  $item.AdoWiki.LinkMarkdown    = $item.AdoWiki.LinkAbsolute.Replace("\(", "(").Replace("\)", ")")
  $item.AdoWiki.LinkDisplay     = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.LinkOrderItem.Replace("\(", "(").Replace("\)", ")").Replace("-", " "))
  $item.AdoWiki.Folder          = (Get-ChildItem -Path $mdFile.Directory -Directory | where-object {$_.Name -eq $item.AdoWiki.LinkOrderItem})
  if ($item.AdoWiki.Folder)
  {
    $item.AdoWiki.FolderName    = $item.AdoWiki.Folder.Name
  }
  $item.AdoWiki.WikiPath        = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.LinkAbsolute.Replace("-", " "))

  $item.DocFxSafe.File            = $item.AdoWiki.File 
  $item.DocFxSafe.FileName        = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.FileName)
  $item.DocFxSafe.FileAbsolute    = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.FileAbsolute)
  $item.DocFxSafe.FileRelative    = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.FileRelative)
  $item.DocFxSafe.FileRelativeUri = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.FileRelativeUri)
  $item.DocFxSafe.LinkRelative    = "$([System.Web.HttpUtility]::UrlDecode($item.AdoWiki.LinkRelative)).md"
  $item.DocFxSafe.LinkAbsolute    = "$([System.Web.HttpUtility]::UrlDecode($item.AdoWiki.LinkAbsolute)).md"
  $item.DocFxSafe.LinkMarkdown    = "$([System.Web.HttpUtility]::UrlDecode($item.AdoWiki.LinkMarkdown)).md"
  $item.DocFxSafe.LinkDisplay     = $item.AdoWiki.LinkDisplay
  $item.DocFxSafe.FolderName      = [System.Web.HttpUtility]::UrlDecode($item.AdoWiki.FolderName)
  
  $item.DocFxSafe.RenameRequired = $item.DocFxSafe.FileName -ne $item.AdoWiki.FileName

  return [PSCustomObject]$item
}

function script:Add-AdoWikiApi
{
  param($Path, $ApisSubfolder, $ApiName, $ApiSubfolder, $ApiUid)

  $private:folder_Root_Path = $Path.FullName
  $private:folder_Apis_Path = Join-Path $folder_Root_Path -ChildPath $ApisSubfolder
  $private:folder_Api_Path  = Join-Path $folder_Apis_Path -ChildPath $ApiSubfolder

  $private:file_Apis_Toc_yaml = Join-Path $folder_Apis_Path -ChildPath "toc.yml"



  if (Test-Path $folder_Apis_Path)
  {
    Write-Host "  Folder $(Resolve-path -Path $folder_Apis_Path -Relative) exists"
  }
  else
  {
    Write-Host "  Folder $($folder_Apis_Path) does not exist, creating"
    New-Item -Path $folder_Apis_Path -ItemType Directory
  }
  


  if (Test-Path $folder_Api_Path)
  {
    Write-Host "  Folder $(Resolve-path -Path $folder_Api_Path -Relative) exists"
  }
  else
  {
    Write-Host "  Folder $($folder_Api_Path) does not exist, creating"
    New-Item -Path $folder_Api_Path -ItemType Directory
  }
  


  if (Test-Path $file_Apis_Toc_yaml)
  {
    Write-Host "  Folder $(Resolve-path -Path $file_Apis_Toc_yaml -Relative) exists, loading it"
    $private:Apis_Toc = get-content $file_Apis_Toc_yaml | ConvertFrom-Yaml -Ordered
  }
  else
  {
    $private:Apis_Toc = [ordered]@{
      items = @()
    }
  }

  $Apis_Toc.items += [ordered]@{
    name = $ApiName
    href = "$($ApiSubfolder)/toc.yml"
    topicUid = $ApiUid
  }

  ConvertTo-Yaml -data $Apis_Toc -OutFile $file_Apis_Toc_yaml -Force

}



function script:Set-SafeWikiFileName
{
  param($mdFileMetadata)

  <#
    $mdFile = $mdFiles | select-object -first 1
    $mdFile = $mdFiles | select-object -first 1 -skip 1
    $mdFile = $mdFiles | select-object -first 1 -skip 2
    $mdFile
  #>
  $item = @{
    originalName = $mdFile.name
    originalAbsolute = $mdFile.FullName
    originalRelative = Resolve-Path $mdFile.FullName -Relative
    originalFolder   = $mdFile.Directory.Name
    originalLinkRelative = $null
    originalFolderRelative = Resolve-Path $mdFile.Directory -Relative
    safeName = [System.Web.HttpUtility]::UrlDecode($mdFile.Name)
    safeAbsolute = $null
    safeRelative = $null
    safeFolder   = $null
    safeLinkRelative = $null
    safeFolderRelative = $null
    itemType = "File"
    display = [System.Web.HttpUtility]::UrlDecode((split-path $mdFile.FullName -LeafBase).Replace("-", " "))
    renamed = $false
    wikiPath = $null
  }

  $item.wikiPath = $item.originalRelative.Substring(2).replace("$([IO.Path]::DirectorySeparatorChar)","/").replace(".md", "")

  Set-MdYamlHeader -file $mdFile -key "wikiPath" -value $item.wikiPath

  if ($item.originalName -ne $item.safeName)
  {
    Write-Host "Found a file that requires renaming"
    Write-Host "         file name: [$($item.originalName)]"
    Write-Host "    safe file name: [$($item.safeName)]"

    Write-Host "    Renaming [$($mdFile.FullName)] to [$($item.safeName)]"
    #$mdFile.MoveTo((join-path -path $Path -childPath $item.safeName))
    if (Test-Path (join-path $mdFile.Directory -ChildPath $item.safeName))
    {
      write-Host "Removing existing file $($item.safeName)"
      remove-item (join-path $mdFile.Directory -ChildPath $item.safeName) -force
    }
    $mdFile = Rename-Item -path $mdFile.FullName -NewName $item.safeName -PassThru -Force
    Write-Host "  Adding original-filename and new-filename values in yaml header"
    Set-MdYamlHeader -file $mdFile -key "original-filename" -value $item.originalName
    Set-MdYamlHeader -file $mdFile -key "new-filename" -value $item.safeName
    $item.renamed = $true
  }
  $item.safeAbsolute = $mdFile.FullName
  $item.safeRelative = Resolve-Path $mdFile.FullName -Relative
  $item.safeFolder   = $mdFile.Directory.Name
  $item.originalLinkRelative = $item.originalRelative.Replace("$([IO.Path]::DirectorySeparatorChar)", "/")
  $item.safeLinkRelative = $item.safeRelative.Replace("$([IO.Path]::DirectorySeparatorChar)", "/")
  $item.safeFolderRelative = Resolve-path $mdFile.Directory.FullName -relative

  if ($item.safeFolderRelative -like "..$([IO.Path]::DirectorySeparatorChar)*")
  {
    $item.safeFolderRelative = "."
  }

  return [PSCustomObject]$item
}

function script:Set-SafeWikiFolderName
{
  param($folder)

  Write-Host "Checking folder from [$($folder)]"

  <#
    $folder = $folderInFolder | select-object -first 1
    $folder = $folderInFolder | select-object -first 1 -skip 1
    $folder = $folderInFolder | select-object -first 1 -skip 2
    $folder
  #>
    
  $item = @{
    originalName = $folder.Name
    originalAbsolute = $folder.FullName
    originalRelative = Resolve-Path $folder.FullName -Relative
    safeName = [System.Web.HttpUtility]::UrlDecode($folder.Name)
    safeAbsolute = $null
    safeRelative = $null
    parentRelative = (Resolve-Path $folder.Parent.FullName -Relative)
    itemType = "Directory"
    display = ([System.Web.HttpUtility]::UrlDecode($folder.Name.Replace("-"," ")))
    renamed = $false
  }

  if ($item.parentRelative -like "..$([IO.Path]::DirectorySeparatorChar)..$([IO.Path]::DirectorySeparatorChar)*")
  {
    $item.parentRelative = $null
  }
  elseif(($item.parentRelative -like "..$([IO.Path]::DirectorySeparatorChar)*"))
  {
    $item.parentRelative = "."
  }

  if ($item.originalName -ne $item.safeName)
  {
    Write-Host "Found a folder that requires renaming"
    Write-Host "original folder name: [$($item.originalName)]"
    Write-Host "    safe folder name: [$($item.safeName)]"

    Write-Host "Renaming folder $($folder) to $($item.safeName)"
    if (Test-Path (join-path $folder.parent -ChildPath $item.safeName))
    {
      write-Host "Removing existing folder $($item.safeName)"
      remove-item (join-path $folder.Parent -ChildPath $item.safeName) -Recurse -Force
    }      
    #$folder.MoveTo((join-path -path $Path -childPath $item.safeName))
    $folder = Rename-Item $folder.FullName -NewName $item.safeName -PassThru -Force

    $item.renamed = $true
  }

  $item.safeAbsolute = $folder.FullName
  $item.safeRelative = Resolve-Path $folder.FullName -Relative

  return [PSCustomObject]$item
}

function script:Get-MdSections
{
  param($Content)

  $codeRegex = "^(?<code>``{3}\s*\w*\s*)$"

  $sections = [System.Collections.ArrayList]::new()

  $codeSections = $Content | select-string $codeRegex
  $lineStart = 0
  $codeBlock = 0

  Write-Host "Code section count: $($codeSections.count)"

  if ($codeSections.count -gt 0)
  {
    for($i=0;$i -lt $codeSections.count/2;$i++)
    {
      $codeBlock = $i*2
      if ($codeSections[$codeBlock].LineNumber-1 -gt $lineStart)
      {
        $sections.Add([PSCustomObject]@{type="Conceptual";content=$content[$lineStart..($codeSections[$codeBlock].LineNumber -2)]}) | out-null
      }
      $sections.Add([PSCustomObject]@{type="Code";content=$content[($codeSections[$codeBlock].LineNumber-1)..($codeSections[$codeBlock+1].LineNumber-1)]}) | out-null
      $lineStart=$codeSections[$codeBlock+1].LineNumber
    }
    if ($lineStart -lt $content.count)
    {
      $sections.Add([PSCustomObject]@{type="Conceptual";content=$content[$lineStart..($content.count-1)]}) | out-null
    }
  }
  else 
  {
    $sections.Add([PSCustomObject]@{type="Conceptual";content=$content}) | out-null
  }

  return $sections

}

function script:Update-Links
{
  param($Content, $ReplaceCode)

  $private:findRegex = "\[(?'display'(?:[^\[\]]|(?<Open>\[)|(?<Content-Open>\]))+(?(Open)(?!)))\]\((?'link'(?:[^\(\)]|(?<Open>\()|(?<Content-Open>\)))+(?(Open)(?!)))\)"

  if ("$content" -ne "" -and $content -match $findRegex)
  {
    $private:sections = Get-MdSections -Content $content
    
    $private:conceptualSectionNumber = 0
  
    foreach($private:conceptual in $sections | where-object type -eq "Conceptual")
    {
      <#
        $conceptual = $sections | where-object type -eq "Conceptual" | select-object -first 1
      #>
      $conceptualSectionNumber++
      if ($conceptual.content -match $f)
      {        
        Write-Verbose "Conceptual section $conceptualSectionNumber"
        Write-Verbose "Before:"
        $conceptual.content | select-string $findRegex -AllMatches | Out-Host
        $conceptual.content = $conceptual.content -replace $findRegex, $replaceCode
        Write-Verbose "After:"
        $conceptual.content | select-string $findRegex -AllMatches | Out-Host
      }
    }
    $Content = $sections | select-object -ExpandProperty content
  }

  return $Content

}

function script:Update-FixAdoWikiEscapes
{
  param($Content)

  $private:r = {
    $private:in = @{
      display = $_.Groups["display"].Value
      link = $_.Groups["link"].Value  
    }
    <#
    $in = @{}
      $in.display = "This is the display"
      $in.link = "https://user:password@www.contoso.com:80/Home/Index.htm?q1=v1&q2=v2#FragmentName"
      $in.link = "xfer:Home_Index#FragmentName"
      $in.link = "/Home \(escaped folder\)/Index.md?q1=v1&q2=v2#FragmentName"
      $in.link = "/Home/Index\(escaped folder\).md?q1=v1&q2=v2#FragmentName"
    #>
    $private:out = @{
      display = $in.display
      link = $in.link
    }
    if ($out.link.StartsWith("/"))
    {
      $out.link = $out.link.replace("\(", "(").replace("\)", ")")
    }

    $private:ret = "[$($out.display)]($($out.link))"
    return $ret

  }
    
  $private:UpdatedContent = Update-Links -Content $Content -ReplaceCode $r

  return $UpdatedContent
}

function script:Update-ToRelativeLinks
{
  param($Content, [Uri]$PageUri)

  $private:r = {
    $private:in = @{
      display = $_.Groups["display"].Value
      link = $_.Groups["link"].Value  
    }
    <#
    $in = @{}
      $in.display = "This is the display"
      $in.link = "https://user:password@www.contoso.com:80/Home/Index.htm?q1=v1&q2=v2#FragmentName"
      $in.link = "xfer:Home_Index#FragmentName"
      $in.link = "/Home/Index.md?q1=v1&q2=v2#FragmentName"
      $in.link = "Home/Index?q1=v1&q2=v2#FragmentName"
      $in.link = "#Anchor"
      $in.link = "/With%20Space/With%20Space%20Too.md?q1=v1&q2=v2#FragmentName"
      $in.link = "/With Space/With Space Too.md?q1=v1&q2=v2#FragmentName"
      $in.display = Read-Host "Display"
      $in.link = Read-Host "Link"
    #>
    $private:out = @{
      display = $in.display
      link = $in.link
    }
    if ($out.link.StartsWith("/"))
    {
      $private:linkUri = [Uri]::new($baseUri, $out.link)
      
      $out.link = $PageUri.MakeRelativeUri($linkUri).ToString()
    }

    $private:ret = "[$($out.display)]($($out.link))"
    return $ret

  }
    
  $private:UpdatedContent = Update-Links -Content $content -ReplaceCode $r

  return $UpdatedContent
 
}

function script:Update-ToMdLinks
{
  param($Content, $AbsoluteLinkList)

  $private:r = {
    $private:in = @{
      display = $_.Groups["display"].Value
      link = $_.Groups["link"].Value
    }
    <#
    $in = @{}
      $in.display = "This is the display"
      $in.link = "https://user:password@www.contoso.com:80/Home/Index.htm?q1=v1&q2=v2#FragmentName"
      $in.link = "xfer:Home_Index#FragmentName"
      $in.link = "/Home/Index.md?q1=v1&q2=v2#FragmentName"
      $in.link = "Home/Index?q1=v1&q2=v2#FragmentName"
      $in.link = "#Anchor"
      $in.link = "/With%20Space/With%20Space%20Too.md?q1=v1&q2=v2#FragmentName"
      $in.link = "/With Space/With Space Too.md?q1=v1&q2=v2#FragmentName"
      $in.display = Read-Host "Display"
      $in.link = Read-Host "Link"
    #>
    $private:out = @{
      display = $in.display
      link = $in.link
      left = ""
      right = ""
    }

    if ($out.link.StartsWith("/") -and -not $out.link.StartsWith("/.attachments"))
    {

      $private:segmentIndex = $out.link.IndexOfAny(("?","#"))

      if ($segmentIndex -ge 0)
      {
        $out.left = $out.link.Substring(0,$segmentIndex)
        $out.right = $out.link.Substring($segmentIndex)
      }
      else
      {
        $out.left = $out.link
      }

      $private:linkSegments = $out.left.split("/")

      # TODO: Might be too harsh/hard core, might need a lot more subtleties, not all links point to a .md file, right ?
      $linkSegments[-1] = "$($linkSegments[-1]).md"

      $out.left = $linkSegments -join "/"

      $out.link = "$($out.left)$($out.right)"
    }
 
    $private:ret = "[$($out.display)]($($out.link))"
    return $ret

  }
    
  $private:updatedContent = Update-Links -Content $content -ReplaceCode $r

  return $updatedContent
 
}

function script:Update-RenamedLinks
{
  param($Content)

  $private:r = {
    $private:in = @{
      display = $_.Groups["display"].Value
      link = $_.Groups["link"].Value
    }
    <#
    $in = @{}
      $in.display = "This is the display"
      $in.link = "https://user:password@www.contoso.com:80/Home/Index.htm?q1=v1&q2=v2#FragmentName"
      $in.link = "xfer:Home_Index#FragmentName"
      $in.link = "/Home/Index.md?q1=v1&q2=v2#FragmentName"
      $in.link = "Home/Index?q1=v1&q2=v2#FragmentName"
      $in.link = "#Anchor"
      $in.link = "/With%20Space/With%20Space%20Too.md?q1=v1&q2=v2#FragmentName"
      $in.link = "/With Space/With Space Too.md?q1=v1&q2=v2#FragmentName"
      $in.display = Read-Host "Display"
      $in.link = Read-Host "Link"
    #>
    $private:out = @{
      display = $in.display
      link = $in.link
      uri = [Uri]::new($baseUri, $in.link)
      left = ""
      right = ""
    }

    if ("$($out.uri.Scheme)://$($out.uri.host)" -ne $baseUrl)
    {
      Write-Verbose "link is external"
    }
    else 
    {
      Write-Verbose "link is internal"

      $private:segmentIndex = $out.link.IndexOfAny(("?","#"))

      if ($segmentIndex -ge 0)
      {
        $out.left = $out.link.Substring(0,$segmentIndex)
        $out.right = $out.link.Substring($segmentIndex)
      }
      else
      {
        $out.left = $out.link
      }

      Write-Verbose "  Left : $($out.left)"
      Write-Verbose "  Right: $($out.right)"

      $private:fixUri = [Uri]::new($baseUri, $out.left)

      $out.left = $fixUri.AbsolutePath

      $out.link = "$($out.left)$($out.right)"
    }
 
    $private:ret = "[$($out.display)]($($out.link))"
    return $private:ret

  }
    
  $private:updatedContent = Update-Links -Content $content -ReplaceCode $r

  return $updatedContent
}

function script:Get-PageUid
{
  param($mdFile)

  $workingDirectory = (Get-Location)

  $relative = (join-path $mdFile.Directory.FullName -ChildPath $mdFile.BaseName).Substring($workingDirectory.Path.Length)
  
  $pageSegments = $relative.Replace(" ", "_").Split("$([IO.Path]::DirectorySeparatorChar)",[System.StringSplitOptions]::RemoveEmptyEntries)
  
  $pageUid = $pageSegments -join "_"
  Write-Verbose "File: $(Resolve-Path -path $mdfile.FullName -Relative) UID: $pageUid"

  return $pageUid

}

function script:Add-ChildWikiOrderItem
{
  param($OrderFile, $OrderTitle, $OrderIndex)

  [System.Collections.ArrayList]$orderItems = Get-Content -Path $OrderFile

  $orderItems.Insert($OrderIndex, $OrderTitle)

  $orderItems | Set-Content -Path $OrderFile
}


function script:Convert-FromWikiOrder
{
  param([System.IO.FileInfo]$Order)

  $workingDirectory = (Get-Location)
  
  $o = [ordered]@{
    orderFile           = $Order
    content             = @() + (Get-Content -path $Order)
    folderAbsolute      = $Order.Directory.FullName
    folderName          = $Order.Directory.Name
    folderRelative      = $null
    folderUri           = $null
    depth               = $null
    orderItems          = [System.Collections.ArrayList]::new()
  }
  $o.folderRelative = $o.folderAbsolute.Substring($workingDirectory.Path.Length)
  $o.folderUri = [Uri]::new($baseUri, $o.folderRelative.replace("$([IO.Path]::DirectorySeparatorChar)", "/"))
  $o.depth = $o.folderRelative.Split("$([IO.Path]::DirectorySeparatorChar)").Count - 1
      
  foreach($orderItem in $o.content)
  {
    <#
      $orderItems

      $orderItem = $o.content | select-object -first 1
      $orderItem = $o.content | select-object -first 1 -skip 1
      $orderItem = $o.content | select-object -last 1

      $orderItem = "Foo"
      $orderItem = "Foo-Bar"
      $orderItem = "Foo-Bar-(Snafu)"
    #>

    if ("$orderItem" -ne "")
    {
    
      Write-Debug "OrderItem: $orderItem"
      
      $oi = [ordered]@{
        orderItem              = $orderItem
        orderItemMd            = "$($orderItem).md"
        orderItemMdUri         = $null
        orderItemFolderPath    = Join-Path -path $order.Directory.FullName -ChildPath $orderItem
        display                = [System.Web.HttpUtility]::UrlDecode($orderItem.Replace("-", " "))
      }
      $oi.orderItemMdUri = [Uri]::new($o.folderUri, $oi.orderItemMd)

      $o.orderItems.Add([PSCustomObject]$oi) | Out-Null
    }
  }

  return [PSCustomObject]$o

}



function script:ConvertTo-DocFxToc
{
  param($OrderItems, $depth)

  $tocItems = [System.Collections.ArrayList]::new()
  
  foreach ($orderItem in $OrderItems)
  {
    <#
      $orderItem = $OrderItems | select-object -first 1
    #>

    $tocItem = [ordered]@{
      name = $orderItem.display
      href = $null
    }

    if (Test-Path $orderItem.orderItemFolderPath)
    {

      if ($depth -eq 0)
      {
        <#
          name: some thing
          href: some-thing/
        #>
        $tocItem.href = "$($orderItem.orderItem)/"
      }
      else
      {
        <#
          name: some thing
          href: some-thing/toc.yml
        #>
        $tocItem.href = "$($orderItem.orderItem)/toc.yml"
      }

      $tocItem.topicHref = "$($orderItem.orderItemMd)"

    }
    else
    {
      <#
        name: some thing
        href: some-thing.md
      #>
      $tocItem.href = $orderItem.orderItemMd

    }

    $tocItems.Add([PSCustomObject]$tocItem) | out-null

  }

  return @{
    items = $tocItems
  }

}


function script:Update-MermaidCodeDelimiter
{
  param($mdfile)

  $content = get-content -path $mdfile.FullName -raw
  if ("" -ne "$content" -and ("$content".Contains(":::mermaid") -or "$content".Contains("::: mermaid")))
  {
    Write-Host "Found Mermaid Code in $($mdfile.FullName). Fixing..."
    $content = $content.replace(":::mermaid", "``````mermaid")
    $content = $content.replace("::: mermaid", "``````mermaid")
    $content = $content.replace(":::", "``````")
    set-content -path $mdfile.FullName -value $content    
  }
}

function script:Update-AdoWikiPreMerge
{
  param($Path, $Depth)
  
  Write-Host "Performing PreMerge actions to [$Path]"
  
  push-location $Path
  
  $private:renameMap = [System.Collections.ArrayList]::new()
  $private:allMetadata = Get-AdoWikiMetadata -Path .
  $private:workingDirectory = (Get-Location)

  # ------------------------------------------------------------------------
  Write-Host "   - Convert .order to toc.yml"

  $private:folders = Get-AdoWikiFolders -Path . -Exclude @(".git", ".attachments")
    
  foreach($private:folder in $folders)
  {
    <#
      $folder = $folders | select-object -first 1
      $folder = $folders | select-object -first 1 -skip 1
    #>

    $private:dot_order = Join-Path $folder -ChildPath ".order"

    if (Test-Path $dot_order)
    {
      $private:dot_order = Get-Item (Join-Path $folder -ChildPath ".order")
  
      # $Order = $order
      # $MetadataItems = $metadataItemsInFolder
      $private:adoWikiOrder = Convert-FromWikiOrder -Order $dot_order
      $private:totalDepth = $Depth + $folder.substring($workingDirectory.Path.Length).split("$([IO.Path]::DirectorySeparatorChar)").count - 1
  
      if (($adoWikiOrder.orderItems | select-object -first 1).orderItem -eq "Index")
      {
        $private:orderItemsExceptIndex = $adoWikiOrder.orderItems | select-object -skip 1
      }
      else
      {
        $private:orderItemsExceptIndex = $adoWikiOrder.orderItems 
      }
  
      <#
        $OrderItems = $orderItemsExceptIndex 
        $depth = $depth
      #>
  
      $private:toc = ConvertTo-DocFxToc -OrderItems $orderItemsExceptIndex -depth $totalDepth
    }
    else
    {
      $private:toc = @{items = @()}
    }
  
    ConvertTo-Yaml $toc -OutFile (Join-Path $folder -ChildPath "toc.yml") -Force
  }





  # ------------------------------------------------------------------------
  Write-Host "   - Set Yaml Headers"
  Write-Host "     - WikiPath"
  Write-Host "     - uid"
  foreach($metadata in $allMetadata)
  {
    $mdFile = $metadata.DocFxSafe.File
    $adoWikiOriginalMd = $mdFile.FullName.Substring($workingDirectory.Path.Length)

    Set-MdYamlHeader -file $mdFile -key "adoWikiPath" -value $metadata.AdoWiki.WikiPath
    Set-MdYamlHeader -file $mdFile -key "adoWikiOriginalMd" -value $adoWikiOriginalMd
  }



  # ------------------------------------------------------------------------
  Write-Host "   - Rename [md Files] to DocFx safe name format"
  foreach($metadata in $allMetadata)
  {
    <#
      $metadata = $allMetadata | select-object -first 1
      
    #>
    $mdFile = $metadata.AdoWiki.File

    if ($metadata.DocFxSafe.RenameRequired)
    {
      $renameMap.Add([PSCustomObject]@{
        from = $metadata.AdoWiki.FileRelativeUri
        to  = $metadata.DocFxSafe.FileRelativeUri
      }) | Out-Null

      Write-Host "   - File $($metadata.AdoWiki.Filename) is not DocFx safe, rename required"
      $filePathToRename = $metadata.AdoWiki.FileAbsolute
      $newName = $metadata.DocFxSafe.FileName

      $metadata.DocFxSafe.File = Rename-Item -Path $filePathToRename -NewName $newName -Force -PassThru
      $metadata.DocFxSafe.FileIsRenamed     = $true

      Set-MdYamlHeader -file $metadata.DocFxSafe.File -key "DocFxSafeFileName" -value $newName

      $toc_yaml = (join-path $mdFile.Directory.FullName -childPath "toc.yml")
      $toc = get-content $toc_yaml | ConvertFrom-yaml -Ordered

      $tocItem = $toc.items | where-object {$_.href -eq $mdFile.Name -or $_.topicHref -eq $mdFile.Name}

      if ($tocItem)
      {
        if ($tocItem.href -eq $mdFile.Name)
        {
          $tocItem.href = $newName
        }
        else
        {
          $tocItem.topicHref = $newName
        }
      }
      else
      {
        Write-Warning "$($mdFile.FullName) not found in $toc_yaml"
      }

      ConvertTo-Yaml -Data $toc -OutFile $toc_yaml -Force
    }     
  }




  # ------------------------------------------------------------------------
  Write-Host "   - Rename [Folders] to DocFx safe name format"

  $foldersMetadata = [System.Collections.ArrayList]::new()
  foreach($metadata in $allMetadata)
  {

    $folder = $metadata.AdoWiki.File.Directory

    if (!($foldersMetadata | where-object {$_.Folder.Fullname -eq $folder.FullName}))
    {
      $foldersMetadata.Add([PSCustomObject]@{
        Folder = $folder
        FolderRelative = $folder.FullName.Substring($workingDirectory.Path.Length)
        Depth = $folder.FullName.Split("$([IO.Path]::DirectorySeparatorChar)").Count
      }) | out-null
    }
  }

  foreach($folderMetadata in $foldersMetadata | sort-object Depth -Descending)
  {
    <#
      $folderMetadata = $foldersMetadata | sort-object Depth -Descending | select-object -first 1
      $folderMetadata = $foldersMetadata | select-object -first 1 -skip 1

    #>    
    $folderUri = [Uri]::new($baseUri, $folderMetadata.FolderRelative.Replace("$([IO.Path]::DirectorySeparatorChar)", "/"))

    if ($folderUri.AbsoluteUri -ne $folderUri.OriginalString)
    {
      Write-Host "   - Folder $($folderMetadata.FolderRelative) is not DocFx safe, rename required"

      $filePathToRename = $folderMetadata.Folder.FullName
      $oldName = $folderMetadata.Folder.Name
      $newName = $folderUri.Segments[-1]
      Write-Host "      From: $($oldName)"
      Write-Host "        To: $($newName)"
      Rename-Item -Path $filePathToRename -NewName $newName -Force

      $renameMap.Add([PSCustomObject]@{
        from = "$($oldName)/"
        to  = "$($newName)/"
      }) | Out-Null

      $toc_yaml = join-path $folderMetadata.Folder.Parent.FullName -ChildPath "toc.yml"
      $toc = get-content $toc_yaml | ConvertFrom-Yaml -Ordered

      foreach($tocItem in $toc.items)
      {
        <#
          $tocItem = $toc.items | select-object -first 1
          $tocItem = $toc.items | select-object -first 1 -skip 1
          $tocItem = $toc.items | select-object -first 1 -skip 2
        #>
        if ($tocItem.href.StartsWith("$($oldName)/"))
        {
          $segments = $tocItem.href.split("/")
          $segments[0] = $newName
          $tocItem.href = $segments -join "/"
        }

        if ("$($tocItem.topicHref)".StartsWith("$($oldName)/"))
        {
          $segments = $tocItem.topicHref.split("/")
          $segments[0] = $newName
          $tocItem.topicHref = $segments -join "/"
        }

      }

      ConvertTo-Yaml -Data $toc -OutFile $toc_yaml -Force

    }
  }




  # ------------------------------------------------------------------------
  Write-Host "   - Update Hyperlinks"
  Write-Host "     - Convert absolute links to relative"
  Write-Host "     - Update wiki links to .md extension"
  Write-Host "     - Update wiki links to match the renamed mdFiles or folder"

  $allMetadata = Get-AdoWikiMetadata -Path .

  $VerbosePreference = 'Continue'
  $DebugPreference = 'Continue'

  foreach($metadata in $allMetadata)
  {
    $private:mdFile = $metadata.DocFxSafe.File

    $private:content = Get-Content -Path $mdFile
    
    $content = Update-FixAdoWikiEscapes -content $Content
    
    # /foo/bar -> /foo/bar.md
    $content = Update-ToMdLinks -content $Content -AbsoluteLinkList $allMetadata.AdoWiki.LinkAbsolute
    
    # /foo/bar.md -> [[../]foo/]bar.md depends on the current page's uri
    $content = Update-RenamedLinks -Content $content
    
    # /foo/bar.md -> [[../]foo/]bar.md depends on the current page's uri
    $private:pageUri = [Uri]::new($baseUri, $mdFile.FullName.Substring($workingDirectory.Path.Length))
    $content = Update-ToRelativeLinks -content $content -PageUri $pageUri
       

    $content | Set-Content -Path $mdFile
  }





  # ------------------------------------------------------------------------
  Write-Host "   - Update Mermaid Code Delimiters"

  foreach($metadata in $allMetadata)
  {
    $mdFile = $metadata.DocFxSafe.File

    Update-MermaidCodeDelimiter -mdfile $mdFile
  }


  pop-location # $target  
}

function script:Update-RootAdoWikiTopNavigation
{
  param([System.IO.DirectoryInfo]$Path)

  Write-Host "Special care for mdFiles that should be in their subfolder"

  $private:mdFilesMoved = [System.Collections.ArrayList]::new()

  $private:toc_yaml = Get-Item (Join-path $Path.FullName -ChildPath "toc.yml")
  $private:toc = get-content $toc_yaml | ConvertFrom-Yaml -Ordered

  foreach($private:tocItem in $toc.items | where-object {"$($_.topicHref)" -ne "" -and -not "$($_.topicHref)".contains("/") })
  {
    <#
      $tocItem = $toc.items | where-object {"$($_.topicHref)" -ne "" -and -not "$($_.topicHref)".contains("/") } | select-object -first 1
    #>
    
    $private:topicHrefSegments = $tocItem.topicHref.Replace("$([IO.Path]::DirectorySeparatorChar)", "/").Split("/")
    $private:sourceFile = $topicHrefSegments[-1]

    if ($topicHrefSegments.count -ne 1)
    {
      Write-Host "Already moved?  Confirm at $($tocItem.topicHref)"
    }
    elseif(-not (test-path -Path (Join-Path $Path.FullName -ChildPath $sourceFile)))
    {
      Write-Host "File not found at $($tocItem.topicHref)"
    }
    else
    {
      $private:targetFolder = $tocItem.href.Replace("$([IO.Path]::DirectorySeparatorChar)", "/").Split("/")[0]
      $private:targetFile = $null

      $private:tryThese = @($sourceFile, "index.md")
      foreach($tryThisfilename in $tryThese)
      {
        $private:testFilename = $tryThisfilename
        $private:testPath = join-path $Path.FullName -ChildPath $targetFolder -AdditionalChildPath $testFilename
  
        if (Test-Path $testPath)
        {
          Write-Host "    Another $testFilename exists in $targetFolder"
        }
        else
        {
          $targetFile = $testFilename
          break
        }
      }
  
      if ($null -eq $targetFile)
      {
        $private:i = 0
        $private:baseFilename = split-path $sourceFile -LeafBase
        do
        {
          
          $i++
          $private:testFilename = "$($baseFilename)_$($i).md"
          $private:testPath = join-path $Path.FullName -ChildPath $targetFolder -AdditionalChildPath $testFilename
    
          if (Test-Path $testFilename)
          {
            Write-Host "    Another $testFilename exists in $targetFolder"
          }
          else
          {
            $targetFile = $testFilename
          }
        } until($targetFile -or $i -gt 1000)
      }
  
      if ($targetFile)
      {
  
        $private:sourcePath = join-path $Path.FullName -ChildPath $sourceFile
        $private:targetPath = join-path $Path.FullName -ChildPath $targetFolder -AdditionalChildPath $targetFile
  
        Write-Host "    Moving $sourceFile to $targetPath"
  
        $private:oldTopicHref = $tocItem.topicHref 
        $private:newTopicHref = "$($targetFolder)/$($targetFile)"

        Move-Item -Path $sourcePath -Destination $targetPath
        $tocItem.topicHref = $newTopicHref
                
        $mdFilesMoved.Add([PSCustomObject]@{
          old_Path = $sourcePath
          old_Folder = $Path.FullName
          old_Filename = $sourceFile
          old_topicHref = $oldTopicHref
          new_Path = $targetPath
          new_Folder = join-path $Path.FullName -ChildPath $targetFolder
          new_Filename = $targetFile
          new_topicHref = $newTopicHref
        }) | out-null
        
      }
      else
      {
        Write-Error "Couldn't find a filename to put $($sourceFile) in $($targetFolder), and I tried..."
      }
    }
  }

  ConvertTo-Yaml -Data $toc -OutFile $toc_yaml -Force

  return $mdFilesMoved
}

function script:Merge-ChildAdoWiki
{  
  param($RootWikiLocation, $ChildWikiSource, $ChildWikiTargetSubfolder, $MenuDisplayName, $MenuPosition)

  <#
    $RootWikiLocation
    $ChildWikiSource
    $ChildWikiTargetSubfolder
    $MenuDisplayName
    $MenuPosition
  #>

  Push-Location $RootWikiLocation
  $private:workingDirectory = (Get-Location)

  $private:targetFolderFullpath = Join-Path $workingDirectory -childPath $ChildWikiTargetSubfolder
  
  Write-Host "    - RoboCopy child wiki to the parent wiki's target folder"
  $private:robocopyChildSource = $ChildWikiSource
  $private:robocopyChildTarget = $targetFolderFullpath
  $private:robocopyLoggingArgs = @("/NS", "/NC", "/NFL", "/NDL", "/NP")
  Write-Host "      From: $robocopyChildSource"
  Write-Host "        To: $robocopyChildTarget"
  robocopy.exe $robocopyChildSource $robocopyChildTarget /E @robocopyLoggingArgs

  Write-Host "    - Insert child wiki orderItem in parent wiki's parent folder's toc.yml (mouthfull)"

  $private:targetFolder = get-item $targetFolderFullpath

  $private:depth = $targetFolder.Parent.FullName.Substring($workingDirectory.Path.Length).split("$([IO.Path]::DirectorySeparatorChar)").Count-1
  
  $private:newItem = [ordered]@{
    name = $MenuDisplayName
    href = $null
    topicHref = $null
  }

  if ($depth -eq 0)
  {
    $newItem.href = "$($targetFolder.Name)/"
  }
  else
  {
    $newItem.href = "$($targetFolder.Name)/toc.yml"
  }

  $private:childToc_yml = Get-ChildItem -Path $targetFolder -Filter "toc.yml"
  $private:childIndex_md = Get-ChildItem -Path $targetFolder -Filter "index.md"

  if ($childIndex_md)
  {
    $newItem.topicHref = "$($targetFolder.Name)/$($childIndex_md.Name)"
  }
  else
  {
    $private:firstChildMenuItem = (Get-Content $childToc_yml | ConvertFrom-Yaml -Ordered).items | where-object {$_.href -like "*.md"} | select-object -first 1

    if (!$firstChildMenuItem)
    {
      $firstChildMenuItem = Get-ChildItem -Path $targetFolder -Filter "*.md" | select-object -first 1
    }

    $newItem.topicHref = "$($targetFolder.Name)/$($firstChildMenuItem.href)"
  }

  $private:parentToc_yml = Get-ChildItem -Path $targetFolder.Parent -Filter "toc.yml"
  $private:parentToc = get-content $parentToc_yml | ConvertFrom-Yaml -Ordered

  $private:parentToc.items.Insert($MenuPosition, $newItem)

  $parentToc | ConvertTo-Yaml -OutFile $parentToc_yml -Force

  Pop-Location

}



function script:Get-AdoWikiMetadata
{
  param($Path)

  push-location $Path

  $metadataList = [System.Collections.ArrayList]::new()

  $folders = Get-AdoWikiFolders -Path . -Exclude @(".git", ".attachments")
    
  foreach($folder in $folders)
  {
    <#
      $folder = $folders | select-object -first 1
    #>
    $mdFiles = Get-WikiMarkdowns -Folder $folder   
  
    foreach($mdFile in $mdFiles)
    {
      $metadata = Get-DocfxItemMetadata -mdFile $mdFile
  
      $metadataList.Add($metadata) | Out-Null
    }    
  }
  
  Pop-Location # docs

  return $metadataList
}


function script:Convert-ToDocfx
{ 
  param($Path)

  push-location $Path

  Write-Host "Converting ADO Wiki to DocFx"

  $folders = Get-AdoWikiFolders -Path . -Exclude @(".git", ".attachments")
  
  Write-Host "  - Generate page's uid"
  Write-Host "  - Convert .order to toc.yml"
  
  foreach($folder in $folders)
  {
    <#
      $folder = $folders | select-object -first 1
      $folder = $folders | select-object -first 1 -skip 1
    #>

    $mdFiles = Get-WikiMarkdowns -Folder $folder
  
    foreach($mdFile in $mdFiles)
    {
      <#
        $mdFile = $mdFiles | select-object -first 1
        $mdFile = $mdFiles | select-object -first 1 -skip 1
      #>
    
      $pageId = Get-PageUid -mdFile $mdFile
      Set-MdYamlHeader -file $mdFile -key "uid" -value $pageId
    
    }


    
  }

  Write-Host "... Converted"
  
  Pop-Location # docs
}
