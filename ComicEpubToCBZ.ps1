param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$FilePath = $args,
    
    [Parameter(Mandatory=$false)]
    [switch]$Help
)

# Display help
if ($Help -or (-not $FilePath -or $FilePath.Count -eq 0)) {
    Write-Host "EPUB to CBZ Processing Tool" -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Cyan
    Write-Host "This script processes EPUB comic/manga files to extract and organize images in the correct order."
    Write-Host "It produces a properly ordered CBZ file in the same directory as the input file."
    Write-Host ""
    Write-Host "Usage:" -ForegroundColor Yellow
    Write-Host "  1. Drag and drop EPUB files onto this script, or"
    Write-Host "  2. Run it from PowerShell with file paths as arguments"
    Write-Host ""
    Write-Host "Parameters:" -ForegroundColor Yellow
    Write-Host "  FilePath : Path(s) to EPUB file(s) - can be multiple"
    Write-Host "  -Help    : Display this help information"
    Write-Host ""
    Write-Host "Example:" -ForegroundColor Green
    Write-Host "  .\ProcessEPUB.ps1 'C:\Comics\MyComic.epub'"
    
    if (-not $Help) {
        Write-Host "`nNo files provided. Please drag and drop EPUB files onto this script."
        if ($Host.Name -eq 'ConsoleHost') {
            Write-Host "`nPress any key to exit..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
    exit
}

Write-Host "Found $($FilePath.Count) files to process..."

# Set the temp directory
$TEMP_DIR = "$env:TEMP\epub_processor"

# Create temp directory if it doesn't exist
New-Item -ItemType Directory -Force -Path $TEMP_DIR | Out-Null

function Get-SpineOrder {
    param (
        [string]$ContentOPFPath
    )
    
    if (-not (Test-Path -LiteralPath $ContentOPFPath)) {
        Write-Error "content.opf file not found at: $ContentOPFPath"
        return $null
    }
    
    try {
        [xml]$opfContent = Get-Content -LiteralPath $ContentOPFPath -Encoding UTF8
        
        # Extract spine order
        $spineItems = $opfContent.package.spine.itemref
        $idRefs = $spineItems | ForEach-Object { $_.idref }
        
        # Map idref to href
        $manifest = $opfContent.package.manifest.item
        $pageOrder = @()
        
        foreach ($idRef in $idRefs) {
            $item = $manifest | Where-Object { $_.id -eq $idRef }
            if ($item) {
                $pageOrder += $item.href
            }
        }
        
        return $pageOrder
    }
    catch {
        Write-Error "Error parsing content.opf: $_"
        return $null
    }
}

function Get-PageMapping {
    param (
        [string]$TOCPath
    )
    
    if (-not (Test-Path -LiteralPath $TOCPath)) {
        Write-Error "TOC file not found at: $TOCPath"
        return $null
    }
    
    try {
        [xml]$tocContent = Get-Content -LiteralPath $TOCPath -Encoding UTF8
        
        # Extract page to label mapping from navMap
        $navPoints = $tocContent.ncx.navMap.navPoint
        $pageMap = @{}
        
        foreach ($navPoint in $navPoints) {
            $label = $navPoint.navLabel.text.'#text'
            $content = $navPoint.content.src
            
            if ($label -and $content) {
                $pageMap[$content] = $label
            }
        }
        
        return $pageMap
    }
    catch {
        Write-Error "Error parsing TOC: $_"
        return $null
    }
}

function Find-ContentFile {
    param (
        [string]$WorkingDir,
        [string]$FilePattern = "*.opf"  # Now accepts a pattern instead of exact filename
    )
    
    Write-Host "  Looking for content file ($FilePattern)..."
    
    # First check if container.xml exists - the standard way to locate content.opf
    $containerPath = Join-Path -Path $WorkingDir -ChildPath "META-INF\container.xml"
    if (Test-Path -LiteralPath $containerPath) {
        Write-Host "  Found container.xml, checking for content path..."
        try {
            [xml]$containerXml = Get-Content -LiteralPath $containerPath -Encoding UTF8
            $rootFilePath = $containerXml.container.rootfiles.rootfile.full_path
            if ($rootFilePath) {
                $fullContentPath = Join-Path -Path $WorkingDir -ChildPath $rootFilePath
                if (Test-Path -LiteralPath $fullContentPath) {
                    Write-Host "  Found content file via container.xml at: $fullContentPath"
                    return $fullContentPath
                }
            }
        }
        catch {
            Write-Warning "  Error parsing container.xml: $_"
        }
    }
    
    # If container.xml didn't work, try common locations
    $commonDirs = @(
        "",              # Root directory
        "OEBPS\",        # Standard EPUB layout
        "OPS\",          # Alternative EPUB layout
        "EPUB\",         # Another alternative
        "content\"       # Sometimes used
    )
    
    foreach ($dir in $commonDirs) {
        $searchPath = Join-Path -Path $WorkingDir -ChildPath $dir
        if (Test-Path -Path $searchPath) {
            $files = Get-ChildItem -Path $searchPath -Filter $FilePattern -File -ErrorAction SilentlyContinue
            if ($files.Count -gt 0) {
                Write-Host "  Found content file at: $($files[0].FullName)"
                return $files[0].FullName
            }
        }
    }
    
    # If still not found, search recursively (limit depth to prevent excessive searching)
    Write-Host "  Searching recursively for $FilePattern..."
    $files = Get-ChildItem -Path $WorkingDir -Filter $FilePattern -File -Recurse -Depth 5 -ErrorAction SilentlyContinue
    
    if ($files.Count -gt 0) {
        # Sort by name to prioritize content.opf or package.opf over other .opf files
        $priorityFiles = $files | Where-Object { $_.Name -eq "content.opf" -or $_.Name -eq "package.opf" }
        if ($priorityFiles.Count -gt 0) {
            Write-Host "  Found preferred content file at: $($priorityFiles[0].FullName)"
            return $priorityFiles[0].FullName
        }
        
        Write-Host "  Found content file at: $($files[0].FullName)"
        return $files[0].FullName
    }
    
    return $null
}

# Add this function to compare file contents
function Compare-FileContent {
    param (
        [string]$Path1,
        [string]$Path2
    )
    
    if (-not (Test-Path -LiteralPath $Path1) -or -not (Test-Path -LiteralPath $Path2)) {
        return $false
    }
    
    $file1Hash = Get-FileHash -LiteralPath $Path1 -Algorithm MD5
    $file2Hash = Get-FileHash -LiteralPath $Path2 -Algorithm MD5
    
    return $file1Hash.Hash -eq $file2Hash.Hash
}

# Add this function to create ComicInfo.xml from EPUB metadata
function Create-ComicInfo {
    param (
        [string]$ContentOPFPath,
        [string]$OutputPath,
        [int]$PageCount
    )
    
    if (-not (Test-Path -LiteralPath $ContentOPFPath)) {
        Write-Warning "Cannot create ComicInfo.xml - content file not found"
        return
    }
    
    try {
        # Create XML document for ComicInfo
        $xmlSettings = New-Object System.Xml.XmlWriterSettings
        $xmlSettings.Indent = $true
        $xmlSettings.IndentChars = "  "
        $xmlSettings.Encoding = [System.Text.Encoding]::UTF8
        
        $xmlWriter = [System.Xml.XmlWriter]::Create($OutputPath, $xmlSettings)
        $xmlWriter.WriteStartDocument()
        
        # Root element with namespace
        $xmlWriter.WriteStartElement("ComicInfo")
        $xmlWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
        $xmlWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")
        
        # Parse EPUB metadata
        [xml]$opfContent = Get-Content -LiteralPath $ContentOPFPath -Encoding UTF8
        $metadata = $opfContent.package.metadata
        
        # Map EPUB metadata to ComicInfo fields
        
        # Title
        $title = $metadata.'dc:title'
        if ($title) {
            $xmlWriter.WriteElementString("Title", $title)
        }
        
        # Series - Try to extract from title if it contains "Vol." or similar
        if ($title -match '(.+?)(?:\s*[,:]\s*|\s+)(?:Vol(?:\.|ume)?\.?\s*)(\d+)') {
            $seriesName = $matches[1].Trim()
            $number = $matches[2]
            $xmlWriter.WriteElementString("Series", $seriesName)
            $xmlWriter.WriteElementString("Number", $number)
        } elseif ($title -match '(.+?)\s+#(\d+)') {
            $seriesName = $matches[1].Trim()
            $number = $matches[2]
            $xmlWriter.WriteElementString("Series", $seriesName)
            $xmlWriter.WriteElementString("Number", $number)
        } else {
            # No volume/issue number found, use title as series
            $xmlWriter.WriteElementString("Series", $title)
            $xmlWriter.WriteElementString("Number", "1")
        }
        
        # Authors - Writers and Artists
        $creators = $metadata.'dc:creator'
        if ($creators) {
            $writers = @()
            $artists = @()
            
            if ($creators -is [array]) {
                foreach ($creator in $creators) {
                    if ($creator.PSObject.Properties['role']) {
                        if ($creator.role -eq "aut") {
                            $writers += $creator.'#text'
                        } elseif ($creator.role -eq "art") {
                            $artists += $creator.'#text'
                        }
                    } else {
                        $writers += $creator
                    }
                }
            } else {
                $writers += $creators
            }
            
            if ($writers.Count -gt 0) {
                $xmlWriter.WriteElementString("Writer", ($writers -join ", "))
            }
            
            if ($artists.Count -gt 0) {
                $xmlWriter.WriteElementString("Penciller", ($artists -join ", "))
                $xmlWriter.WriteElementString("Inker", ($artists -join ", "))
            }
        }
        
        # Publisher
        $publisher = $metadata.'dc:publisher'
        if ($publisher) {
            $xmlWriter.WriteElementString("Publisher", $publisher)
        }
        
        # Publication Date
        $date = $metadata.'dc:date'
        if ($date) {
            if ($date -match '(\d{4})') {
                $year = $matches[1]
                $xmlWriter.WriteElementString("Year", $year)
            }
            
            try {
                $pubDate = [DateTime]::Parse($date)
                $xmlWriter.WriteElementString("Month", $pubDate.Month)
                $xmlWriter.WriteElementString("Day", $pubDate.Day)
            } catch {
                Write-Verbose "Could not parse full date from $date"
            }
        }
        
        # Genre
        $subject = $metadata.'dc:subject'
        if ($subject) {
            if ($subject -is [array]) {
                $xmlWriter.WriteElementString("Genre", ($subject -join ", "))
            } else {
                $xmlWriter.WriteElementString("Genre", $subject)
            }
        }
        
        # Summary/Description
        $description = $metadata.'dc:description'
        if ($description) {
            $xmlWriter.WriteElementString("Summary", $description)
        }
        
        # Language
        $language = $metadata.'dc:language'
        if ($language) {
            $xmlWriter.WriteElementString("LanguageISO", $language)
        }
        
        # Page Count
        $xmlWriter.WriteElementString("PageCount", $PageCount)
        
        # Manga - look for right-to-left reading direction
        if ($opfContent.package.spine -and $opfContent.package.spine.HasAttribute("page-progression-direction")) {
            $direction = $opfContent.package.spine.GetAttribute("page-progression-direction")
            if ($direction -eq "rtl") {
                $xmlWriter.WriteElementString("Manga", "Yes")
            } else {
                $xmlWriter.WriteElementString("Manga", "No")
            }
        } elseif ($title -match "manga|manhua|manhwa" -or $subject -match "manga|manhua|manhwa") {
            $xmlWriter.WriteElementString("Manga", "Yes")
        }
        
        # Close everything
        $xmlWriter.WriteEndElement() # ComicInfo
        $xmlWriter.WriteEndDocument()
        $xmlWriter.Close()
        
        Write-Host "  - Created ComicInfo.xml with metadata"
    }
    catch {
        Write-Warning "Error creating ComicInfo.xml: $_"
    }
}

function Process-EPUB {
    param (
        [string]$EpubFile
    )
    
    if (-not (Test-Path -LiteralPath $EpubFile)) {
        Write-Error "EPUB file not found: $EpubFile"
        return
    }
    
    # Get filename without extension
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($EpubFile)
    $outputDir = [System.IO.Path]::GetDirectoryName($EpubFile)
    $cbzFile = Join-Path -Path $outputDir -ChildPath "$filename.cbz"
    
    # Create working directory
    $workingDir = Join-Path -Path $TEMP_DIR -ChildPath $filename
    New-Item -ItemType Directory -Force -Path $workingDir | Out-Null
    
    # Create temp directory for ordered images
    $orderedImagesDir = Join-Path -Path $workingDir -ChildPath "ordered_images"
    New-Item -ItemType Directory -Force -Path $orderedImagesDir | Out-Null
    
    try {
        Write-Host "Processing: $EpubFile"
        
        # Create temporary zip file and extract it
        $tempZip = Join-Path -Path $TEMP_DIR -ChildPath "$filename.zip"
        Copy-Item -Path $EpubFile -Destination $tempZip -Force
        Expand-Archive -Path $tempZip -DestinationPath $workingDir -Force
        
        # Find content.opf using the updated function
        $contentOPFPath = Find-ContentFile -WorkingDir $workingDir -FilePattern "*.opf"
        
        if (-not $contentOPFPath) {
            Write-Error "Could not find content.opf in EPUB: $EpubFile"
            return
        }
        
        # Get content directory (parent folder of content.opf)
        $contentDir = Split-Path -Path $contentOPFPath
        
        # Find toc.ncx using the function
        $tocPath = Find-ContentFile -WorkingDir $workingDir -FilePattern "*.ncx"
        
        # Get spine order from content.opf
        $pageOrder = Get-SpineOrder -ContentOPFPath $contentOPFPath
        if (-not $pageOrder) {
            Write-Error "Failed to get page order from content.opf"
            return
        }
        
        # Get page mapping from TOC if available
        $pageMap = @{}
        if ($tocPath) {
            $pageMap = Get-PageMapping -TOCPath $tocPath
            if (-not $pageMap) {
                Write-Warning "Failed to get page mapping from toc.ncx - will continue without page labels"
            }
        }
        
        # Keep track of the cover path for duplicate detection
        $coverImageFullPath = $null
        
        # First, process the cover page if it exists
        $coverProcessed = $false
        
        # Look for cover in multiple ways
        try {
            # 1. Try to find cover image directly from manifest
            [xml]$opfContent = Get-Content -LiteralPath $contentOPFPath -Encoding UTF8
            
            # Look for cover meta element
            $coverItem = $opfContent.package.metadata.meta | Where-Object { $_.name -eq "cover" }
            if ($coverItem -and $coverItem.content) {
                $coverId = $coverItem.content
                $coverManifestItem = $opfContent.package.manifest.item | Where-Object { $_.id -eq $coverId }
                
                if ($coverManifestItem -and $coverManifestItem.href) {
                    $coverHref = $coverManifestItem.href
                    $fullCoverPath = Join-Path -Path $contentDir -ChildPath $coverHref
                    
                    if (Test-Path -LiteralPath $fullCoverPath) {
                        $coverImageFullPath = $fullCoverPath
                        $extension = [System.IO.Path]::GetExtension($fullCoverPath)
                        $newCoverName = "000 - Cover{0}" -f $extension
                        $outputCoverPath = Join-Path -Path $orderedImagesDir -ChildPath $newCoverName
                        
                        # Copy the cover file
                        Copy-Item -LiteralPath $fullCoverPath -Destination $outputCoverPath
                        Write-Host "  - Added: $newCoverName (Cover from manifest)"
                        $coverProcessed = $true
                    }
                }
            }
            
            # 2. If no cover found yet, try guide reference
            if (-not $coverProcessed -and $opfContent.package.guide) {
                $coverRef = $opfContent.package.guide.reference | Where-Object { $_.type -eq "cover" }
                if ($coverRef -and $coverRef.href) {
                    # Find the corresponding XHTML file
                    $coverXhtmlPath = Join-Path -Path $contentDir -ChildPath $coverRef.href
                    
                    if (Test-Path -LiteralPath $coverXhtmlPath) {
                        # Extract image from the XHTML
                        $coverContent = Get-Content -LiteralPath $coverXhtmlPath -Raw
                        if ($coverContent -match '<img\s+[^>]*src=["'']([^"'']+)["''][^>]*>') {
                            $coverImagePath = $matches[1]
                            $fullCoverPath = Join-Path -Path (Split-Path -Path $coverXhtmlPath) -ChildPath $coverImagePath
                            
                            if (Test-Path -LiteralPath $fullCoverPath) {
                                $coverImageFullPath = $fullCoverPath
                                $extension = [System.IO.Path]::GetExtension($fullCoverPath)
                                $newCoverName = "000 - Cover{0}" -f $extension
                                $outputCoverPath = Join-Path -Path $orderedImagesDir -ChildPath $newCoverName
                                
                                # Copy the cover file
                                Copy-Item -LiteralPath $fullCoverPath -Destination $outputCoverPath
                                Write-Host "  - Added: $newCoverName (Cover from guide)"
                                $coverProcessed = $true
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Warning "Error processing cover: $_"
        }
        
        # Process the rest of the pages in order
        $pageNumber = 1
        $firstPageProcessed = $false
        
        # Add a function to extract image URLs from XHTML content
        function Extract-ImageUrls {
            param (
                [string]$HtmlContent
            )
            
            $imageUrls = @()
            
            # Match standard img tags
            $imgMatches = [regex]::Matches($HtmlContent, '<img\s+[^>]*src=["'']([^"'']+)["''][^>]*>')
            foreach ($match in $imgMatches) {
                $imageUrls += $match.Groups[1].Value
            }
            
            # Some EPUBs use image tags in SVG wrappers
            $svgMatches = [regex]::Matches($HtmlContent, '<image\s+[^>]*xlink:href=["'']([^"'']+)["''][^>]*>')
            foreach ($match in $svgMatches) {
                $imageUrls += $match.Groups[1].Value
            }
            
            return $imageUrls
        }
        
        foreach ($page in $pageOrder) {
            # Skip cover page if we already processed it
            if ($page -match "cover\.xhtml$" -or $page -match "PAGE_cover\.xhtml$") {
                continue
            }
            
            # Extract image filename from xhtml
            $xhtmlPath = Join-Path -Path $contentDir -ChildPath $page
            
            if (Test-Path -LiteralPath $xhtmlPath) {
                $xhtmlContent = Get-Content -LiteralPath $xhtmlPath -Raw
                # Extract all image paths from the XHTML file
                $imageUrls = Extract-ImageUrls -HtmlContent $xhtmlContent
                
                foreach ($imagePath in $imageUrls) {
                    # Try multiple approaches for resolving the image path
                    $possiblePaths = @(
                        # Path 1: Relative to XHTML file
                        [System.IO.Path]::GetFullPath([System.IO.Path]::Combine((Split-Path -Path $xhtmlPath), $imagePath))
                        # Path 2: Relative to content directory
                        Join-Path -Path $contentDir -ChildPath $imagePath
                        # Path 3: OEBPS special case - some EPUBs store images in OEBPS/images folder
                        Join-Path -Path $workingDir -ChildPath "OEBPS\images\$([System.IO.Path]::GetFileName($imagePath))"
                        # Path 4: Standard image folder
                        Join-Path -Path $workingDir -ChildPath "images\$([System.IO.Path]::GetFileName($imagePath))"
                        # Path 5: Root level images folder
                        Join-Path -Path $workingDir -ChildPath "$([System.IO.Path]::GetFileName($imagePath))"
                    )
                    
                    $fullImagePath = $null
                    foreach ($path in $possiblePaths) {
                        if (Test-Path -LiteralPath $path) {
                            $fullImagePath = $path
                            break
                        }
                    }
                    
                    if (-not $fullImagePath) {
                        # If still not found, try recursive search as last resort
                        $imageFilename = [System.IO.Path]::GetFileName($imagePath)
                        $imageFiles = Get-ChildItem -Path $workingDir -Filter $imageFilename -File -Recurse -ErrorAction SilentlyContinue
                        
                        if ($imageFiles.Count -gt 0) {
                            $fullImagePath = $imageFiles[0].FullName
                            Write-Host "  Found image by filename search: $fullImagePath"
                        }
                    }
                    
                    if ($fullImagePath) {
                        $foundAnyImages = $true
                        # Check if this is the first page and if it's identical to the cover
                        $isDuplicate = $false
                        
                        if (!$firstPageProcessed -and $coverProcessed -and $coverImageFullPath) {
                            # Compare the files to see if they're the same image
                            if (Compare-FileContent -Path1 $fullImagePath -Path2 $coverImageFullPath) {
                                Write-Host "  - Skipping duplicate first page (identical to cover)"
                                $firstPageProcessed = $true
                                $isDuplicate = $true
                            }
                        }
                        
                        if (-not $isDuplicate) {
                            # Get page label from TOC if available
                            $pageLabel = if ($pageMap.ContainsKey($page)) { $pageMap[$page] } else { $pageNumber }
                            
                            # Determine file extension
                            $extension = [System.IO.Path]::GetExtension($fullImagePath)
                            
                            # Create padded filename
                            $newFilename = "{0:D3} - {1}{2}" -f $pageNumber, $pageLabel, $extension
                            $outputPath = Join-Path -Path $orderedImagesDir -ChildPath $newFilename
                            
                            # Copy the file
                            Copy-Item -LiteralPath $fullImagePath -Destination $outputPath
                            Write-Host "  - Added: $newFilename"
                            
                            $pageNumber++
                            
                            if (-not $firstPageProcessed) {
                                $firstPageProcessed = $true
                            }
                        }
                    } else {
                        Write-Verbose "Could not find image: $imagePath"
                    }
                }
            } else {
                Write-Warning "  XHTML file not found: $xhtmlPath"
            }
        }
        
        # If no images were processed through spine order, try a fallback approach
        if (-not $foundAnyImages) {
            Write-Host "  No images found using spine order, trying fallback image discovery..."
            
            # Search for all image files in the EPUB
            $imageExtensions = @("*.jpg", "*.jpeg", "*.png", "*.gif", "*.webp")
            $allImages = @()
            
            foreach ($ext in $imageExtensions) {
                $allImages += Get-ChildItem -Path $workingDir -Filter $ext -File -Recurse -ErrorAction SilentlyContinue
            }
            
            # Sort images by path to maintain some order
            $allImages = $allImages | Sort-Object -Property FullName
            
            if ($allImages.Count -gt 0) {
                Write-Host "  Found $($allImages.Count) images using direct file search"
                $pageNumber = 1
                
                foreach ($image in $allImages) {
                    # Skip non-content images (like icons, logos, etc.) which are typically small
                    $fileInfo = Get-Item -LiteralPath $image.FullName
                    if ($fileInfo.Length -lt 10KB) {
                        Write-Verbose "Skipping small image (likely non-content): $($image.Name)"
                        continue
                    }
                    
                    # Skip images in META-INF directory
                    if ($image.FullName -match "META-INF") {
                        continue
                    }
                    
                    # Determine file extension
                    $extension = [System.IO.Path]::GetExtension($image.FullName)
                    
                    # Create padded filename
                    $newFilename = "{0:D3} - Page {0}{1}" -f $pageNumber, $extension
                    $outputPath = Join-Path -Path $orderedImagesDir -ChildPath $newFilename
                    
                    # Copy the file
                    Copy-Item -LiteralPath $image.FullName -Destination $outputPath
                    Write-Host "  - Added: $newFilename (fallback method)"
                    
                    $pageNumber++
                    $foundAnyImages = $true
                }
            }
        }
        
        if (-not $coverProcessed) {
            Write-Warning "No cover image found to process."
        }
        
        # Create the CBZ file
        if ((Get-ChildItem -Path $orderedImagesDir | Measure-Object).Count -gt 0) {
            # Get page count for ComicInfo.xml
            $pageCount = (Get-ChildItem -Path $orderedImagesDir -File | Measure-Object).Count
            
            # Create ComicInfo.xml and add it to the ordered images directory
            $comicInfoPath = Join-Path -Path $orderedImagesDir -ChildPath "ComicInfo.xml"
            Create-ComicInfo -ContentOPFPath $contentOPFPath -OutputPath $comicInfoPath -PageCount $pageCount
            
            # Remove existing CBZ if it exists
            if (Test-Path -LiteralPath $cbzFile) {
                Remove-Item -LiteralPath $cbzFile -Force
            }
            
            # Create a temporary ZIP file first
            $tempZipOutput = Join-Path -Path $TEMP_DIR -ChildPath "$filename.zip"
            if (Test-Path -LiteralPath $tempZipOutput) {
                Remove-Item -LiteralPath $tempZipOutput -Force
            }
            
            # Create ZIP file
            Compress-Archive -Path "$orderedImagesDir\*" -DestinationPath $tempZipOutput
            
            # Rename ZIP to CBZ
            Move-Item -Path $tempZipOutput -Destination $cbzFile -Force
            
            Write-Host "Created CBZ: $cbzFile"
        } else {
            Write-Error "No images were processed from $EpubFile"
        }
    }
    catch {
        Write-Error "Error processing $EpubFile`: $_"
    }
    finally {
        # Clean up temporary files for this iteration
        if ($tempZip -and (Test-Path $tempZip)) {
            Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue
        }
        if ($workingDir -and (Test-Path $workingDir)) {
            Remove-Item -Path $workingDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

try {
    foreach ($epub_file in $FilePath) {
        if (-not ($epub_file -match "\.epub$")) {
            Write-Host "Skipping non-epub file: $epub_file" -ForegroundColor Yellow
            continue
        }
        
        Process-EPUB -EpubFile $epub_file
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    # Final cleanup of temp directory
    if (Test-Path $TEMP_DIR) {
        Remove-Item -Path $TEMP_DIR -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    Write-Host "`nProcessing complete."
    if ($Host.Name -eq 'ConsoleHost') {
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}