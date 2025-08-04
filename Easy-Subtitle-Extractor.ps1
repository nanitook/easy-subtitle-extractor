
# ====================================================================================================
# EASY SUBTITLE EXTRACTOR - PowerShell Script
# Requirements: FFmpeg (and FFprobe)
# Version: 1.0
# ====================================================================================================

param(
  [Parameter(Mandatory = $false, HelpMessage = "Video file path")]
  [string]$VideoPath,
  
  [Parameter(Mandatory = $false, HelpMessage = "Output directory for subtitles")]
  [string]$OutputDir,
  
  [Parameter(Mandatory = $false, HelpMessage = "Specific stream to extract (e.g.: 0, 1, 2)")]
  [string]$StreamIndex,
  
  [Parameter(Mandatory = $false, HelpMessage = "List streams only, do not extract")]
  [switch]$ListOnly,
  
  [Parameter(Mandatory = $false, HelpMessage = "Extract all streams automatically")]
  [switch]$ExtractAll,
  
  [Parameter(Mandatory = $false, HelpMessage = "Interactive mode (GUI)")]
  [switch]$Interactive,
  
  [Parameter(Mandatory = $false, HelpMessage = "Custom path to ffmpeg.exe")]
  [string]$FFmpegPath,
  
  [Parameter(Mandatory = $false, HelpMessage = "Custom path to ffprobe.exe")]
  [string]$FFprobePath
)

# Color and style configuration
$Host.UI.RawUI.WindowTitle = "Easy Subtitle Extractor - PowerShell"

## Function to write colored text
function Write-ColorText {
  param(
    [string]$Text,
    [string]$Color = "White",
    [switch]$NoNewline
  )
    
  $colorMap = @{
    "Red"     = [ConsoleColor]::Red
    "Green"   = [ConsoleColor]::Green
    "Yellow"  = [ConsoleColor]::Yellow
    "Blue"    = [ConsoleColor]::Blue
    "Cyan"    = [ConsoleColor]::Cyan
    "Magenta" = [ConsoleColor]::Magenta
    "White"   = [ConsoleColor]::White
    "Gray"    = [ConsoleColor]::Gray
  }
    
  if ($NoNewline) {
    Write-Host $Text -ForegroundColor $colorMap[$Color] -NoNewline
  }
  else {
    Write-Host $Text -ForegroundColor $colorMap[$Color]
  }
}

## Function to display the banner
function Show-Banner {
  Clear-Host
  Write-ColorText "╔══════════════════════════════════════════════════════════════════════════╗" "Cyan"
  Write-ColorText "║                        🎬 SUBTITLE EXTRACTOR 🎬                          ║" "Cyan"
  Write-ColorText "║                         PowerShell Edition v1.0                          ║" "Cyan"
  Write-ColorText "╚══════════════════════════════════════════════════════════════════════════╝" "Cyan"
  Write-ColorText ""
}

## Function to check if FFmpeg is installed
function Test-FFmpegInstallation {
  Write-ColorText "🔍 Checking FFmpeg installation..." "Yellow"
  $ffmpegCmd = if ($FFmpegPath) { $FFmpegPath } else { "ffmpeg" }
  try {
    $ffmpegVersion = & $ffmpegCmd -version 2>$null | Select-Object -First 1
    if ($ffmpegVersion -match "ffmpeg version") {
      Write-ColorText "✅ FFmpeg found: $($ffmpegVersion -replace 'ffmpeg version ', '')" "Green"
      return $true
    }
  }
  catch {
    Write-ColorText "❌ Error: FFmpeg is not installed or not in PATH" "Red"
    Write-ColorText ""
    Write-ColorText "To install FFmpeg:" "Yellow"
    Write-ColorText "1. Download from: https://ffmpeg.org/download.html" "White"
    Write-ColorText "2. Or install with chocolatey: choco install ffmpeg" "White"
    Write-ColorText "3. Or install with winget: winget install ffmpeg" "White"
    return $false
  }
}

## Function to validate video file
function Test-VideoFile {
  param([string]$FilePath)
    
  if (-not (Test-Path $FilePath)) {
    Write-ColorText "❌ Error: File does not exist: $FilePath" "Red"
    return $false
  }
    
  $validExtensions = @('.mp4', '.mkv', '.avi', '.mov', '.wmv', '.flv', '.webm', '.m4v', '.3gp', '.ts', '.m2ts', '.mpg', '.mpeg', '.ogv')
  $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    
  if ($extension -notin $validExtensions) {
    Write-ColorText "❌ Error: Unsupported file format: $extension" "Red"
    Write-ColorText "Supported formats: $($validExtensions -join ', ')" "Gray"
    return $false
  }
    
  $fileSize = (Get-Item $FilePath).Length
  Write-ColorText "📁 File: $(Split-Path $FilePath -Leaf)" "Green"
  Write-ColorText "📊 Size: $([math]::Round($fileSize / 1MB, 2)) MB" "Green"
    
  return $true
}

## Function to list subtitle streams
function Get-SubtitleStreams {
  param([string]$VideoPath)
  Write-ColorText "🔍 Analyzing subtitle streams..." "Yellow"
  $ffprobeCmd = if ($FFprobePath) { $FFprobePath } else { "ffprobe" }
  $ffmpegCmd = if ($FFmpegPath) { $FFmpegPath } else { "ffmpeg" }
  try {
    # Run ffprobe to get detailed stream information
    $ffprobeArgs = @(
      '-v', 'quiet'
      '-select_streams', 's'
      '-show_entries', 'stream=index,codec_name,codec_type,codec_tag_string:stream_tags=language,title'
      '-of', 'csv=p=0'
      $VideoPath
    )
    $streamInfo = & $ffprobeCmd @ffprobeArgs 2>$null
    # If ffprobe fails, use alternative method with ffmpeg
    if (-not $streamInfo) {
      Write-ColorText "📋 Using alternative detection method..." "Yellow"
      $ffmpegOutput = & $ffmpegCmd -i $VideoPath -f null - 2>&1 | Out-String
      $streamInfo = @()
      # Parse ffmpeg output to find subtitle streams
      $ffmpegOutput -split "`n" | ForEach-Object {
        if ($_ -match "Stream #(\d+):(\d+).*?: Subtitle: (.+)" -or 
          $_ -match "Stream #(\d+):(\d+).*?subtitle.*?: (.+)") {
          $streamInfo += "$($matches[2]),$($matches[3]),subtitle,"
        }
      }
    }
    if (-not $streamInfo) {
      Write-ColorText "ℹ️  No subtitle streams found using automatic detection." "Yellow"
      Write-ColorText "🔍 Trying manual detection..." "Yellow"
      # Manual method: try subtitle indexes
      $manualStreams = @()
      for ($i = 0; $i -lt 10; $i++) {
        try {
          $testResult = & $ffmpegCmd -i $VideoPath -map "0:s:$i" -t 0.1 -f null - 2>&1
          if ($LASTEXITCODE -eq 0) {
            $manualStreams += [PSCustomObject]@{
              Index    = $i
              Codec    = "subtitle"
              Language = "unknown"
              Title    = "Subtitle stream $i (auto-detected)"
            }
          }
        }
        catch {
          break
        }
      }
      return $manualStreams
    }
    # Process stream information
    $streams = @()
    $streamInfo | ForEach-Object {
      if ($_ -and $_.Trim()) {
        $parts = $_ -split ','
        if ($parts.Length -ge 3) {
          $streams += [PSCustomObject]@{
            Index    = $parts[0]
            Codec    = if ($parts[1]) { $parts[1] } else { "subtitle" }
            Language = if ($parts.Length -gt 4 -and $parts[4]) { $parts[4] } else { "unknown" }
            Title    = if ($parts.Length -gt 5 -and $parts[5]) { $parts[5] } else { "Subtitle stream $($parts[0])" }
          }
        }
      }
    }
    return $streams
  }
  catch {
    Write-ColorText "❌ Error analyzing streams: $($_.Exception.Message)" "Red"
    return @()
  }
}

## Function to display found streams
function Show-SubtitleStreams {
  param([array]$Streams)
    
  if ($Streams.Count -eq 0) {
    Write-ColorText "❌ No subtitle streams found in the file." "Red"
    Write-ColorText ""
    Write-ColorText "Possible causes:" "Yellow"
    Write-ColorText "• The file does not contain embedded subtitles" "Gray"
    Write-ColorText "• The subtitles are in an unsupported format" "Gray"
    Write-ColorText "• The file is corrupted" "Gray"
    return
  }
    
  Write-ColorText ""
  Write-ColorText "📋 SUBTITLE STREAMS FOUND:" "Green"
  Write-ColorText "═══════════════════════════════════════" "Green"
    
  for ($i = 0; $i -lt $Streams.Count; $i++) {
    $stream = $Streams[$i]
    Write-ColorText "[$($i + 1)] " "Cyan" -NoNewline
    Write-ColorText "Stream $($stream.Index) " "White" -NoNewline
    Write-ColorText "│ " "Gray" -NoNewline
    Write-ColorText "$($stream.Codec) " "Yellow" -NoNewline
    Write-ColorText "│ " "Gray" -NoNewline
    Write-ColorText "$($stream.Language) " "Magenta" -NoNewline
    Write-ColorText "│ " "Gray" -NoNewline
    Write-ColorText "$($stream.Title)" "White"
  }
  Write-ColorText ""
}

## Function to extract a specific subtitle
function Export-SubtitleStream {
  param(
    [string]$VideoPath,
    [object]$Stream,
    [string]$OutputDir,
    [int]$StreamNumber
  )
  $ffmpegCmd = if ($FFmpegPath) { $FFmpegPath } else { "ffmpeg" }
  $videoFileName = [System.IO.Path]::GetFileNameWithoutExtension($VideoPath)
  $language = if ($Stream.Language -and $Stream.Language -ne "unknown") { "_$($Stream.Language)" } else { "" }
  $outputFileName = "${videoFileName}_subtitle_${StreamNumber}${language}.srt"
  $outputPath = Join-Path $OutputDir $outputFileName
  Write-ColorText "════════════════════════════════════════════════════" "Cyan"
  Write-ColorText "⏳ Extracting stream $StreamNumber (real index: $($Stream.Index))..." "Yellow"
  Write-ColorText "📝 Output file: $outputFileName" "Gray"
  # Multiple extraction strategies
  $extractionCommands = @(
    # Main command: convert to SRT
    @('-i', $VideoPath, '-map', "0:s:$($Stream.Index)", '-c:s', 'srt', '-avoid_negative_ts', 'make_zero', $outputPath),
    @('-i', $VideoPath, '-map', "0:$($Stream.Index)", '-c:s', 'srt', '-avoid_negative_ts', 'make_zero', $outputPath),
    # Alternative command for ASS/SSA
    @('-i', $VideoPath, '-map', "0:s:$($Stream.Index)", '-f', 'srt', $outputPath),
    # Generic alternative command
    @('-i', $VideoPath, '-map', "0:$($Stream.Index)", '-c:s', 'copy', $outputPath)
  )
  $extractionSuccessful = $false
  foreach ($command in $extractionCommands) {
    try {
      Write-ColorText "⚙️  Running FFmpeg..." "Gray"
      & $ffmpegCmd @command 2>&1 | Out-Null
      if ($LASTEXITCODE -eq 0 -and (Test-Path $outputPath)) {
        $fileSize = (Get-Item $outputPath).Length
        if ($fileSize -gt 0) {
          Write-ColorText "✅ Subtitle extracted successfully!" "Green"
          Write-ColorText "📁 Location: $outputPath" "Green"
          Write-ColorText "📊 Size: $([math]::Round($fileSize / 1KB, 2)) KB" "Green"
          $extractionSuccessful = $true
          break
        }
      }
    }
    catch {
      continue
    }
  }
  if (-not $extractionSuccessful) {
    Write-ColorText "❌ Error: Could not extract stream $($Stream.Index)" "Red"
    Write-ColorText "Possible causes:" "Yellow"
    Write-ColorText "• The stream does not contain text subtitles" "Gray"
    Write-ColorText "• The subtitle format is not supported" "Gray"
    Write-ColorText "• The subtitles are images (PGS/VobSub)" "Gray"
  }
  return $extractionSuccessful
}

## Function to select file interactively
function Select-VideoFile {
  Add-Type -AssemblyName System.Windows.Forms
    
  $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $openFileDialog.Title = "Select video file"
  $openFileDialog.Filter = "Video Files|*.mp4;*.mkv;*.avi;*.mov;*.wmv;*.flv;*.webm;*.m4v;*.3gp;*.ts;*.m2ts;*.mpg;*.mpeg;*.ogv|All files|*.*"
  $openFileDialog.Multiselect = $false
    
  if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    return $openFileDialog.FileName
  }
    
  return $null
}

## Function to select output directory
function Select-OutputDirectory {
  Add-Type -AssemblyName System.Windows.Forms
    
  $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
  $folderDialog.Description = "Select output directory for subtitles"
  $folderDialog.ShowNewFolderButton = $true
    
  if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    return $folderDialog.SelectedPath
  }
    
  return $null
}

## Main interactive function
function Start-InteractiveMode {
  do {
    Show-Banner

    # Check FFmpeg, if it fails and there are no custom paths, prompt the user
    while (-not (Test-FFmpegInstallation)) {
      Write-ColorText ""
      Write-ColorText "Do you want to specify the path to ffmpeg.exe and ffprobe.exe manually? (Y/N): " "Yellow" -NoNewline
      $setPath = Read-Host
      if ($setPath.ToUpper() -eq "Y") {
        Write-ColorText "Enter the full path to ffmpeg.exe: " "Yellow" -NoNewline
        $script:FFmpegPath = Read-Host
        Write-ColorText "Enter the full path to ffprobe.exe: " "Yellow" -NoNewline
        $script:FFprobePath = Read-Host
      }
      else {
        Write-ColorText "Press any key to exit..." "Gray"
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
      }
    }
    Write-ColorText ""
    Write-ColorText "🎯 INTERACTIVE MODE" "Blue"
    Write-ColorText "══════════════════" "Blue"
    Write-ColorText ""
    Write-ColorText "[1] Select video file" "White"
    Write-ColorText "[2] Enter path manually" "White"
    Write-ColorText "[3] Exit" "White"
    Write-ColorText ""
    Write-ColorText "Choose an option: " "Yellow" -NoNewline
    $choice = Read-Host
    $videoPath = $null
    switch ($choice) {
      "1" {
        $videoPath = Select-VideoFile
        if (-not $videoPath) {
          Write-ColorText "❌ No file selected." "Red"
          Start-Sleep 2
          continue
        }
      }
      "2" {
        Write-ColorText "Enter the video file path: " "Yellow" -NoNewline
        $videoPath = Read-Host
        if (-not $videoPath) {
          Write-ColorText "❌ Invalid path." "Red"
          Start-Sleep 2
          continue
        }
      }
      "3" {
        Write-ColorText "👋 Goodbye!" "Cyan"
        return
      }
      default {
        Write-ColorText "❌ Invalid option." "Red"
        Start-Sleep 2
        continue
      }
    }
    # Validate file
    if (-not (Test-VideoFile $videoPath)) {
      Write-ColorText "Press any key to continue..." "Gray"
      $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
      continue
    }
    # Get streams
    $streams = Get-SubtitleStreams $videoPath
    Show-SubtitleStreams $streams
    if ($streams.Count -eq 0) {
      Write-ColorText "Press any key to continue..." "Gray"
      $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
      continue
    }
    # Extraction menu
    Write-ColorText "🎯 EXTRACTION OPTIONS:" "Blue"
    Write-ColorText "═════════════════════════" "Blue"
    Write-ColorText "[A] Extract all streams" "White"
    Write-ColorText "[1-$($streams.Count)] Extract specific stream" "White"
    Write-ColorText "[R] Return to main menu" "White"
    Write-ColorText ""
    Write-ColorText "Choose an option: " "Yellow" -NoNewline
    $extractChoice = Read-Host
    if ($extractChoice.ToUpper() -eq "R") {
      continue
    }
    # Select output directory
    $outputDir = Select-OutputDirectory
    if (-not $outputDir) {
      $outputDir = Split-Path $videoPath -Parent
      Write-ColorText "📁 Using video directory: $outputDir" "Yellow"
    }
    # Extraction process
    if ($extractChoice.ToUpper() -eq "A") {
      Write-ColorText ""
      Write-ColorText "🚀 Extracting all streams..." "Green"
      Write-ColorText "═══════════════════════════════════" "Green"
      $successCount = 0
      for ($i = 0; $i -lt $streams.Count; $i++) {
        Write-ColorText ""
        if (Export-SubtitleStream $videoPath $streams[$i] $outputDir ($i + 1)) {
          $successCount++
        }
      }
      Write-ColorText ""
      Write-ColorText "📊 SUMMARY: $successCount of $($streams.Count) streams extracted successfully." "Green"
    }
    elseif ($extractChoice -match '^[\d]+$' -and [int]$extractChoice -ge 1 -and [int]$extractChoice -le $streams.Count) {
      $selectedIndex = [int]$extractChoice - 1
      Write-ColorText ""
      Export-SubtitleStream $videoPath $streams[$selectedIndex] $outputDir $extractChoice
    }
    else {
      Write-ColorText "❌ Invalid option." "Red"
    }
    Write-ColorText ""
    Write-ColorText "Press any key to return to the main menu..." "Gray"
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
  } while ($true)
}

## ====================================================================================================
## MAIN FUNCTION
## ====================================================================================================

function Main {
  # Check if running in interactive mode
  if ($Interactive) {
    Start-InteractiveMode
    return
  }
    
  Show-Banner
    
  # Check FFmpeg
  if (-not (Test-FFmpegInstallation)) {
    exit 1
  }
    
  # If no file is provided, prompt for one
  if (-not $VideoPath) {
    Write-ColorText "Enter the video file path: " "Yellow" -NoNewline
    $VideoPath = Read-Host
  }
    
  # Validate file
  if (-not (Test-VideoFile $VideoPath)) {
    exit 1
  }
    
  # Set output directory
  if (-not $OutputDir) {
    $OutputDir = Split-Path $VideoPath -Parent
  }
    
  if (-not (Test-Path $OutputDir)) {
    try {
      New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
    catch {
      Write-ColorText "❌ Error: Could not create output directory: $OutputDir" "Red"
      exit 1
    }
  }
    
  # Get streams
  $streams = Get-SubtitleStreams $VideoPath
  Show-SubtitleStreams $streams
    
  if ($streams.Count -eq 0) {
    exit 1
  }
    
  # Only list streams
  if ($ListOnly) {
    Write-ColorText "✅ Listing completed." "Green"
    return
  }
    
  # Extract all streams
  if ($ExtractAll) {
    Write-ColorText ""
    Write-ColorText "🚀 Extracting all streams..." "Green"
    Write-ColorText "═══════════════════════════════════" "Green"
    $successCount = 0
    for ($i = 0; $i -lt $streams.Count; $i++) {
      Write-ColorText ""
      if (Export-SubtitleStream $VideoPath $streams[$i] $OutputDir ($i + 1)) {
        $successCount++
      }
    }
    Write-ColorText ""
    Write-ColorText "📊 SUMMARY: $successCount of $($streams.Count) streams extracted successfully." "Green"
    return
  }
    
  # Extract specific stream
  if ($StreamIndex) {
    if ($StreamIndex -match '^\d+$' -and [int]$StreamIndex -ge 1 -and [int]$StreamIndex -le $streams.Count) {
      $selectedIndex = [int]$StreamIndex - 1
      Write-ColorText ""
      Export-SubtitleStream $VideoPath $streams[$selectedIndex] $OutputDir $StreamIndex
    }
    else {
      Write-ColorText "❌ Error: Invalid stream index. Use a number between 1 and $($streams.Count)" "Red"
      exit 1
    }
    return
  }
    
  # Default to interactive mode
  Start-InteractiveMode
}

# Run main function in interactive mode by default if no parameters are passed directly
if ($MyInvocation.BoundParameters.Count -eq 0) {
  $script:Interactive = $true
  Main
}
# Run main function
else {
  Main
}