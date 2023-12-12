Function Start-Nastran {
  [CmdletBinding()]
  param(
    [string[]] $File,
    [switch]$BatchFile,
    [switch]$PrintLog,
    [string[]]$Keywords
  )

  # check if batch file is specified
  if ($BatchFile) {
    $filelist = Get-Content -Path $File | Get-ChildItem
  } else {
    $filelist = Get-ChildItem $File
  }

  $batchstart = Get-Date
  foreach ($deck in $filelist) {
    # PrintLog will print out the log file in realtime
    $jobstart = Get-Date
    If ($PrintLog) {
      # start as a background process
      $job = Start-Job -name $deck.BaseName -ScriptBlock{
        Start-Process $input -ArgumentList $args -NoNewWindow -Wait
        } -ArgumentList $deck, ($Keywords -join ",") -InputObject nastran

      Start-Sleep -Seconds 1  # to allow for log file to be created
      # print out log file
      $log = Get-ChildItem ($deck.DirectoryName + "\" + $deck.BaseName + ".log")

      # wait until analysis has started and then print log until that point
      Do {
        $lineidx = (Select-String $log -Pattern "Nastran started").LineNumber
      } Until ($lineidx -gt 0)
      $startidx = $lineidx
      Get-Content $log -TotalCount $lineidx

      # catch up printout if necessary
      Do {
        $lineidx += 1
        $line = (Get-Content $log)[$lineidx]
      } While ($lineidx -lt (Get-Content $log).Length)

      # track file and output new lines
      $oldline = $line
      Do {
        $finished = (Select-String $log -Pattern "Nastran finished") -ne $null
        $newline = Get-Content $log -Tail 1
        If ($newline -ne $oldline) {
          Write-Host $newline
        }
        $oldline = $newline
      } Until ($finished)

    # run nastran normally without printing out the log file
    } else {
      nastran $deck @Keywords
      $jobend = Get-Date
      $jobspan = New-Timespan -Start $jobstart -End $jobend
    }
    $jobend = Get-Date
    $jobspan = New-Timespan -Start $jobstart -End $jobend

    # check for pass or fail
    $f06 = Get-ChildItem ($deck.DirectoryName + "\" + $deck.BaseName + ".f06")
    $fail = Select-String -Path $f06 -Pattern "FATAL"

    # notify in console
    $icon = $null
    If ($fail) {
      Write-Host "!!! Nastran job " -NoNewLine -ForegroundColor Red
      Write-Host $deck.BaseName -NoNewLine -ForegroundColor DarkYellow
      Write-Host " failed after " -NoNewLine -ForegroundColor Red
      Write-Host $jobspan.ToString("hh\:mm\:ss") -NoNewLine -ForegroundColor DarkYellow
      Write-Host " !!!" -ForegroundColor Red
      $icon = 48 + 4096  # exclamation point
    } else {
      Write-Host "Nastran job " -NoNewLine -ForegroundColor White
      Write-Host $deck.BaseName -NoNewLine -ForegroundColor Yellow
      Write-Host " completed successfully in " -NoNewLine -ForegroundColor White
      Write-Host $jobspan.ToString("hh\:mm\:ss") -ForegroundColor Yellow
      $icon = 4096
    }
  }
  $batchend = Get-Date
  $batchspan = New-Timespan -Start $batchstart -End $batchend

  # notify in console
  If ($filelist.Length -gt 1) {
    Write-Host "Batch job finished in " -NoNewLine
    Write-Host $batchspan.ToString("hh\:mm\:ss") -ForegroundColor Blue
    $notification = 'Batch job finished in ' + $batchspan.ToString("hh\:mm\:ss")
  } else {
    $notification = 'Nastran job finished in ' + $batchspan.ToString("hh\:mm\:ss")
  }

  # notification popup
  $wshell = New-Object -ComObject Wscript.Shell
  $null = $wshell.Popup($notification, 0, "Nastran Finished", $icon)

  <#
  .SYNOPSIS

  Augmented Nastran execution command.

  .DESCRIPTION

  Runs Nastran and will notify user of job completion. Includes options to run
  single files or batches as well as printing out LOG file data real-time.

  .PARAMETER File
  File or list of files to submit to Nastran. Also accepts strings with
  wildcards. Use of wildcards will submit all jobs matching the search pattern
  as a batch run.

  .PARAMETER BatchFile
  If flag is enabled, the `File` parameter will be assumed to be a text file
  containing a list of Nastran decks to run in a batch run.

  .PARAMETER PrintLog
  If flag is enabled, will print out the resulting LOG file to the console
  for real-time monitoring.

  .PARAMETER Keywords
  Nastran keywords to added to the Nastran command. (See QRG)

  .INPUTS

  None.

  .OUTPUTS

  None.
  #>
}

Set-Alias nas Start-Nastran