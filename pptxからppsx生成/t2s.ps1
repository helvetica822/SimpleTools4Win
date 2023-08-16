function Resolve-PowerPointComAssembly() {
    $powerPnt = $null

    try {
        $powerPnt = New-Object -ComObject PowerPoint.Application
    }
    catch {
    }

    return $powerPnt
}

function Convert-PowerPoint2SlideShow( [System.Object]$powerPnt, [string[]]$paths) {
    $presentation = $null
  
    foreach ($p in $paths) {
        try {
            $dirPath = [IO.Path]::GetDirectoryName($p)
            $fileName = [IO.Path]::GetFileNameWithoutExtension($p)
            $srcPath = Join-Path  $dirPath ($fileName + ".pptx")
            $destPath = Join-Path  $dirPath ($fileName + ".ppsx")
            
            Write-Host $srcPath
            Write-Host " >" $destPath "... " -NoNewline -ForegroundColor Blue
            
            $presentation = $powerPnt.Presentations.Open($srcPath, $null, $null, [Microsoft.Office.Core.MsoTriState]::msoFalse)
            $presentation.SaveAs($destPath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsOpenXMLShow)

            Write-Host "done" -ForegroundColor Blue
        }
        catch {
            Write-Host "error" -ForegroundColor Red
        }
        finally {
            if ($null -ne $presentation) {
                $presentation.Close()
            }       
        }
    }
}

Write-Host "######################################## pptx から ppsx 生成 ########################################"
Write-Host "#                                                                                                   #"
Write-Host "#     PowerPoint ファイル(.pptx)を スライドショーファイル(.ppsx) に変換します。                     #"
Write-Host "#                                                                                                   #"
Write-Host "#     動作要件                                                                                      #"
Write-Host "#       ・" -NoNewline
Write-Host "実行端末に PowerPoint がインストールされていること                             " -NoNewline -ForegroundColor Red
Write-Host "           #"
Write-Host "#                                                                                                   #"
Write-Host "#     機能                                                                                          #"
Write-Host "#       ・単一の PowerPoint ファイルへのパスを指定した場合、指定した 1 ファイルのみ変換します。     #"
Write-Host "#       ・フォルダのパスを指定した場合、フォルダ配下の PowerPoint ファイルを全て変換します。        #"
Write-Host "#                                                                                                   #"
Write-Host "#####################################################################################################"
Write-Host

Write-Host "本ツールの動作要件をチェックしています ... "
Write-Host " > PowerPoint インストールチェック ... " -NoNewline
$powerPntApp = Resolve-PowerPointComAssembly

if ($null -eq $powerPntApp) {
    Write-Host "error" -ForegroundColor Red
    exit
}

Write-Host "done" -ForegroundColor Blue
Write-Host

try {
    while ($true) {
        $path = Read-Host "変換対象のパス"

        if ($path.Length -eq 0) {
            continue
        }
    
        if (-not( Test-Path $path)) {
            Write-Host "不正なパスまたは存在しないパスです。"
            continue
        }

        break;
    }

    $path = Convert-Path $path
    
    if ((Get-Item $path).PSIsContainer) {
        $files = Get-ChildItem $path -Include *.pptx -Recurse
    
        if ($files.Count -eq 0) {
            Write-Host "PowerPoint ファイルが存在しません。"
            exit
        }
    
        Write-Host
        Write-Host "スライドショーファイル(.ppsx) 変換対象ファイル"
    
        $files | ForEach-Object { Write-Host " >" $_.FullName -ForegroundColor Green }
    
        Write-Host
        $yesno = Read-Host "実行してよろしいですか? (y/n)"
    
        if ($yesno.ToLower() -ne "y") {
            Write-Host "実行をキャンセルしました。"
            exit
        }
    
        Write-Host
        
        $filesPaths = @()
        $files | ForEach-Object { $filesPaths += $_.FullName }
        
        Convert-PowerPoint2SlideShow  $powerPntApp $filesPaths
    }
    else {
        $extension = (Get-ChildItem $path | ForEach-Object { $_.Extension }).ToLower()
        if ($extension -ne ".pptx") {
            Write-Host "指定したファイルは PowerPoint ファイルではありません。"
            exit
        }
    
        Convert-PowerPoint2SlideShow  $powerPntApp (, $path)
    }    
}
catch {
    Write-Host "PowerPoint がインストールされていないか使用できません。" -ForegroundColor Red
}
finally {
    if ($null -ne $powerPntApp) {
        $powerPntApp.Quit()
        $powerPntApp = $null
    }
    [GC]::Collect()
}
