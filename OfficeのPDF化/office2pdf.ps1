function Resolve-ExcelComAssembly() {
    $excel = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
    }
    catch {      
        if ($null -ne $excel) {
            $excel.Quit()
            $excel = $null
            [GC]::Collect()
        }
    }

    return $excel
}

function Resolve-WordComAssembly() {
    $word = $null

    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
    }
    catch {      
        if ($null -ne $word) {
            $word.Quit()
            $word = $null
            [GC]::Collect()
        }
    }

    return $word
}

function Convert-Excel2Pdf( [System.Object]$excel, [string[]]$paths) {
    $book = $null

    try {
        foreach ($p in $paths) {
            $dirPath = [IO.Path]::GetDirectoryName($p)
            $fileName = [IO.Path]::GetFileNameWithoutExtension($p)
            $srcPath = Join-Path  $dirPath ($fileName + ".xlsx")
            $destPath = Join-Path  $dirPath ($fileName + ".pdf")
            
            Write-Host $srcPath
            Write-Host " >" $destPath "... " -NoNewline -ForegroundColor Blue

            $book = $excel.Workbooks.Open($srcPath)
            $book.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $destPath)

            Write-Host "done" -ForegroundColor Blue
        }
    }
    catch {
        Write-Host "error" -ForegroundColor Red
    }
    finally {
        if ($null -ne $book) {
            $book.Close($false)
        }
    }
}

function Convert-Word2Pdf( [System.Object]$word, [string[]]$paths) {
    $doc = $null

    try {
        foreach ($p in $paths) {
            $dirPath = [IO.Path]::GetDirectoryName($p)
            $fileName = [IO.Path]::GetFileNameWithoutExtension($p)
            $srcPath = Join-Path  $dirPath ($fileName + ".docx")
            $destPath = Join-Path  $dirPath ($fileName + ".pdf")
            
            Write-Host $srcPath
            Write-Host " >" $destPath "... " -NoNewline -ForegroundColor Blue

            $doc = $word.Documents.Open($srcPath)
            $doc.ExportAsFixedFormat($destPath, [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF)
            # SaveAs版(バージョンによってはこっちの方が良いかも)
            # $doc.SaveAs([ref] $destPath, [ref] 17)
            
            Write-Host "done" -ForegroundColor Blue
        }
    }
    catch {
        Write-Host "error" -ForegroundColor Red
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close($false)
        }
    }
}

function Convert-Office2Pdf( [System.Object]$excel, [System.Object]$word, [string[]]$paths) {
    $xlsxFiles = @()
    $docxFiles = @()

    foreach ($p in $paths) {
        $extension = [System.IO.Path]::GetExtension($p)
        if ($extension -eq ".xlsx") {
            $xlsxFiles += $p
        }
        elseif ($extension -eq ".docx") {
            $docxFiles += $p
        }
    }

    Convert-Excel2Pdf $excel $xlsxFiles
    Convert-Word2Pdf $word $docxFiles
}


Write-Host "######################################## Office ファイルの PDF 化 ########################################"
Write-Host "#                                                                                                        #"
Write-Host "#     Excel ファイル(.xlsx)または Word ファイル(.docx)を PDF に変換します。                              #"
Write-Host "#                                                                                                        #"
Write-Host "#     動作要件                                                                                           #"
Write-Host "#       ・" -NoNewline
Write-Host "実行端末に Excel がインストールされていること                             " -NoNewline -ForegroundColor Red
Write-Host "                     #"
Write-Host "#       ・" -NoNewline
Write-Host "実行端末に Word がインストールされていること                              " -NoNewline -ForegroundColor Red
Write-Host "                     #"
Write-Host "#                                                                                                        #"
Write-Host "#     機能                                                                                               #"
Write-Host "#       ・単一の Excel, Word ファイルへのパスを指定した場合、指定した 1 ファイルのみ変換します。         #"
Write-Host "#       ・フォルダのパスを指定した場合、フォルダ配下の Excel, Word ファイルを全て変換します。            #"
Write-Host "#                                                                                                        #"
Write-Host "##########################################################################################################"
Write-Host

Write-Host "本ツールの動作要件をチェックしています ... "
Write-Host " > Excel インストールチェック ... " -NoNewline
$excelApp = Resolve-ExcelComAssembly

if ($null -eq $excelApp) {
    Write-Host "error" -ForegroundColor Red
    exit
}

Write-Host "done" -ForegroundColor Blue

Write-Host " > Word インストールチェック ... " -NoNewline
$wordApp = Resolve-WordComAssembly

if ($null -eq $wordApp) {
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
        $files = Get-ChildItem $path -Include *.xlsx, *.docx -Recurse
    
        if ($files.Count -eq 0) {
            Write-Host "Excel, Word ファイルが存在しません。"
            exit
        }
    
        Write-Host
        Write-Host "PDF 変換対象ファイル"
    
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
        
        Convert-Office2Pdf  $excelApp $wordApp $filesPaths
    }
    else {
        $extension = (Get-ChildItem $path | ForEach-Object { $_.Extension }).ToLower()
        if ($extension -ne ".xlsx" -and $extension -ne ".docx") {
            Write-Host "指定したファイルは Excel, Word ファイルではありません。"
            exit
        }
    
        Convert-Office2Pdf  $excelApp $wordApp (, $path)
    }    
}
catch {
    Write-Host "Excel, Word がインストールされていないか使用できません。" -ForegroundColor Red
}
finally {
    if ($null -eq $excelApp) {
        $excelApp.Quit()
        $excelApp = $null
    }
    if ($null -eq $wordApp) {
        $wordApp.Quit()
        $wordApp = $null
    }
    [GC]::Collect()
}
