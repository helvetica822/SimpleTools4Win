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
            # SaveAs��(�o�[�W�����ɂ���Ă͂������̕����ǂ�����)
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


Write-Host "######################################## Office �t�@�C���� PDF �� ########################################"
Write-Host "#                                                                                                        #"
Write-Host "#     Excel �t�@�C��(.xlsx)�܂��� Word �t�@�C��(.docx)�� PDF �ɕϊ����܂��B                              #"
Write-Host "#                                                                                                        #"
Write-Host "#     ����v��                                                                                           #"
Write-Host "#       �E" -NoNewline
Write-Host "���s�[���� Excel ���C���X�g�[������Ă��邱��                             " -NoNewline -ForegroundColor Red
Write-Host "                     #"
Write-Host "#       �E" -NoNewline
Write-Host "���s�[���� Word ���C���X�g�[������Ă��邱��                              " -NoNewline -ForegroundColor Red
Write-Host "                     #"
Write-Host "#                                                                                                        #"
Write-Host "#     �@�\                                                                                               #"
Write-Host "#       �E�P��� Excel, Word �t�@�C���ւ̃p�X���w�肵���ꍇ�A�w�肵�� 1 �t�@�C���̂ݕϊ����܂��B         #"
Write-Host "#       �E�t�H���_�̃p�X���w�肵���ꍇ�A�t�H���_�z���� Excel, Word �t�@�C����S�ĕϊ����܂��B            #"
Write-Host "#                                                                                                        #"
Write-Host "##########################################################################################################"
Write-Host

Write-Host "�{�c�[���̓���v�����`�F�b�N���Ă��܂� ... "
Write-Host " > Excel �C���X�g�[���`�F�b�N ... " -NoNewline
$excelApp = Resolve-ExcelComAssembly

if ($null -eq $excelApp) {
    Write-Host "error" -ForegroundColor Red
    exit
}

Write-Host "done" -ForegroundColor Blue

Write-Host " > Word �C���X�g�[���`�F�b�N ... " -NoNewline
$wordApp = Resolve-WordComAssembly

if ($null -eq $wordApp) {
    Write-Host "error" -ForegroundColor Red
    exit
}

Write-Host "done" -ForegroundColor Blue
Write-Host

try {
    while ($true) {
        $path = Read-Host "�ϊ��Ώۂ̃p�X"

        if ($path.Length -eq 0) {
            continue
        }
    
        if (-not( Test-Path $path)) {
            Write-Host "�s���ȃp�X�܂��͑��݂��Ȃ��p�X�ł��B"
            continue
        }

        break;
    }

    $path = Convert-Path $path
    
    if ((Get-Item $path).PSIsContainer) {
        $files = Get-ChildItem $path -Include *.xlsx, *.docx -Recurse
    
        if ($files.Count -eq 0) {
            Write-Host "Excel, Word �t�@�C�������݂��܂���B"
            exit
        }
    
        Write-Host
        Write-Host "PDF �ϊ��Ώۃt�@�C��"
    
        $files | ForEach-Object { Write-Host " >" $_.FullName -ForegroundColor Green }
    
        Write-Host
        $yesno = Read-Host "���s���Ă�낵���ł���? (y/n)"
    
        if ($yesno.ToLower() -ne "y") {
            Write-Host "���s���L�����Z�����܂����B"
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
            Write-Host "�w�肵���t�@�C���� Excel, Word �t�@�C���ł͂���܂���B"
            exit
        }
    
        Convert-Office2Pdf  $excelApp $wordApp (, $path)
    }    
}
catch {
    Write-Host "Excel, Word ���C���X�g�[������Ă��Ȃ����g�p�ł��܂���B" -ForegroundColor Red
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
