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

Write-Host "######################################## pptx ���� ppsx ���� ########################################"
Write-Host "#                                                                                                   #"
Write-Host "#     PowerPoint �t�@�C��(.pptx)�� �X���C�h�V���[�t�@�C��(.ppsx) �ɕϊ����܂��B                     #"
Write-Host "#                                                                                                   #"
Write-Host "#     ����v��                                                                                      #"
Write-Host "#       �E" -NoNewline
Write-Host "���s�[���� PowerPoint ���C���X�g�[������Ă��邱��                             " -NoNewline -ForegroundColor Red
Write-Host "           #"
Write-Host "#                                                                                                   #"
Write-Host "#     �@�\                                                                                          #"
Write-Host "#       �E�P��� PowerPoint �t�@�C���ւ̃p�X���w�肵���ꍇ�A�w�肵�� 1 �t�@�C���̂ݕϊ����܂��B     #"
Write-Host "#       �E�t�H���_�̃p�X���w�肵���ꍇ�A�t�H���_�z���� PowerPoint �t�@�C����S�ĕϊ����܂��B        #"
Write-Host "#                                                                                                   #"
Write-Host "#####################################################################################################"
Write-Host

Write-Host "�{�c�[���̓���v�����`�F�b�N���Ă��܂� ... "
Write-Host " > PowerPoint �C���X�g�[���`�F�b�N ... " -NoNewline
$powerPntApp = Resolve-PowerPointComAssembly

if ($null -eq $powerPntApp) {
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
        $files = Get-ChildItem $path -Include *.pptx -Recurse
    
        if ($files.Count -eq 0) {
            Write-Host "PowerPoint �t�@�C�������݂��܂���B"
            exit
        }
    
        Write-Host
        Write-Host "�X���C�h�V���[�t�@�C��(.ppsx) �ϊ��Ώۃt�@�C��"
    
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
        
        Convert-PowerPoint2SlideShow  $powerPntApp $filesPaths
    }
    else {
        $extension = (Get-ChildItem $path | ForEach-Object { $_.Extension }).ToLower()
        if ($extension -ne ".pptx") {
            Write-Host "�w�肵���t�@�C���� PowerPoint �t�@�C���ł͂���܂���B"
            exit
        }
    
        Convert-PowerPoint2SlideShow  $powerPntApp (, $path)
    }    
}
catch {
    Write-Host "PowerPoint ���C���X�g�[������Ă��Ȃ����g�p�ł��܂���B" -ForegroundColor Red
}
finally {
    if ($null -ne $powerPntApp) {
        $powerPntApp.Quit()
        $powerPntApp = $null
    }
    [GC]::Collect()
}
