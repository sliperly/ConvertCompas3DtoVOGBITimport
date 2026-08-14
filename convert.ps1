# =================================================================
#   КОНВЕРТЕР: Состав изделия -> Файлы импорта Vogbit
#   Использование: запускать через start.bat (двойной клик)
# =================================================================

# ======================== КОНФИГУРАЦИЯ ==========================
$SRC_MASK    = "*Состав*"
$WFILES_MASK = "*файлами*"
$WTECH_MASK  = "*технолог*"
$WNOM_MASK   = "*номенклатур*"

$EXT_TO_KIND = [ordered]@{
    ".pdf" = "Чертёж"
    ".dxf" = "развёртка"
    ".jpg" = "Внешний вид"
}
# =================================================================

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Определяем папку скрипта надёжным способом
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}
Write-Host "Папка скрипта: $ScriptDir"

function Resolve-Single([string]$mask) {
    $found = @(Get-ChildItem $ScriptDir -File | Where-Object { $_.Name -like $mask })
    if ($found.Count -eq 0) { throw "Не найден файл по маске '$mask' в папке: $ScriptDir" }
    if ($found.Count -gt 1) { Write-Warning "По маске '$mask' найдено несколько файлов, берём: $($found[0].Name)" }
    return $found[0].FullName
}

function Get-MatUnit([string]$material, [string]$section) {
    if ($material -match "Труба") { return "м"  }
    if ($material -match "Лист")  { return "кг" }
    return "шт"
}

function Get-LinkedFiles([string]$designation) {
    $result = [System.Collections.Generic.List[hashtable]]::new()
    if ([string]::IsNullOrWhiteSpace($designation)) { return $result }
    $found = Get-ChildItem $ScriptDir -Recurse -File -ErrorAction SilentlyContinue |
             Where-Object { $_.BaseName -like "$designation*" }
    foreach ($f in $found) {
        $ext = $f.Extension.ToLower()
        if ($EXT_TO_KIND.Contains($ext)) {
            $result.Add(@{ Kind = $EXT_TO_KIND[$ext]; FileName = $f.Name })
        }
    }
    return $result
}

function Clear-DataRows($ws) {
    try {
        $last = $ws.UsedRange.Rows.Count
        Write-Host "    Очистка строк 2..$last"
        if ($last -ge 2) {
            $ws.Rows("2:$last").Delete() | Out-Null
        }
    } catch {
        Write-Warning "Clear-DataRows: $($_.Exception.Message)"
    }
}

# ================================================================
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Конвертер: Состав изделия -> Vogbit"          -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

try {
    $pathSrc   = Resolve-Single $SRC_MASK
    $pathFiles = Resolve-Single $WFILES_MASK
    $pathTech  = Resolve-Single $WTECH_MASK
    $pathNom   = Resolve-Single $WNOM_MASK

    Write-Host "Источник  : $(Split-Path -Leaf $pathSrc)"
    Write-Host "С файлами : $(Split-Path -Leaf $pathFiles)"
    Write-Host "С технол. : $(Split-Path -Leaf $pathTech)"
    Write-Host "Номенклат.: $(Split-Path -Leaf $pathNom)"
    Write-Host ""

    Write-Host "Запускаем Excel..." -ForegroundColor Yellow
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible        = $false
    $xl.DisplayAlerts  = $false
    $xl.AskToUpdateLinks = $false
    Write-Host "Excel запущен." -ForegroundColor Green

    try {

        # ============================================================
        # ШАГ 1: Читаем источник
        # Колонки: 1=N 2=Наименование 3=Обозначение 4=Количество
        #          5=Материал 6=Масса 7=Позиция 8=Раздел спецификации
        # ============================================================
        Write-Host "Читаем источник..." -ForegroundColor Yellow
        $wbSrc  = $xl.Workbooks.Open($pathSrc, 0, $true)  # открыть read-only
        $wsSrc  = $wbSrc.Sheets.Item(1)
        $lastR  = $wsSrc.UsedRange.Rows.Count
        Write-Host "  Строк в источнике: $lastR"

        $items = [System.Collections.Generic.List[PSObject]]::new()

        for ($r = 2; $r -le $lastR; $r++) {
            $section  = [string]$wsSrc.Cells.Item($r, 8).Value2
            $material = [string]$wsSrc.Cells.Item($r, 5).Value2
            $section  = $section.Trim()
            $material = $material.Trim()

            if ($section -notin @("Детали","Стандартные изделия")) { continue }
            if ($material -match "ЛДСП") { continue }

            $desig = [string]$wsSrc.Cells.Item($r, 3).Value2
            $name  = [string]$wsSrc.Cells.Item($r, 2).Value2
            $qty   = $wsSrc.Cells.Item($r, 4).Value2

            $items.Add([PSCustomObject]@{
                Designation = $desig.Trim()
                Name        = $name.Trim()
                Qty         = $qty
                Material    = $material
                Section     = $section
                MatUnit     = (Get-MatUnit $material $section)
            })
        }

        $wbSrc.Close($false)
        Write-Host "  Загружено позиций: $($items.Count)" -ForegroundColor Green
        Write-Host ""

        # ============================================================
        # ШАГ 2: Только номенклатура
        # A=Обозначение B=Наименование C=Кол-во D=Ед.изм E=L F=b
        # ============================================================
        Write-Host "Заполняем 'только номенклатура'..." -ForegroundColor Yellow
        $wbNom = $xl.Workbooks.Open($pathNom)
        $wsNom = $wbNom.Sheets.Item(1)
        Clear-DataRows $wsNom

        $r = 2
        foreach ($item in $items) {
            $wsNom.Cells.Item($r,1) = $item.Designation
            $wsNom.Cells.Item($r,2) = $item.Name
            $wsNom.Cells.Item($r,3) = $item.Qty
            $wsNom.Cells.Item($r,4) = "шт"
            $r++
        }
        $wbNom.Save()
        $wbNom.Close($false)
        Write-Host "  Записано строк: $($r-2)" -ForegroundColor Green

        # ============================================================
        # ШАГ 3: С технологией
        # A=Уровень B=Вид связи C=Тип D=Обозначение E=Наименование
        # F=Количество G=Ед.изм H=Тшт,час
        # ============================================================
        Write-Host "Заполняем 'с технологией'..." -ForegroundColor Yellow
        $wbTech = $xl.Workbooks.Open($pathTech)
        $wsTech = $wbTech.Sheets.Item(1)
        Clear-DataRows $wsTech

        $r = 2
        foreach ($item in $items) {
            $wsTech.Cells.Item($r,1) = 0
            $wsTech.Cells.Item($r,2) = "Деталь"
            $wsTech.Cells.Item($r,3) = ""
            $wsTech.Cells.Item($r,4) = $item.Designation
            $wsTech.Cells.Item($r,5) = $item.Name
            $wsTech.Cells.Item($r,6) = $item.Qty
            $wsTech.Cells.Item($r,7) = "шт"
            $wsTech.Cells.Item($r,8) = ""
            $r++

            if ($item.Section -eq "Детали" -and $item.Material -ne "") {
                $wsTech.Cells.Item($r,1) = 1
                $wsTech.Cells.Item($r,2) = "Материал"
                $wsTech.Cells.Item($r,3) = "T"
                $wsTech.Cells.Item($r,4) = ""
                $wsTech.Cells.Item($r,5) = $item.Material
                $wsTech.Cells.Item($r,6) = ""
                $wsTech.Cells.Item($r,7) = $item.MatUnit
                $wsTech.Cells.Item($r,8) = ""
                $r++
            }
        }
        $wbTech.Save()
        $wbTech.Close($false)
        Write-Host "  Записано строк: $($r-2)" -ForegroundColor Green

        # ============================================================
        # ШАГ 4: С файлами
        # A=уровень B=Тип C=вид связи D=обозначение E=наименование
        # F=количество G=ед.изм H=вид связи файла I=файл J=... M=...
        # ============================================================
        Write-Host "Заполняем 'с файлами'..." -ForegroundColor Yellow
        $wbFls = $xl.Workbooks.Open($pathFiles)
        $wsFls = $wbFls.Sheets.Item(1)
        Clear-DataRows $wsFls

        $r = 2
        foreach ($item in $items) {
            $linked = Get-LinkedFiles $item.Designation

            $wsFls.Cells.Item($r,1) = 0
            $wsFls.Cells.Item($r,2) = ""
            $wsFls.Cells.Item($r,3) = "деталь"
            $wsFls.Cells.Item($r,4) = $item.Designation
            $wsFls.Cells.Item($r,5) = $item.Name
            $wsFls.Cells.Item($r,6) = $item.Qty
            $wsFls.Cells.Item($r,7) = "шт"

            $col = 8
            foreach ($lf in $linked) {
                if ($col -gt 12) { break }
                $wsFls.Cells.Item($r,$col)   = $lf.Kind
                $wsFls.Cells.Item($r,$col+1) = $lf.FileName
                $col += 2
            }
            $r++

            if ($item.Section -eq "Детали" -and $item.Material -ne "") {
                $wsFls.Cells.Item($r,1) = 1
                $wsFls.Cells.Item($r,2) = "T"
                $wsFls.Cells.Item($r,3) = "материал"
                $wsFls.Cells.Item($r,4) = ""
                $wsFls.Cells.Item($r,5) = $item.Material
                $wsFls.Cells.Item($r,6) = ""
                $wsFls.Cells.Item($r,7) = $item.MatUnit
                $r++
            }
        }
        $wbFls.Save()
        $wbFls.Close($false)
        Write-Host "  Записано строк: $($r-2)" -ForegroundColor Green

    } finally {
        Write-Host "Закрываем Excel..." -ForegroundColor Yellow
        $xl.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
        [System.GC]::Collect()
        Write-Host "Excel закрыт." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "================================================" -ForegroundColor Green
    Write-Host "  ГОТОВО. Три файла заполнены."                   -ForegroundColor Green
    Write-Host "================================================" -ForegroundColor Green

} catch {
    Write-Host ""
    Write-Host "ОШИБКА: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Строка: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Проверьте:" -ForegroundColor Yellow
    Write-Host "  1. Все четыре Excel-файла в той же папке, что start.bat"
    Write-Host "  2. Файлы не открыты в Excel"
    Write-Host "  3. Файлы не помечены как 'только чтение'"
}

Write-Host ""
Read-Host "Нажмите Enter для закрытия"
