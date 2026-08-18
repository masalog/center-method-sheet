# insert_image.ps1
# Excelに画像を挿入するPowerShellスクリプト（COM経由）
# 既存の吹き出し図形は保持されます

# --- 設定 ---
$excelPath = Join-Path $PSScriptRoot "center_sheet_C-1-2.xlsx"
$imgPath = Join-Path $PSScriptRoot "images\image_001.png"

# --- 画像サイズ・位置（ポイント単位） ---
$imgWidth = 267.7
$imgHeight = 642.5
$imgLeft = 187.6
$imgTop = 150.8

# --- Excel操作 ---
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($excelPath)
    $ws = $wb.Sheets.Item("C1_2")

    # 既存の画像のうち、同じ位置にあるものを削除（吹き出し画像は残す）
    for ($i = $ws.Pictures().Count; $i -ge 1; $i--) {
        $pic = $ws.Pictures().Item($i)
        if ([Math]::Abs($pic.Left - $imgLeft) -lt 5 -and [Math]::Abs($pic.Top - $imgTop) -lt 5) {
            $name = $pic.Name
            $pic.Delete()
            Write-Host "Deleted: $name"
        }
    }

    # 新しい画像を挿入
    $pic = $ws.Pictures().Insert($imgPath)
    $pic.Left = $imgLeft
    $pic.Top = $imgTop
    $pic.Width = $imgWidth
    $pic.Height = $imgHeight

    # 保存して閉じる
    $wb.Save()
    $wb.Close()
    $excel.Quit()

    Write-Host "Done: Image inserted successfully."
}
catch {
    Write-Host "Error: $_"
    if ($excel) {
        try { $excel.Quit() } catch [System.Exception] {}
    }
    exit 1
}
finally {
    if ($excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
}
