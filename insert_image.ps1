# insert_image.ps1                                                              # スクリプト名
# Excelに画像を挿入するPowerShellスクリプト（COM経由）                           # 概要説明
# 既存の吹き出し図形は保持されます                                               # 注意事項

# --- 設定 ---
$excelPath = Join-Path $PSScriptRoot "center_sheet_C-1-2.xlsx"                   # Excelファイルのパス（スクリプトと同じフォルダ）
$imgPath = Join-Path $PSScriptRoot "images\image_001.png"                        # 挿入する画像のパス

# --- 画像サイズ・位置（ポイント単位） ---
$imgWidth = 267.7                                                                # 画像の幅
$imgHeight = 642.5                                                               # 画像の高さ
$imgLeft = 187.6                                                                 # 画像の左位置
$imgTop = 150.8                                                                  # 画像の上位置

# --- Excel操作 ---
$excel = $null                                                                   # Excel変数を初期化
try {                                                                            # エラー処理の開始
    $excel = New-Object -ComObject Excel.Application                             # Excelアプリを起動
    $excel.Visible = $false                                                      # 非表示で起動
    $excel.DisplayAlerts = $false                                                # 警告ダイアログを抑制

    $wb = $excel.Workbooks.Open($excelPath)                                      # ブックを開く
    $ws = $wb.Sheets.Item("C1_2")                                                # シート「C1_2」を選択

    # 既存の画像のうち、同じ位置にあるものを削除（吹き出し画像は残す）
    for ($i = $ws.Pictures().Count; $i -ge 1; $i--) {                            # 画像を末尾から順に走査
        $pic = $ws.Pictures().Item($i)                                           # i番目の画像を取得
        if ([Math]::Abs($pic.Left - $imgLeft) -lt 5 -and [Math]::Abs($pic.Top - $imgTop) -lt 5) {  # 位置が一致するか判定（誤差5pt以内）
            $name = $pic.Name                                                    # 削除前に名前を保存
            $pic.Delete()                                                        # 画像を削除
            Write-Host "Deleted: $name"                                          # 削除した画像名を表示
        }
    }

    # 新しい画像を挿入
    $pic = $ws.Pictures().Insert($imgPath)                                       # 画像を挿入
    $pic.Left = $imgLeft                                                         # 左位置を設定
    $pic.Top = $imgTop                                                           # 上位置を設定
    $pic.Width = $imgWidth                                                       # 幅を設定
    $pic.Height = $imgHeight                                                     # 高さを設定

    # 保存して閉じる
    $wb.Save()                                                                   # ブックを上書き保存
    $wb.Close()                                                                  # ブックを閉じる
    $excel.Quit()                                                                # Excelアプリを終了

    Write-Host "Done: Image inserted successfully."                              # 完了メッセージ
}
catch {                                                                          # エラー発生時
    Write-Host "Error: $_"                                                       # エラー内容を表示
    if ($excel) {                                                                # Excelが起動していれば
        try { $excel.Quit() } catch [System.Exception] {}                        # 強制終了（失敗しても無視）
    }
    exit 1                                                                       # 異常終了
}
finally {                                                                        # 成功・失敗に関わらず必ず実行
    if ($excel) {                                                                # Excelオブジェクトがあれば
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null  # COMオブジェクトを解放
    }
    [GC]::Collect()                                                              # ガベージコレクション実行
}
