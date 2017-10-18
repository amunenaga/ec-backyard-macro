
# GUIのファイルダイアログ、メッセージボックスを使うためのコマンドレット
Add-Type -Assembly System.Windows.Forms

function FetchSlimsCsv {
    # SlimsのダウンロードCSVをファイル指定ダイアログから指定してもらって、CSVオブジェクトを返す
    
    # PowerShellからファイル指定ダイアログを表示する
    # @url:https://letspowershell.blogspot.jp/2015/06/powershell_10.html
    $dialog = New-Object System.Windows.Forms.OpenFileDialog

    if ($dialog.ShowDialog() -eq "OK") {
        $CsvPath = $dialog.FileName
    } else {
        [System.Windows.Forms.MessageBox]::Show("SLIMSファイルの指定がキャンセルされました。処理を終了します", "キャンセル")
        exit
    }

    $Slims = Get-Content $CsvPath | ConvertFrom-Csv

    # ロケーション集約データの場合はロケーションが入っている、在庫数取得ができないのでアラート出して終了
    if ($Slims[1].WLOCCD -notlike "") {
        [System.Windows.Forms.MessageBox]::Show("ロケーション集約データは使えません。ダウンロードをやり直して下さい。")
        exit
    }

    return $Slims

}

function FetchclossMallCsv {
    # 所定のフォルダのクロスモールCSVを全て連結したCSVオブジェクトを返す

    $CsvFiles = Get-Childitem \\server02\商品部\ネット販売関連\ピッキング\クロスモール\order*.csv | Sort-Object LastWriteTime -Descending
    $Crossmall = Get-content $CsvFiles[0].FullName | ConvertFrom-Csv

    return $Crossmall

}

function GetOrderedProducts($OrderId) {
# クロスモールCSVから管理番号が一致する行を取得
    $RecordSet = $Crossmall | Where-Object { $_.管理番号 -like $OrderId }

    return $RecordSet 

}

function GetSlimsInventry($Code) {
# ロケーションと在庫数を判定して、在庫数を返す
# Slimsにデータなしは在庫ゼロとみなす
    
    # 受注時商品コードでハイフンがあれば、それより後ろは削除、頭の0も削除
    $Code = $Code -replace "\-.",""
    $Code = $Code -replace "^0",""
    
    # 5ケタか6ケタなら、コードで照合、それ以外はJANで照合、照合失敗は在庫ゼロとみなす
    # ロケーション集約データでレコードが2個ある時にキャストエラーがでるので考える

    try {
        $Result = $Slims | Where-Object { $_.WSHOCD -eq $Code }
        [long]$Qty = $Result.WKSBQT
    } catch {

        try {
            $Result = $Slims | Where-Object { $_.WJANCD -eq $Code }
            [long]$Qty = $Result.WKSBQT
        } catch {
            [Long]$Qty = 0 
        }

    }

    Return $Qty

}

function AllowShipping($RecordSet) {
    # 商品配列に対して、全ての商品がSLIMS在庫有りなら、Trueを返す
    
    # 商品コードをキーとして、在庫数を格納するタイプセーフなハッシュを作成
    $Products = New-Object 'System.Collections.Generic.Dictionary[string, long]'
    
    foreach ($Record in $RecordSet){
        $Code = $Record.商品コード        
        $Qty = GetSlimsInventry ($Code)            
        
        # クロスモールCSVが同一商品で行が分かれる場合があるので、キー重複時はそのままContinueする
        try {
            $Products.add($Code, $Qty)            
        }
        catch {
            continue
        }
    }

    # 1点でも在庫数0の商品があれば、False
    if ( $Products.ContainsValue(0) ) {
        return $false
    } else {
        return $true
    }

}

function UpdateYamato($Csv) {
# ヤマトの出荷可能なお客様側管理番号に対して、出荷予定日を本日で更新する
# また、送り状記載用に別途、項目追加有り

    $TodayDate= get-date -Format "yyyy/M/d"    

    $Yamato = Get-Content $Csv.Fullname | ConvertFrom-Csv

    $Yamato | ForEach-Object {

        # $_.add("お客様管理番号(送り状用)", $_.お客様管理番号)
        try {
            $OrderedProducts = GetOrderedProducts($_.お客様管理番号)            
        } catch {
            continue
        }
            
        if (AllowShipping($OrderedProducts)) {
            $_.出荷予定日 = $TodayDate            
        }
    }

    $OutPutPath = "~\desktop\"+ $Csv.Name.Replace(".csv","_today") + ".csv" 
    $yamato | Export-Csv $OutPutPath -Encoding default -noType

    # 処理時間確認用のEcho
    Get-date -Format "HH:mm:ss"
    Echo $csv.name + " 処理完了"

}

function UpdatePostal($Csv) {
# ゆうパックの出荷予定日を本日で更新する
    
    $TodayDate= get-date -Format "yyyy/M/d"    

    $Script:Postal = Get-Content $Csv.Fullname | ConvertFrom-Csv
    
    $Script:Postal | ForEach-Object {

        try {
            $OrderedProducts = GetOrderedProducts($_.お客様側管理番号)            
        } catch {
            continue
        }
            
        if (AllowShipping($OrderedProducts)) {
            $_.発送予定日 = $TodayDate
        }

    }

    $OutPutPath = "~\desktop\" + $Csv.Name.Replace(".csv","_today") + ".csv" 
    $Postal | Export-Csv $OutPutPath -Encoding default -noType

    # 処理時間確認用のEcho
    Get-date -Format "HH:mm:ss"
    Echo $csv.name + " 処理完了"

}

$BasePath = "\\server02\商品部\ネット販売関連\梱包室データ\送り状データ\"
$today = Get-Date -Format "MMdd"
$Folder = $BasePath + $today + "\出荷データ"

# クロスモールCSVとSlimsCSVのオブジェクトは、全メソッドから同一オブジェクトを参照する
$Script:Slims = FetchSlimsCsv
$Script:Crossmall = FetchClossmallCsv

echo "ヤマト処理中"

$AllYamatoCsv = ls $Folder -filter *ヤマト*.csv
ForEach ($Csv in $AllYamatoCsv) {

    UpdateYamato($Csv)

    #処理済みCSVを待避

}

echo "ゆうパック処理中"

$AllPostalCsv = ls $Folder -filter *ゆうパック*.csv
ForEach ($Csv in $AllPostalCsv) {
    UpdatePostal($Csv)

    #処理済みCSVを待避
    
}

Echo "ヤマト・ゆうパック処理完了"
Get-date -Format "HH:mm:ss"

Read-Host "終了するにはENTERキーを押して下さい" 