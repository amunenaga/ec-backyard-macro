# 各CSV取得

function FetchSlimsCsv {
    $Slims = Return Get-content ~\desktop\SISJM.csv | ConvertFrom-Csv
    return = $Slims
}

function FetchYamatoCsv {
    $Yamato = Get-content ~\desktop\1003_ヤマト_A.csv | ConvertFrom-Csv
    return $Yamato

}
function FetchPostalCsv {
    $Postal = Get-content ~\desktop\1003_ゆうパック_A.csv | ConvertFrom-Csv
    return $Postal

}

function FetchclossMallCsv {
    $Crossmall = Get-content ~\desktop\order.csv | ConvertFrom-Csv
    return $Crossmall

}

function GetOrderedProducts($OrderId) {
# クロスモールCSVから管理番号が一致する行を取得
    $RecordSet = $Crossmall | Where-Object { $_.管理番号 -like $OrderId }

    return $RecordSet 

}

function GetSlimsInventry($Code) {
# ロケーションと在庫数を判定して、在庫数を返す
# ロケーションなしは在庫ゼロとみなす

    if ($Code.length -le 7) {
        $Result = $Slims | Where-Object { $_.WSHOCD -eq [int]$Code }
        $Qty = $Result.WKSBQT
    } else {
        # 5ケタ・6ケタでないコードはSlims在庫 0で返す
        $Qty = 0
    }
    
    return $Qty

}

function AllowShipping($RecordSet) {
    # 商品配列に対して、全ての商品がSLIMS在庫有りなら、Trueを返す
    
    # 商品コードをキーとして、在庫数を格納するハッシュを作成
    $Products = @{}
    
    foreach ($Record in $RecordSet){
        $Code = $Record.商品コード        
        $Qty = GetSlimsInventry ($Code)            
        
        # 同一商品で行が分かれる場合があるので、既に商品コードキーがあるかチェックして追加
        if (! $Products.ContainsKey($Code)) {
            $Products.add($Code, $Qty)            
        }
    }

    # 1点でも在庫数0の商品があれば、False
    if ( $Products.ContainsValue(0) ) {
        return $false
    } else {
        return $true
    }

}

function UpdateYamato($OrderId) {
# ヤマトの出荷可能なお客様管理番号に対して、出荷予定日を本日で更新する

    $EstimatedDate = get-date -Format "yyyy/M/d"

    $OrderedProducts = GetOrderedProducts($OrderId)

    if (AllowShipping($OrderedProducts)) {
        $Yamato | Where-Object {$_.お客様管理番号 -like $OrderId} | foreach {$_.出荷予定日 = $EstimatedDate }
    }

}

function UpdatePostal($OrderId) {
# ゆうパックの出荷予定日を本日で更新する
    $EstimatedDate = get-date -Format "yyyy/M/d"

    $OrderedProducts = GetOrderedProducts($OrderId)

    if (AllowShipping($OrderedProducts)) {
        $Postal | Where-Object {$_.お客様側管理番号 -like $OrderId} | foreach {$_.発送予定日 = $EstimatedDate }
    }

}

cd ~\

Get-date -Format "HH:mm:ss"

$Script:Slims = FetchSlimsCsv
$Script:Crossmall = FetchClossmallCsv

$Script:Yamato = FetchYamatoCsv
$Script:Yamato = FetchYamatoCsv
$Script:Postal = FetchPostalCsv

echo "ヤマト処理中"
$yamato | foreach {

    UpdateYamato($_.お客様管理番号)

}

Get-date -Format "HH:mm:ss"

echo "ゆうパック処理中"

$Postal | foreach {
    
        UpdatePostal($_.お客様側管理番号)
    
    }

$yamato | Export-Csv ~\desktop\test_data_yamato.csv -Encoding default -noType
$Postal | Export-Csv ~\desktop\test_data_postal.csv -Encoding default -noType

Get-date -Format "HH:mm:ss"

Read-Host "終了するにはENTERキーを押して下さい" 