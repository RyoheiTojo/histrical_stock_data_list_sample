# histrical_stock_data_list_sample

## 概要

指定銘柄のヒストリカルデータを指定期間でCSVダウンロード出来るように取得＆整形するマクロ。実態としてはGoogleFinanceのAPIをセルに埋め込んで、取得したデータを整形しているだけ。

## 使い方

### 「対象銘柄情報」シートを作る

![スクリーンショット 2021-07-10 13 01 44](https://user-images.githubusercontent.com/49355629/125151003-0079a380-e17f-11eb-9cfb-935cc34d3a23.png)
```
C1：="all"
C2：開始日 （例：=DATE(2018,3,1)）
C3：終了日 （例：=DATE(2021,6,7))
C4：インターバル（例：="DAILY")

F列（1行目から順に）NYSEARCA:VOOなど。
```

### CreateStockStickDataを動かす

![スクリーンショット 2021-07-10 13 02 52](https://user-images.githubusercontent.com/49355629/125151047-564e4b80-e17f-11eb-8113-afeec7387efa.png)

「対象銘柄中間結果」シートにGoogleFinanceの結果が出る。
※ただし、指定した期間の中で非営業日の日数分空行が出来る（営業日計算を避けている為）
![スクリーンショット 2021-07-10 13 10 45](https://user-images.githubusercontent.com/49355629/125151182-42571980-e180-11eb-81fb-f1e229ff5792.png)

### FormatStockStickDataを動かす

![スクリーンショット 2021-07-10 13 03 12](https://user-images.githubusercontent.com/49355629/125151057-623a0d80-e17f-11eb-9c32-c4e63ae93fc0.png)

「対象銘柄結果」シートに整形された結果が出る。

![スクリーンショット 2021-07-10 13 04 08](https://user-images.githubusercontent.com/49355629/125151062-66fec180-e17f-11eb-9894-c963ceabb4f1.png)

## 注意

1. マクロ実行はGASの仕様上数分でタイムアウトになるので、処理が重い場合は作りの見直しが必要
2. GASのAPI呼び出しが増えると異常に遅くなってタイムアウトし易い為、セル操作のRangeは大きく取っている
3. 後続処理を動かす際に先行して呼び出したAPIが完了する保証があるのか未確認

## 発展

1. SpreadSheetをPythonAPIで動かす仕組みがある気がするので、それと組み合わせればPythonから直接コールする作りにすることも可能かも
