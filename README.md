# SNMP Data Import & Chart Generator

Excel VBA マクロ。SNMP MIB / トラップのテキストデータを読み込み、シートへのインポートとグラフ生成を自動化します。

## 機能

- 指定フォルダ内の `.txt` ファイルを一括処理
- ファイルごとにデータシート (`Data_*`) とグラフシート (`Chart_*`) を自動作成
- SNMP MIB / トラップデータのパースと可視化

## 使い方

1. Excel を開き、`Alt+F11` で VBE（Visual Basic Editor）を起動
2. **挿入 > 標準モジュール** を選択し、`SNMP_Import.bas` の内容を貼り付け
3. `Alt+F8` を押して「`ImportSNMPData`」を選択し、実行
4. ダイアログでテキストファイルが入ったフォルダを選択

## ファイル構成

| ファイル | 説明 |
|---|---|
| `SNMP_Import.bas` | メインマクロ（VBA モジュール） |
| `sample_data/` | サンプル用テキストデータ |

## 動作環境

- Microsoft Excel（Windows）
- VBA マクロを有効にした状態で実行してください
