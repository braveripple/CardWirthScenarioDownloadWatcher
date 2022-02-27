# CardWirthScenarioDownloadWatcher
CardWirthのシナリオダウンロードを監視して
CardWirthのシナリオフォルダへの移動を自動で行うスクリプト

## 起動方法
`CardWirthScenarioDownloadWatcher.ps1`をダブルクリックすると、タスクバーの通知領域に`CardWirthScenarioDownloadWatcher`の通知アイコンが出る。(アイコンは今のところPowerShellのアイコンと同じ。)

## 終了方法
通知アイコンを右クリックして「Exit」をクリックすると終了する。

## 設定方法
`CardWirthScenarioDownloadWatcher.ps1`
をメモ帳などで開き、以下の変数を変更する。
* `$WATCH_DIRECTORY_PATH` … ダウンロードのディレクトリのパス
* `$CARDWIRTH_SCENARIO_DIRECTORY_PATH` … CardWirthのシナリオフォルダのパスを指定する。
* `$CARDWIRTHNEXT_SCENARIO_DIRECTORY_PATH` … CardWirthNextのシナリオフォルダのパスを指定する。

## 動作環境
Windows11,PowerShell7で動作確認。
以下のPowerShellのモジュールが必要。
* [CardWirthScenarioSummaryReader(自作)](https://github.com/braveripple/CardWirthScenarioSummaryReader)
* [BurntToast](https://github.com/Windos/BurntToast)

## 課題
* 設定項目を外部ファイル化したい。
* エンジンの設定を可変にしたい。
