# gas-tutor-payment-statement

月次の報酬計算から、講師別のPDF明細と Gmail の下書きまでを自動で作る Google Apps Script の見本です。
架空の教室「サンプル学習塾」、架空の講師20名、架空のデータだけで動きます。

## 課題
毎月、実績の表と単価の表を見比べて講師ごとの報酬を計算し、明細を1人ずつ作って、メールを書いていました。
計算の間違い、作り直し、送る前の確認に時間がかかる作業でした。

## 仕組み
1. 「実績」「マスタ_講師」「マスタ_品目」「設定」のシートを読む（元のデータは書き換えない）
2. ①チェック：足りない値や合わない行を先に洗い出す
3. ②計算：端数処理と税区分を設定どおりに当てて報酬を出す
4. ③PDF作成：講師別の明細をPDFにして保存
5. ④下書き作成：講師ごとの Gmail 下書きを一括で作る（送信は人が行う）
6. ⑤確定：シートを保護し、監査ログを残す
メニュー「月次処理」とサイドバーから、この順に実行します。順番を飛ばしたボタンは押せません。

## 効果
講師の人数が増えても、毎月同じ手順で、計算から下書きまでがそろいます。送る前の確認は下書きの一覧で行えます。

## 画面
画像は後で追加します。

## 設置の手順
SETUP_GUIDE.md の「1. 初期設定」を参照してください。data/ に架空のサンプルデータ（CSV）があります。

## 止まるときと記録
- 失敗した処理は「エラーログ」シートに記録し、確定した月は動かしません
- 確定の前に、同じファイル内の非表示シートにバックアップを取ります

## ファイル
| ファイル | 役割 |
|---|---|
| Code.gs | メニューとサイドバーの入口 |
| Config.gs | 設定値・シート名・列番号 |
| Validate.gs・Calculate.gs | チェックと報酬計算 |
| Pdf.gs・Drive.gs | PDF明細と保存 |
| Draft.gs | Gmail 下書きの一括作成 |
| Finalize.gs | 確定とシートの保護・監査ログ |
| Batch.gs・State.gs | 続きからの再開・進み具合と二重処理の記録 |
| Dashboard.gs | 月別の集計とグラフ |
| Util.gs・Sheets.gs | 共通の処理 |
| Sidebar.html | サイドバー |
| test/ | Node.js で動くテスト |

## 作りの工夫
- 日付：対象年月や今日の日付は、スクリプトのタイムゾーン（appsscript.json の timeZone。見本では Asia/Tokyo）で読みます。海外の会社で使うときは、この設定を会社の地域に変えてください。
- テスト：`node test/run.js`（test/gas_mock の偽物の上で、正常・異常・大量件数・二重処理・6分制限からの再開・送信系が呼ばれないことを確かめる。実際の Apps Script では動かしていません）
- 6分制限からの再開：Batch.gs の runBatchOverInstructors_ が講師1人ずつ処理し、Config.gs の BATCH_TIME_LIMIT_MS（4.5分）を超えたら進み具合を State.gs の saveBatchProgress_ に残し、scheduleContinuation_ でトリガーを作って resumeBatch から続きを動かす
- 二重処理の防ぎ方：State.gs の getIdempotencyMap_／saveIdempotencyMap_ で、作ったPDF（Pdf.gs）と下書き（Draft.gs）を月ごとに記録し、二度作らない。確定した月は isMonthFinalized_ で止める
- 1日の上限：State.gs の assertDailyLimitNotReached_

## 注意
- 架空のデータで動く見本です。実際の運用前に、自分の環境で動きを確かめてください
- Gmail・ドライブ・Apps Script には1日の上限があります

## 作った人
Reypage｜GAS業務自動化

## ライセンス
MIT（LICENSE を参照）
