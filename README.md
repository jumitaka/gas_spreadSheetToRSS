# gas_spreadSheetToRSS
SpreadSheet To RSS(Google Apps Script)

# これは何？
GoogleスプレッドシートをRSS化します。

## 利用想定
IFTTTでtwitter検索の結果をSpreadSheetに挿入しておき、その結果をRSSリーダーで読む、など。

# 導入
1. 作成済のGoogleスプレッドシートに、このスクリプトをApps Scriptで導入。
2. 行1に列タイトルを設定。（後述）
3. 行2以降、RSS化するデータを挿入。（その際タイトル列とセットするデータを合わせること）

## 列タイトル
列タイトル設定は以下となっています。

|  itemタグ |  列タイトル  |
| ---- | ---- |
| title       |  Title  |
| author      |  Author |
| link        |  Link   |
| description |  Title  |
| pubDate     |  Date   |
| guid        |  Link   |

getColumnIndexに渡しているcolumnNamesを変更することで、他列タイトルでも対応できます。

# FAQ
## pubDateについて
pubDate他、RSSの日付はRFC822の日付形式にする必要があります。
適宜convertDateを修正してください。
