# 家計簿アプリ セットアップ手順

スプレッドシートをデータベースに、GitHub Pagesをフロントにした割り勘家計簿アプリ。

## 1. スプレッドシート + Apps Script の準備

1. Google Sheets で新しいスプレッドシートを作成する
2. メニューの「拡張機能」→「Apps Script」を開く
3. デフォルトの `Code.gs` の中身を全部消して、このリポジトリの `gas/Code.gs` の内容を貼り付ける
4. 保存する（Ctrl+S）
5. 右上の「デプロイ」→「新しいデプロイ」をクリック
6. 種類の選択で「ウェブアプリ」を選ぶ
7. 設定:
   - 実行するユーザー: **自分**
   - アクセスできるユーザー: **全員**
8. 「デプロイ」をクリックし、承認を求められたら自分のGoogleアカウントで許可する
9. 発行された **ウェブアプリのURL** をコピーする（`https://script.google.com/macros/s/xxxx/exec` の形）

https://script.google.com/macros/s/AKfycbxwMA2biwcJz5GbjjR5xyYwqSbw6cRSE5sA-aRtwQg0_bO7s1zU6H03NKZiZVfMU50-/exec

AKfycbxwMA2biwcJz5GbjjR5xyYwqSbw6cRSE5sA-aRtwQg0_bO7s1zU6H03NKZiZVfMU50-

初回実行時にシートが自動で作られます（シート名: `支出`）。

### 合言葉の変更

`gas/Code.gs` 内の `PASSPHRASE` を書き換えて、再度「デプロイ」→「デプロイを管理」→ 既存デプロイの編集で新バージョンとしてデプロイし直せば変更できます。

## 2. フロント（index.html）の設定

`index.html` 内の以下の行を、手順1でコピーしたURLに書き換える。

```js
const GAS_URL = 'PUT_YOUR_GAS_WEBAPP_URL_HERE';
```

## 3. GitHub Pages で公開

1. このフォルダの内容でGitHubリポジトリを作成し、push する
2. リポジトリの Settings → Pages → Source を「main ブランチ / (root)」に設定
3. 発行されたURL（`https://<user>.github.io/<repo>/`）にアクセス

## 使い方

- 初回、記録を追加/削除しようとすると合言葉の入力を求められる（一度入力すればブラウザに保存され、以後聞かれない）
- 「支出」は自動で半額ずつ負担として計算される
- 「精算」はどちらかが相手に実際にお金を渡したときに記録する（全額がそのまま残高に反映される）
- 画面上部に「誰が誰にいくら借りているか」が自動計算されて表示される
