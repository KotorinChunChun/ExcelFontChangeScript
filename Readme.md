# Excel游フォント変更プログラム



## 概要

本ソフトウェアは、Excelのフォントを游書体からMSゴシックに変更するためのスクリプトとExcelマクロです。



作成までの経緯はブログ記事を御覧ください。

[Excelから游ゴシック体を徹底的に駆逐する Part1](https://www.excel-chunchun.com/entry/2019/02/18/010021)

[Excelから游ゴシック体を徹底的に駆逐する Part2](https://www.excel-chunchun.com/entry/FontChange2)

[Excelから游ゴシック体を徹底的に駆逐する Part3](https://www.excel-chunchun.com/entry/FontChange3)



## 内容

- Excel_FontChange_Install.vbs
    - Excel_FontChange.vbsをタスクスケジューラに登録
    - 必要なデータ一式を「C:\Program Files\Excel_FontChange」はインストール
    - 新規作成にxlsmやxlsを追加（レジストリを変更）
- Excel_FontChange.vbs
    - テンプレートファイルを置換
    - 新規作成ファイルを置換
- ExtFontChange.bas.vba
    - ExcelVBAにインポートして使用する
    - アクティブブックの既定のフォントを統一するマクロ（游ゴシック、MSPゴシック、メイリオ）



##  お約束

- コードの安全性／正確性について、私は保証できません。

- 自己責任で利用できる人だけがご利用ください。

- Excel VBAでデバッグを行っているため、`#IF～#End IF`の配置の都合上、おかしな位置に変数の宣言があったりしますがご了承ください。

- 主に自分用コードのため、失敗事例や廃案になったモノも全てコメントで残されています。不満があればセルフサービスでお願いします。

- すべての環境に対しての動作は保証できません。（Office365 / Excel 2016 32bit/64bit で検証）

- 不具合等は詳しい環境情報とエラーメッセージを頂ければ、対処できるかもしれません。来るものは拒みませんので、遠慮なく教えてください。
- スクリプトの転用、改造、会社利用、好きにしていただいて構いませんが、差し支え無ければ引用元としてブログを紹介して頂けると幸いです。



##  インストール方法

- 必ずZIPを展開
- Excel_FontChange_Install.vbsをダブルクリック
- ユーザーアカウント制御が出るので「はい」を選択
- インストールを選択



##  アンインストール方法

- 必ずZIPを展開
- Excel_FontChange_Install.vbsをダブルクリック
- ユーザーアカウント制御が出るので「はい」を選択
- アンインストールを選択


## 既知のバグ

- 本スクリプトを使用しても、次の操作を行った場合のみ游ゴシックが復活することを確認しています。こちらの対処法は見つかっておらず、同梱しているVBAマクロなどにより対応頂く必要があります。
    1. シートを右クリック
    2. シートの移動またはコピー
    3. 移動先ブック名で「(新しいブック)」を選択
    4. OK
- 図形の変形問題は、条件次第で完全には対策しきれない場合があります。
- 2016から増えたストアアプリ版（UWP版）Officeのセキュリティが突破できないため対応させていません。

<br>

UWP版Officeの`EXCEL.EXE`のパスは下記のような感じです。

```
C:\Program Files\WindowsApps\Microsoft.Office.Desktop.Excel_バージョンなどEXCEL.EXE
```

<br>

コントロールパネルのアプリと機能を見た時に

* デスクトップアプリ版

```
Microsoft Office XXXXXXXXXX 2016 ‐ ja-jp
```

* ストアアプリ版

```
Microsoft Office Desktop Apps
```

となっているので見分けることができます。

<br>

そもそもUWP版は

- 外部で入手したアドオンがインストールできない。
- COMが使えないので、サードパーティ製ソフトとの連携が取れない。

と言った問題があるので、従来と同じように利用したい方は注意が必要です。

<br>

昨今はWindows10にUWP版がプリインストールされていることが多いので、知らずにUWP版を使っている人が増えていますが、UWP版では制約があり動作しません。

問題が起きた人は[MicrosoftのOffice削除ツール](https://support.office.com/ja-jp/article/pc-%E3%81%8B%E3%82%89-office-%E3%82%92%E3%82%A2%E3%83%B3%E3%82%A4%E3%83%B3%E3%82%B9%E3%83%88%E3%83%BC%E3%83%AB%E3%81%99%E3%82%8B-9dd49b83-264a-477a-8fcc-2fdf5dbf61d8#OfficeVersion=%E3%82%AF%E3%82%A4%E3%83%83%E3%82%AF%E5%AE%9F%E8%A1%8C%E3%81%BE%E3%81%9F%E3%81%AF_MSI)などで完全消去してから、公式のオンラインダウンロードにてデスクトップ版アプリを入れ直すようにすることを推奨します。

情シスの方はデスクトップアプリ版に入れ替えてから出荷してあげると良いかと思います。

<br>


## VBAマクロをリボンへ登録する方法

ワンクリックで使えるようにするために、マクロを個人用マクロブックに入れて、リボンに登録しておきましょう。

1. 個人用マクロを作成
   - 無い場合は「マクロの記録」から「個人用マクロブック」を選択して適当に記録すると勝手に作成してくれる。
   - [f:id:Kotori-ChunChun:20190720235456p:plain]
2. VBEを開く
   - Alt+F11
3. 標準モジュールを作成
   - VBAProjectを右クリック
   - →挿入
   - →標準モジュール
4. 上記プログラムをコピペ
5. VBEを閉じる
6. リボンのユーザー設定
   - 適当にリボン領域を右クリック
7. マクロを登録
   - 右側のリストにて「ホーム」を選択
   - 右側のリスト下部にて「新しいグループ」をクリック
   - 左側のリスト上部にて「マクロ」を選択
   - 「PERSONAL.XLSB!ブック全体の游フォントをMSPフォントに変更」を選択して「追加」
   - 「OK」

あとはボタンを押して実行するだけ♪

言葉で分からない人は[Excel マクロ 登録方法 で Google検索](https://www.google.com/search?q=Excel+マクロ+登録方法)してください。



##  作者情報

作者：ことりちゅん

Twitter：@KotorinChunChun

ブログ：<[えくせるちゅんちゅん](https://www.excel-chunchun.com/)>

[GitHubダウンロード](https://github.com/KotorinChunChun/ExcelFontChangeScript/archive/master.zip)

[GitHubリポジトリを閲覧](https://github.com/KotorinChunChun/ExcelFontChangeScript)



##  更新履歴

| 日付     | 概要         |
| -------- | ------------ |
| 2019/6/5 | 初回リリース |
| 2020/8/23 | 64bit対応とGitHub公開とフォント相互変換に対応 |


## 謝辞

64bit対応に際して、[furyutei様](https://github.com/furyutei) が行った[変更](https://twitter.com/furyutei/status/1181222249371582464?s=20)を反映させていただきました。
並びにTwitterやBlogからの指摘を元に修正を加えております。

