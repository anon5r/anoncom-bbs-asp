# anoncomBBS の使い方

## 動作環境

Windows ServerなどのレガシーASP(≠ASP.NET)が動作する環境を用意。
また、_Microsoft Accessデータベース(*.mdb)_が扱えるようにする必要があります。
_FileSystemObject_、_ADODB_が利用できる必要があります。

## インストール

ドキュメントルート以下の適当なところに設置してください。


## 初期設定

<設置したパス>/install.asp にアクセスし、初期設定を実行します。

>http://<example.com>/bbs/install.asp


## 管理者ランクを作る

管理者ランクは __Session(_"AdminLevel"_)__ で設定

データベースには __adminrank__ フィールドを設定


|ランク   |説明          |
|:------:|:-------------|
|       9|最高権限|anoncomBBSに対して何でもできる（データベースの移動など）|
|       8|準最高権限|掲示板管理者管理以外|
|       7|掲示板の削除はできない|
|       6|掲示板の作成はできない|
|       5||
|       4|掲示板設定の初期化ができない、掲示板レスの全消去ができない|
|       3|掲示板設定が変更できない|
|       2|管理者情報の変更ができない|
|       1|掲示板レスの削除ができない|
|       0|何もできない（アカウント停止・権限剥奪）|


[権限詳細](adminrank.html)


* __管理者ユーザアカウントの作成ページ__
  - アカウント作成処理
* __管理者ユーザアカウントの削除ページ__
  - アカウント抹消処理
* __管理者ユーザアカウントの編集ページ__
  - 名前、メールアドレス、アカウント停止、管理権限など



## デバッグモード

**debug_flag** 変数により、その掲示板のデバッグレベルを設定し、レベルに応じて出力する内容を変える。

|項目      |説明                        |
|:--------|:--------------------------|
|デバッグなし|書き込み後「書き込みました」だけ|
|簡易デバッグ|書き込み後に通知配信の場合はメール送信結果を表示|
|デバッグ|書き込み後、書き込み処理時間とメール送信処理時間を表示。通知配信でエラーがあった場合、詳細なエラーを表示|
|完全デバッグ|書き込み後、通知配信をチェックしても実際には配信せず、配信内容のみを書き出す。|
