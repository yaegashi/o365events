# o365events

## 概要

Microsoft Graph API を使って指定した Office 365 アカウントのカレンダーにアクセスし、
イベントと出席者のデータを Excel ないし JSON 形式で出力する CLI ツールです。
SharePoint Online のドキュメントライブラリにアップロードすることもできます。

Linux/Windows/macOS 用の実行ファイルが
[Releases](https://github.com/yaegashi/o365events/releases)
よりダウンロードできます。

## 使用法

```console
$ o365events -h
Usage of o365events:
  -client-id string
        Client ID (default "b7dbe94f-2f3a-4b98-a372-a99d0edff196")
  -end string
        End date (YYYYMMDD)
  -exclude
        Exclude calendar owner from attendees
  -output string
        Output path (default "events.xlsx")
  -start string
        Start date(YYYYMMDD) (default "20200404")
  -tenant-id string
        Tenant ID (default "common")
  -token-cache-path string
        Token cache path (default "token_cache.json")
```

### 認証

o365events を初めて実行すると次のように URL (https://microsoft.com/devicelogin) とコード (`HDZSYAKLD`) を表示します。

```console
$ o365events 
To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code HDZSYAKLD to authenticate.
```

指示のとおりに Web ブラウザで URL を開き、コードを入力してください。
その後サインインを要求されますので、カレンダーにアクセス権限のあるアカウントでサインインしてください。

サインインに成功するとカレントディレクトリに `token_cache.json` ファイルを作成します。
2 回目以降の実行ではこのファイルを読み取ってユーザー認証しますのでサインインを求められることはありません。
このファイルは他人がアクセスできない安全な場所に保存してください。

### me

o365events を引数なしで実行すると、
サインインしたユーザー (`me`) の現在の月のイベントを events.xlsx ファイルに出力します。

```console
$ o365events
2020/04/04 14:33:09 I: User me
2020/04/04 14:33:09 I: Fetching events of admin@l0wdev.onmicrosoft.com (d2a07c12-3806-4f0b-9f86-c39d88de1c83)
2020/04/04 14:33:10 I: Got 1 events
2020/04/04 14:33:10 I: Writing to events.xlsx
```

### JSON 出力

`-output` で拡張子が `.json` のファイル名を指定すると JSON で出力します。
`-` を指定すると標準出力に JSON で出力します。

```console
$ o365events -output events.json
```

### 期間指定

`-start` と `-end` で出力するイベントの期間を日単位で指定できます。

```console
$ o365events -start 20190401 -end 20200331
```

### 会議室

会議室のイベントを出力するには、次のように会議室アカウントのメールアドレスを引数に並べます。
会議室アカウントを出席者リストから除外するために `-exclude` を指定することを推奨します。

```console
$ o365events -start 20200401 -end 20200430 -exclude room-a@l0wdev.onmicrosoft.com room-b@l0wdev.onmicros
oft.com room-c@l0wdev.onmicrosoft.com
2020/04/04 14:34:30 I: User room-a@l0wdev.onmicrosoft.com
2020/04/04 14:34:30 I: Fetching events of room-a@l0wdev.onmicrosoft.com (5fc084ba-b8fb-479f-bbdc-456ea8b7880b)
2020/04/04 14:34:31 I: Got 28 events
2020/04/04 14:34:31 I: User room-b@l0wdev.onmicrosoft.com
2020/04/04 14:34:31 I: Fetching events of room-b@l0wdev.onmicrosoft.com (9349a25d-9f8b-47fd-85d2-22b2a512409d)
2020/04/04 14:34:32 I: Got 3 events
2020/04/04 14:34:32 I: User room-c@l0wdev.onmicrosoft.com
2020/04/04 14:34:33 I: Fetching events of room-c@l0wdev.onmicrosoft.com (85b5f7d3-47f3-4c95-978f-1d846cbaae7d)
2020/04/04 14:34:33 I: Got 5 events
2020/04/04 14:34:33 I: Writing to events.xlsx
```
