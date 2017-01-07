# WSH RequestKit

Windows Scripting Host を楽に使うためのユーティリティ

よく使うイディオムをパッケージ化し、使いやすくしました。

社内の、非エンジニアの日常業務を簡単に自動化させるのを目的としています。


## 例

### サイトへのログイン

    <?xml version="1.0" encoding="utf-8" standalone="yes" ?>
    <job>
    <script language="JScript" src="wsf-request-toolkit.js"></script>
    <script language="JScript">

    var credential = RequestKit.getJson('https://example.com/get-credential');
    var ie = new RequestKit.IE();
    ie.navigate('http://example.com/login');
    ie.login(credential.data.user_id, credential.data.password);

    </script>
    </job>

※ https://example.com/get-credential へのアクセスで、

    {
      "user_id": "hoge-user",
      "password": "hoge-password"
    }

の JSON が帰ってくると仮定しています。



## リファレンス

### RequestKit

#### new RequestKit.IE()

##### 返り値

InternetExproler.Application のラッパーインスタンス (IEインスタンス)


#### RequestKit.getJson(string URL)

Json を取得します。

##### 返り値

  {
    xmlHttp: xmlHttpオブジェクト,
    headers: 取得ヘッダ,
    data: パースしたデータ
  }

#### RequestKit.downloadSave(string URL, string filePath)

URLの内容をローカルファイルに保存します。

#### RequestKit.getDesktopSize()

デスクトップのサイズを取得します。

##### 返り値

    {
      width: 幅,
      height: 高さ
    }

#### RequestKit.sendMail(Object smtpSetting, Object mail)

メールを送信します。

    sendMail({
        host: "smtp.example.com",
        port: 465,
        useSsl: true,
        userId: "spamman",
        password: "eggs"
      },
      {
        fromAddress: "admin@example.com",
        toAddress: "you@yourdomain.com",
        subject: "件名です。テストメール",
        body: "本文です。テストメール"
      });


#### RequestKit.wmiExecQuery()

WMI Wbem に WQL クエリを発行し、結果を RequestKit.SWbemRecord インスタンスの
配列として取得します。


### RequestKit.IE オブジェクト

#### .navigate(string URL)

url へ遷移します。遷移後はブラウザが準備完了になるまで待ちます。


#### .login(string loginId, string password)

表示しているフォームにログインIDとパスワードを入力してログインします。
フォームの形式によってはログインできない可能性もあります。

#### .script(string sourceCode)

ブラウザでJavaScriptを実行します

#### .clickByQuerySelector(string querySelector)

CSSセレクタにマッチしたエレメントをクリックします

#### .fillInputs(object)

フォームの name と value からなる連想配列を引数で与えると、
ページ内のフォームにその値を入力します。

##### 例

    ie.fillInputs({
        "name": "ほげ太郎",
        "email": "xxxxx@example.com",
        "gender": "1",
        "country": "Japan"
    })

#### .downloadSave(string URL, string filePath)

URLの内容をローカルファイルに保存します。
現在表示しているページのクッキーをリクエストに含めるので、
ログイン済みページのコンテンツをダウンロードできます。

#### .activate()

IEのウインドウをアクティブにします。
ただし複数のIEが起動している場合不具合があり、起動しているIEオブジェクトが
最前面にならずに別のIEが最前面になるかもしれません。
(修正案わからず)

#### .applySaveDialog()

「ファイルを保存しますか?」のダイアログが出てきたとき、「保存」のショートカットを押します。
上記 .activate() と SendKeys() を使うので、動作は不安定です。

#### .close()

IEを閉じます

#### .application

InternetExproler.Application の実体です。


### RequestKit.SWbemRecord オブジェクト

WQL 結果の各フィールドがプロパティとして格納されています。

#### .showAttrs()

フィールドをポップアップで表示します。デバッグ用
