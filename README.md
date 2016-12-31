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

の JSON が帰ってくる想定。



## リファレンス

### new RequestKit.IE()

#### 返り値

InternetExproler.Application のラッパーインスタンス (IEインスタンス)


### RequestKit.getJson(URL)

Json を取得します。

#### 返り値

  {
    xmlHttp: xmlHttpオブジェクト,
    headers: 取得ヘッダ,
    data: パースしたデータ
  }


### RequestKit.getDesktopSize()

デスクトップのサイズを取得します

#### 返り値

    {
      width: 幅,
      height: 高さ
    }

### RequestKit.sendMail(object smtpSetting, object mail)

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

### IEオブジェクト

### ie.navigate(string "url")

url へ遷移します。遷移後はブラウザが準備完了になるまで待ちます。


### ie.login(string "login_id", string "password")

表示しているフォームにログインIDとパスワードを入力してログインします。
フォームの形式によってはログインできない可能性もあります。

### ie.script(string sourceCode)

ブラウザでJavaScriptを実行します

### ie.clickByQuerySelector(string querySelector)

CSSセレクタにマッチしたエレメントをクリックします

### ie.close()

IEを閉じます

### ie.application

InternetExproler.Application の実体です。
