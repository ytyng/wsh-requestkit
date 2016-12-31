var RequestKit = {

    /**
     * デバッグ用 objの属性を列挙して表示
     */
    showAttrs: function (obj) {
        var attrs = [];
        for (attr in obj) {
            attrs.push(attr);
        }
        WScript.echo(attrs.join(','));
    },

    /**
     * デスクトップサイズを取得
     * @param ie: this.IE あれば高速に処理する
     * @returns {{width: Number, height: Number}}
     */
    getDesktopSize: function (ie) {
        var ieCreated = false;
        if (!ie) {
            ie = new this.IE({visible: false});
            ie.navigate('about:blank');
            ieCreated = true;
        }
        var size = {
            width: ie.application.document.parentWindow.screen.availWidth,
            height: ie.application.document.parentWindow.screen.availHeight
        };
        if (ieCreated) {
            ie.close();
        }
        return size;
    },

    /**
     * IEを起動
     */
    IE: function (options) {
        this.application = WScript.CreateObject("InternetExplorer.Application");
        if (!options) {
            options = {};
        }
        if (options.visible == null) {
            this.application.Visible = true;
        } else {
            this.application.Visible = options.visible;
        }
        while (this.application.busy) WScript.Sleep(100);
    },

    /**
     * Jsonを取得
     */
    getJson: function (url) {
        var xmlHttp = WScript.CreateObject("Msxml2.ServerXMLHTTP");
        xmlHttp.open("GET", url, false);
        xmlHttp.send();
        eval("var data = " + xmlHttp.responseText + ";");
        return {
            xmlHttp: xmlHttp,
            status: xmlHttp.status,
            headers: xmlHttp.getAllResponseHeaders(),
            data: data
        };
    },

    /**
     * URLリソースをファイルに保存
     */
    downloadSave: function (url, filePath) {
      var xmlHttp = WScript.CreateObject("Msxml2.ServerXMLHTTP");
      xmlHttp.open("GET", url, false);
      xmlHttp.send();
      var stream = WScript.CreateObject("ADODB.Stream");
      stream.type = 1;
      stream.open();
      stream.write(xmlHttp.responseBody);
      stream.saveToFile(filePath, 2);
    },

    /**
     * メールを送信する
     * smtp: host, port, useSsl, userId, password
     * mail: fromAddress, toAddress, subject, body
     */
    sendMail: function (smtp, mail) {
        var cdoSchemas = "http://schemas.microsoft.com/cdo/configuration/";
        var cdoMessage = WScript.CreateObject("CDO.Message");
        cdoMessage.From = mail.fromAddress;
        cdoMessage.To = mail.toAddress;
        cdoMessage.Subject = mail.subject;
        cdoMessage.TextBody = mail.body + "\r\n";
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "sendusing") = 2;
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "smtpserver") = smtp.host;
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "smtpserverport") = smtp.port;
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "sendusername") = smtp.userId;
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "sendpassword") = smtp.password;
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "smtpauthenticate") = 1;
        cdoMessage.Configuration.Fields.Item(cdoSchemas + "smtpusessl") = smtp.useSsl;
        cdoMessage.Configuration.Fields.Update();
        cdoMessage.Send();
    },

    /**
     * 発声する
     */
    say: function(message) {
      var sapi = WScript.CreateObject("SAPI.SpVoice");
      sapi.Speak(message);
    }

};

/**
 * IEでページ移動
 */
RequestKit.IE.prototype.navigate = function (url) {
    this.application.navigate(url);
    while (this.application.busy) WScript.Sleep(100);
    while (this.application.document.readyState != "complete") WScript.Sleep(100);
};

/**
 * ログインフォームにログインする
 */
RequestKit.IE.prototype.login = function (login_id, password) {
    for (var i = 0; i < this.application.document.forms.length; i++) {
        var filled = false;
        var form = this.application.document.forms(i);
        var inputs = form.getElementsByTagName('input');
        var submitButton = null;
        var buttons = form.getElementsByTagName('button');
        for (var j = 0; j < inputs.length; j++) {
            var input = inputs[j];
            if (input.type == 'text') {
                input.value = login_id;
            }
            if (input.type == 'password') {
                input.value = password;
                filled = true;
            }
            if (input.type == 'submit') {
                submitButton = input;
            }
            if (input.type == 'image') {
                submitButton = input;
            }
        }
        if (filled) {
            if (submitButton) {
                submitButton.click();
            } else if (buttons.length) {
                // 最後のボタン (雑)
                buttons[buttons.length - 1].click();
            } else {
                form.submit();
            }
            while (this.application.busy) WScript.Sleep(100);
            while (this.application.document.readyState != "complete") WScript.Sleep(100);
            break;
        }
    }
};


/**
 * ブラウザ上でスクリプトを実行
 */
RequestKit.IE.prototype.script = function (sourceCode, options) {
    while (this.application.busy) WScript.Sleep(100);
    while (this.application.document.readyState != "complete") WScript.Sleep(100);
    if (!options) {
      options = {};
    }
    if (options.mode == "tag") {
      var scriptTag = this.application.document.createElement("script");
      scriptTag.text = sourceCode;
      this.application.document.body.appendChild(scriptTag);
    } else {
      this.application.navigate("javascript:" + sourceCode + ";void(0)");
    }
    while (this.application.busy) WScript.Sleep(100);
    while (this.application.document.readyState != "complete") WScript.Sleep(100);
};

/**
 * CSSセレクタにマッチしたものをクリック
 */
RequestKit.IE.prototype.clickByQuerySelector = function(querySelector) {
    var element = this.application.document.querySelector(querySelector);
    element.click();
    while (this.application.busy) WScript.Sleep(100);
    while (this.application.document.readyState != "complete") WScript.Sleep(100);
};

/**
 * URLリソースを保存 ログインクッキーを使う
 */
RequestKit.IE.prototype.downloadSave = function(url, filePath) {
    var xmlHttp = WScript.CreateObject("Msxml2.ServerXMLHTTP");
    xmlHttp.open("GET", url, false);
    xmlHttp.setRequestHeader('Cookie', this.application.document.cookie);
    xmlHttp.send();
    var stream = WScript.CreateObject("ADODB.Stream");
    stream.type = 1;
    stream.open();
    stream.write(xmlHttp.responseBody);
    stream.saveToFile(filePath, 2);
};

/**
 * IEを閉じる
 */
RequestKit.IE.prototype.close = function () {
    this.application.Quit();
};
