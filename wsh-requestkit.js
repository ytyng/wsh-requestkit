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
            width: ie.application.Document.parentWindow.screen.availWidth,
            height: ie.application.Document.parentWindow.screen.availHeight
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
        var xmlHttp = new ActiveXObject("Msxml2.ServerXMLHTTP");
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
     * メールを送信する
     * smtp: host, port, useSsl, userId, password
     * mail: fromAddress, toAddress, subject, body
     */
    sendMail: function (smtp, mail) {
        var cdoSchemas = "http://schemas.microsoft.com/cdo/configuration/";
        var cdoMessage = new ActiveXObject("CDO.Message");
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
    while (this.application.Document.readyState != "complete") WScript.Sleep(100);
};

/**
 * ログインフォームにログインする
 */
RequestKit.IE.prototype.login = function (login_id, password) {
    for (var i = 0; i < this.application.Document.forms.length; i++) {
        var filled = false;
        var form = this.application.Document.forms(i);
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
        }
        if (filled) {
            if (submitButton) {
                submitButton.click();
            } else if (buttons) {
                // 最後のボタン (雑)
                buttons[buttons.length - 1].click();
            } else {
                form.submit();
            }
            while (this.application.busy) WScript.Sleep(100);
            while (this.application.Document.readyState != "complete") WScript.Sleep(100);
            break;
        }
    }
};

/**
 * IEを閉じる
 */
RequestKit.IE.prototype.close = function () {
    this.application.Quit();
};
