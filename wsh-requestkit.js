var RequestKit = {
    /**
     * 簡易 enumerate。コレクションに順に func を適用する。
     */
    enumerate: function (collection, func, _this) {
        if (!_this) {
            _this = this;
        }
        var e = new Enumerator(collection);
        while (!e.atEnd()) {
            func.call(_this, e.item());
            e.moveNext();
        }
    },

    /**
     * 簡易 each。配列要素に順に func を適用する。
     */
    each: function (ary, func, _this) {
        if (!_this) {
            _this = this;
        }
        for (var i = 0; i < ary.length; i++) {
            func.call(_this, i, ary[i]);
        }
    },

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
        if (!options) {
            options = {};
        }
        if (options.application) {
            this.application = options.application;
        } else {
            this.application = WScript.CreateObject("InternetExplorer.Application");
        }
        if (options.visible == null) {
            this.application.Visible = true;
        } else {
            this.application.Visible = options.visible;
        }
        while (this.application.busy) WScript.Sleep(100);
    },

    /**
     * 起動済み IE をタイトル一致で探して IE インスタンスにする
     */
    findIEByTitle: function (title) {
        var shell = WScript.CreateObject('Shell.Application');
        var windows = shell.windows();
        for (i = 0; i < windows.count; i++) {
            var w = windows.Item(i);
            if (w.document) {
                if (w.document.title == title) {
                    return new this.IE({application: w});
                }
            }
        }
    },

    /**
     * 起動済み IE をURL一致で探して IE インスタンスにする
     */
    findIEByUrl: function (url) {
        var shell = WScript.CreateObject('Shell.Application');
        var windows = shell.windows();
        for (i = 0; i < windows.count; i++) {
            var w = windows.Item(i);
            if (w.LocationUrl == url) {
                return new this.IE({application: w});
            }
        }
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
    say: function (message) {
        var sapi = WScript.CreateObject("SAPI.SpVoice");
        sapi.Speak(message);
    },

    /**
     * WMI Wbem にクエリを発行し、結果を SWbemPropertyRecord の配列で取得
     */
    wmiExecQuery: function (query, namespace) {
        if (!namespace) {
            namespace = "root\\cimv2";
        }
        var swbemLocator = WScript.CreateObject("WbemScripting.SWbemLocator");
        var wmiServer = swbemLocator.ConnectServer(null, namespace);
        var swbemObjectSet = wmiServer.ExecQuery(query);
        var e = new Enumerator(swbemObjectSet);
        var swbemRecords = [];
        var self = this;
        this.enumerate(swbemObjectSet, function (swbemObject) {
            swbemRecords.push(new self.SWbemRecord(swbemObject));
        });
        return swbemRecords;
    },

    /**
     * アプリケーションを exe 名指定でアクティブにする。
     * 通常は、wscriptShell.AppActivate にアプリ名を入れれば後方一致でアクティブ化されるはずだが、
     * それだと IE がアクティブにならないのでこのメソッドでを使う
     * 例: RequestKit.activateProcessByExeName("iexplore.exe");
     * クエリにマッチしたすべてのIEを AppActivate しているが、
     * アクティブになるのは1つだけ (複数IEで最前面のものだけ?) のようだ
     */
    activateProcessByExeName: function (exeName) {
        var wscriptShell = WScript.CreateObject("WScript.Shell");
        var query = "SELECT Caption, ProcessId FROM Win32_Process WHERE Caption='" + exeName + "'";
        var propRecords = this.wmiExecQuery(query, null);
        this.each(propRecords, function(r){
            wscriptShell.AppActivate(r.ProcessId);
        });
    },

    /**
     * SWbemObject をラップしたクラス
     * wmiExecQuery メソッドで作られる。
     */
    SWbemRecord: function (swbemObject) {
        this.swbemObject = swbemObject;
        this.propertyNames = [];
        RequestKit.enumerate(swbemObject.Properties_, function (swbemProperty) {
            if (swbemProperty.Value != null) {
                if (swbemProperty.IsArray) {
                    this[swbemProperty.Name] = new VBArray(swbemProperty.Value).toArray();
                } else {
                    this[swbemProperty.Name] = swbemProperty.Value;
                }
                this.propertyNames.push(swbemProperty.Name);
            }
        }, this);
    }

};

// ============================================================================================================
// RequestKit.IE
// ------------------------------------------------------------------------------------------------------------

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
RequestKit.IE.prototype.clickByQuerySelector = function (querySelector) {
    var element = this.application.document.querySelector(querySelector);
    element.click();
    while (this.application.busy) WScript.Sleep(100);
    while (this.application.document.readyState != "complete") WScript.Sleep(100);
};

/**
 * console.log()
 */
RequestKit.IE.prototype.log = function (message) {
    this.script("console.log('" + message + "')");
};

RequestKit.IE.prototype.logAtters = function (obj) {
    for (var attr in obj) {
        this.log(attr + ": " + obj[attr]);
    }
};

/**
 * フォームの入力をする
 */
RequestKit.IE.prototype.fillInputs = function (nameValues) {
    for (var name in nameValues) {
        var value = nameValues[name];
        var querySelector = '[name="' + name + '"]';
        var elements = this.application.document.querySelectorAll(querySelector);
        if (!elements) {
            continue;
        }
        for (var j = 0; j < elements.length; j++) {
            var element = elements[j];
            // this.logAtters(element);
            // WScript.Echo(element.tagName);
            //RequestKit.showAttrs(element);
            var tagName = element.tagName.toLowerCase();
            if (tagName == 'input') {
                // WScript.Echo(element.type);
                if (element.type.toLowerCase() == 'radio') {
                    if (element.value == value) {
                        element.checked = true;
                    }
                } else if (element.type.toLowerCase() == 'checkbox') {
                    if (element.value == value) {
                        element.checked = true;
                    }
                } else {
                    element.value = value;
                }
            } else if (tagName == 'textarea') {
                element.innerText = value;
            } else if (tagName == 'select') {
                var options = element.getElementsByTagName('option');
                for (var i = 0; i < options.length; i++) {
                    var option = options[i];
                    var text = option.innerText.replace(/(^\s+)|(\s+$)/g, "");
                    if (option.value == value || text == value) {
                        element.selectedIndex = i;
                        break;
                    }
                }
            }
        }
    }
};

/**
 * ダウンロード確認ダイアログを閉じる
 * 自身のIEだけをアクティブにしたかったがやり方がわからなかった。
 * AppActivate は、ウインドウ名を後方一致で検索する
 */
RequestKit.IE.prototype.activate = function () {
    // var wscriptShell = WScript.CreateObject("WScript.Shell");
    // wscriptShell.AppActivate('Internet Explorer');
    RequestKit.activateProcessByExeName("iexplore.exe");
};

/**
 * 保存するか確認しているダイアログを閉じる(保存する)
 */
RequestKit.IE.prototype.applySaveDialog = function () {
    this.activate();
    var wscriptShell = WScript.CreateObject("WScript.Shell");
    wscriptShel.SendKeys('%{s}');
};

/**
 * URLリソースを保存 ログインクッキーを使う
 */
RequestKit.IE.prototype.downloadSave = function (url, filePath) {
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

// ============================================================================================================
// RequestKit.SWbemRecord
// ------------------------------------------------------------------------------------------------------------
/**
 * 全プロパティを表示 (デバッグ用)
 */
RequestKit.SWbemRecord.prototype.showAttrs = function () {
    var log = [];
    RequestKit.each(this.propertyNames, function(v){
        log.push(v + ": " + this[v]);
    }, this);
    WScript.Echo(log.join("\n"));
};
