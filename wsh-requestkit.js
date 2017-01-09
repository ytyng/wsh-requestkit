/* -*- coding: cp932 -*- */
/* vim: set expandtab fenc=cp932 */
var RequestKit = {
    /**
     * �Ȉ� enumerate�B�R���N�V�����ɏ��� func ��K�p����B
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
     * �Ȉ� each�B�z��v�f�ɏ��� func ��K�p����B
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
     * �f�o�b�O�p obj�̑�����񋓂��ĕ\��
     */
    showAttrs: function (obj) {
        var attrs = [];
        for (attr in obj) {
            attrs.push(attr);
        }
        WScript.echo(attrs.join(','));
    },

    /**
     * �f�X�N�g�b�v�T�C�Y���擾
     * @param ie: this.IE ����΍����ɏ�������
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
     * IE���N��
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
     * �N���ς� IE ���^�C�g����v�ŒT���� IE �C���X�^���X�ɂ���
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
     * �N���ς� IE ��URL��v�ŒT���� IE �C���X�^���X�ɂ���
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
     * Json���擾
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
     * URL���\�[�X���t�@�C���ɕۑ�
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
     * ���[���𑗐M����
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
     * ��������
     */
    say: function (message) {
        var sapi = WScript.CreateObject("SAPI.SpVoice");
        sapi.Speak(message);
    },

    /**
     * WMI Wbem �� WQL �N�G���𔭍s���A���ʂ� SWbemRecord �̔z��Ŏ擾
     */
    wmiExecQuery: function (query, namespace) {
        if (!namespace) {
            namespace = "root\\cimv2";
        }
        var swbemLocator = WScript.CreateObject("WbemScripting.SWbemLocator");
        var wmiServer = swbemLocator.ConnectServer(null, namespace);
        var swbemObjectSet = wmiServer.ExecQuery(query);
        // var e = new Enumerator(swbemObjectSet);
        var swbemRecords = [];
        this.enumerate(swbemObjectSet, function (swbemObject) {
            swbemRecords.push(new this.SWbemRecord(swbemObject));
        }, this);
        return swbemRecords;
    },

    /**
     * �A�v���P�[�V������ exe ���w��ŃA�N�e�B�u�ɂ���B
     * �ʏ�́AwscriptShell.AppActivate �ɃA�v����������Ό����v�ŃA�N�e�B�u�������͂������A
     * ���ꂾ�� IE ���A�N�e�B�u�ɂȂ�Ȃ��̂ł��̃��\�b�h�ł��g��
     * ��: RequestKit.activateProcessByExeName("iexplore.exe");
     * �N�G���Ƀ}�b�`�������ׂĂ�IE�� AppActivate ���Ă��邪�A
     * �A�N�e�B�u�ɂȂ�̂�1���� (����IE�ōőO�ʂ̂��̂���?) �̂悤��
     */
    activateProcessByExeName: function (exeName) {
        var wscriptShell = WScript.CreateObject("WScript.Shell");
        var query = "SELECT Caption, ProcessId FROM Win32_Process WHERE Caption='" + exeName + "'";
        var propRecords = this.wmiExecQuery(query, null);
        this.each(propRecords, function(i, r){
            wscriptShell.AppActivate(r.ProcessId);
        });
    },

    /**
     * SWbemObject �����b�v�����N���X
     * wmiExecQuery ���\�b�h�ō����B
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
 * IE�Ńy�[�W�ړ�
 */
RequestKit.IE.prototype.navigate = function (url) {
    this.application.navigate(url);
    while (this.application.busy) WScript.Sleep(100);
    while (this.application.document.readyState != "complete") WScript.Sleep(100);
};

/**
 * ���O�C���t�H�[���Ƀ��O�C������
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
                // �Ō�̃{�^�� (�G)
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
 * �u���E�U��ŃX�N���v�g�����s
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
 * CSS�Z���N�^�Ƀ}�b�`�������̂��N���b�N
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
 * �t�H�[���̓��͂�����
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

            var tagName = element.tagName.toLowerCase();
            if (tagName == 'input') {
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
 * IE �̃E�C���h�E���őO�ʂɂ��܂��B
 * �������A�����̃E�C���h�E������ꍇ�A�ǂꂪ�őO�ʂɂȂ邩�s��ł��B
 * ���C�Ă킩�炸
 */
RequestKit.IE.prototype.activate = function () {
    // var wscriptShell = WScript.CreateObject("WScript.Shell");
    // wscriptShell.AppActivate('Internet Explorer');
    RequestKit.activateProcessByExeName("iexplore.exe");
};

/**
 * �ۑ����邩�m�F���Ă���_�C�A���O�����(�ۑ�����)
 */
RequestKit.IE.prototype.applySaveDialog = function () {
    this.activate();
    var wscriptShell = WScript.CreateObject("WScript.Shell");
    wscriptShel.SendKeys('%{s}');
};

/**
 * URL���\�[�X��ۑ� ���O�C���N�b�L�[���g��
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
 * IE�����
 */
RequestKit.IE.prototype.close = function () {
    this.application.Quit();
};

// ============================================================================================================
// RequestKit.SWbemRecord
// ------------------------------------------------------------------------------------------------------------
/**
 * �S�v���p�e�B��\�� (�f�o�b�O�p)
 */
RequestKit.SWbemRecord.prototype.showAttrs = function () {
    var log = [];
    RequestKit.each(this.propertyNames, function(i, v){
        log.push(v + ": " + this[v]);
    }, this);
    WScript.Echo(log.join("\n"));
};
