//
// Mail.js
// Extension to simplify sending of e-mails via CDO.Message
//
// Copyright (c) 2012 by Ildar Shaimordanov
//

function Mail(options)
{
        this.driver = new ActiveXObject('CDO.Message');
        this.configure(options);
};

(function()
{

var fieldsUpdate = function(fields, prefix, options)
{
        for (var p in options) {
                if ( ! p || ! options.hasOwnProperty(p) ) {
                        continue;
                }
                fields.Item(prefix + p) = options[p];
        }

        fields.Update();
};

Mail.prototype.configure = function(options)
{
        fieldsUpdate(
                this.driver.Configuration.Fields,
                'http://schemas.microsoft.com/cdo/configuration/',
                options);
};

Mail.prototype.setMailHeaders = function(options)
{
        fieldsUpdate(
                this.driver.Fields,
                'urn:schemas:mailheader:',
                options);
};

Mail.prototype.addAttachment = function(URL, mimetype, options)
{
        options = options || {};
        var attachment = this.driver.AddAttachment(URL, options.username, options.password);
        if ( mimetype ) {
                attachment.ContentMediaType = mimtype;
        }
};

Mail.prototype.prepareMessage = function(options)
{
        var driver = this.driver;

        for (var p in options) {
                if ( ! p || ! options.hasOwnProperty(p) ) {
                        continue;
                }
                driver[p] = options[p];
        }
};

Mail.prototype.send = function(options)
{
        this.prepareMessage(options);
        this.driver.Send();
};

var prepareSmtpOptions = function(options, host, port)
{
        options = options || {};
        options.sendusing = 2;
        /* options.smtpusessl = true; */
		/* options.smtpauthenticate = 1; */
        options.smtpserver = host;
        options.smtpserverport = port;
        return options;
};

Mail.VOE = function(options)
{
		return new Mail(prepareSmtpOptions(options, 'mail.voel.ru', 25));
};


Mail.MailRu = function(options)
{
        return new Mail(prepareSmtpOptions(options, 'smtp.mail.ru', 25));
};

Mail.Yandex = function(options)
{
        return new Mail(prepareSmtpOptions(options, 'smtp.yandex.ru', 465));
};

Mail.GMail = function(options)
{
        return new Mail(prepareSmtpOptions(options, 'smtp.gmail.com', 465));
};

})();