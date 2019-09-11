// libary import
const xlsx = require('xlsx');
const {google} = require('googleapis');
const fs = require('fs');
const path = require('path');

// sender excel sheet import
const senderBook = xlsx.readFile(path.join(__dirname, 'emails', 'emails.xlsx'));
const senderSheet = senderBook.SheetNames;
let senderRow = xlsx.utils.sheet_to_json(senderBook.Sheets[senderSheet[0]]);

// lead excel sheet import
const leadBook = xlsx.readFile(path.join(__dirname, 'emails', 'leads.xlsx'));
const leadSheet = leadBook.SheetNames;
let leadRow = xlsx.utils.sheet_to_json(leadBook.Sheets[leadSheet[0]]);

// global variable
var gmail, mailSendInterval, mailSubject, mailBody, sendCount = 50;

// fatch config file
fs.readFile(path.join(__dirname, 'config.json'), (err, res) => {
    if (err) {
        console.log('config.json file not found');
        process.exit(1);
    }

    res = JSON.parse(res);

    mailSubject = res.mailSubject;
    mailBody = res.mailBody;
    sendCount = res.sendFromPerMail;
    mailSendInterval = res.mailSendInterval;

    senderRow.map(function(value) {
        console.log(value);
        const oAuth2Client = new google.auth.OAuth2(value.client_id, value.client_secret, value.redirect_uris);
        const token = {
            access_token    : value.access_token,
            refresh_token   : value.refresh_token,
            scope           : value.scope,
            token_type      : value.token_type,
            expiry_date     : value.expiry_date
        };
        
        gmail = google.gmail({version: 'v1', auth: oAuth2Client.setCredentials(token)});
        sendEmail(gmail, value.email)
    });
});

function sendEmail(gmail, email) {
    
    var i = 0;

    var sendEmailInterval = setInterval(() => {
        var leadMail = leadRow.pop();
        
        if ( typeof leadMail === "undefined") {
            console.log('All lead is completed to send email');
            process.exit(1);
        }

        var email_lines = [];
        email_lines.push(`From: ${email}`);
        email_lines.push(`To: ${leadMail}`);
        email_lines.push('Content-type: text/html;charset=iso-8859-1');
        email_lines.push('MIME-Version: 1.0');
        email_lines.push(`Subject: ${mailSubject}`);
        email_lines.push('');
        email_lines.push(mailBody);
      
        var email = email_lines.join('\r\n').trim();

        var base64EncodedEmail = new Buffer.from(email).toString('base64');
        base64EncodedEmail = base64EncodedEmail.replace(/\+/g, '-').replace(/\//g, '_');


        gmail.users.messages.send({
            userId: 'me',
            resource: {
                raw: base64EncodedEmail
            }
        }, (err, res) => {
            if (err) return console.log('The API returned an error: ' + err);
            console.log(res);
        })

        i++;

        if (i === sendCount) clearInterval(sendEmailInterval);

    }, mailSendInterval);
}