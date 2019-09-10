const xlsx = require('xlsx');
const {google} = require('googleapis');
const senderBook = xlsx.readFile(path.join(__dirname, 'emails', 'emails.xlsx'));
const senderSheet = senderBook.SheetNames;

let senderRow = xlsx.utils.sheet_to_json(senderBook.Sheets[senderSheet[0]]);

const leadBook = xlsx.readFile(path.join(__dirname, 'emails', 'emails.xlsx'));
const leadSheet = leadBook.SheetNames;

let leadRow = xlsx.utils.sheet_to_json(leadBook.Sheets[leadSheet[0]]);

senderRow.map(function(value) {
    const oAuth2Client = new google.auth.OAuth2(value.client_id, value.client_secret, value.redirect_uris);
    const token = {
        access_token    : value.access_token,
        refresh_token   : value.refresh_token,
        scope           : value.scope,
        token_type      : value.token_type,
        expiry_date     : value.expiry_date
    };
    oAuth2Client.setCredentials(token);
    sendEmail(oAuth2Client, value.email)
})

function sendEmail(auth, email) {
    const gmail = google.gmail({version: 'v1', auth});

    var email_lines = [];

    email_lines.push(`From: ${email}`);
    email_lines.push('To: , ');
    email_lines.push('Content-type: text/html;charset=iso-8859-1');
    email_lines.push('MIME-Version: 1.0');
    email_lines.push('Subject: this would be the subject');
    email_lines.push('');
    email_lines.push('And this would be the content.<br/>');
    email_lines.push('The body is in HTML so <b>we could even use bold</b>');

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
}