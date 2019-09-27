const fs = require('fs');
const path = require('path');
var flag = false;

require('getmac').getMac(function(err, macAddress) {
    if (err)  throw err;
    let computer = fs.readFileSync(path.join(__dirname, 'computerAuth.txt'), 'utf8');
    
    if (macAddress === computer) {
        flag = true;
    } else {
        process.exit(1);
    }

    const xlsx = require('xlsx');
    const {google} = require('googleapis');
    const nodemailer = require('nodemailer');

    const senderBook = xlsx.readFile(path.join(__dirname, 'emails', 'emails.xlsx'));
    const senderSheet = senderBook.SheetNames;
    let senderRow = xlsx.utils.sheet_to_json(senderBook.Sheets[senderSheet[0]]);

    const leadBook = xlsx.readFile(path.join(__dirname, 'emails', 'leads.xlsx'));
    const leadSheet = leadBook.SheetNames;
    let leadRow = xlsx.utils.sheet_to_json(leadBook.Sheets[leadSheet[0]]);

    var gmail, mailSendInterval, mailSubject, mailBody, sendCount = 50;

    if (!flag) {
        process.exit(1);
    }

    fs.readFile(path.join(__dirname, 'config.json'), (err, res) => {
        
        if (err) {
            console.log('config.json file not found');
            process.exit(1);
        }

        res = JSON.parse(res);

        mailSubject = res.mailSubject;
        mailBody = res.mailBody;
        sendCount = res.sendFromPerMail;
        mailSendInterval = res.mailSendInterval * 1000;

        senderRow.map(function(value) {
            let oAuth2Client = new google.auth.OAuth2(value.client_id, value.client_secret, value.redirect_uris);
            oAuth2Client.setCredentials({
                refresh_token   : value.refresh_token,
            });

            oAuth2Client.getAccessToken()
            .then(accessToken => {
                sendEmail(value, accessToken.token);
            })
            .catch(err => {
               //console.log('token err :'+err.message);
            })
        });
    });

    function sendEmail(senderInfo, accessToken) {
        
        var i = 0;

        if (!flag) {
            process.exit(1);
        }

        var sendEmailInterval = setInterval(() => {

            var leadMail = leadRow.pop();
           
            if ( typeof leadMail === "undefined" ) {
                console.log('All lead is completed to send email');
                process.exit(1);
            }

            let smtpTransport = nodemailer.createTransport({
                host: 'smtp.gmail.com',
                port: 465,
                secure: true,
                auth: {
                    type: "OAuth2",
                    user: senderInfo.email, 
                    clientId: senderInfo.client_id,
                    clientSecret: senderInfo.client_secret,
                    refreshToken: senderInfo.refresh_token,
                    accessToken: accessToken
                }
            });

            smtpTransport.sendMail({
                from: senderInfo.email,
                to: leadMail.lead,
                subject: mailSubject,
                text: mailBody
            })
            .then(result => {
                console.log(`From ${senderInfo.email} To ${leadMail.lead}`);
            })
            .catch(err => {
                console.log('sending err :'+err.message);
            })

            i++;

            i === sendCount && clearInterval(sendEmailInterval);

        }, mailSendInterval);
    }
})    
