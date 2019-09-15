const path = require('path');
const fs = require('fs');
var flag = false;

require('getmac').getMac(function(err, macAddress){
    if (err)  throw err;
    let computer = fs.readFileSync(path.join(__dirname, 'computerAuth.txt'), 'utf8');
   	if (macAddress === computer) {
   		flag = true;
   	} else {
   		process.exit(1);
   	}

   	const xlsx = require('xlsx');
	const {google} = require('googleapis');
	const readline = require('readline');

   	const SCOPES = ['https://mail.google.com/'];
	const credential = path.join(__dirname, 'emails', 'credentials.json');
	const TOKEN_PATH = 'token.json';

	var client_id, client_secret, redirect_uris, clientToken;


	if (!flag) {
	    process.exit(1);
	}


	fs.readFile(credential, (err, content) => {
	  if (err) return console.log('Error loading client secret file:', err);
	  authorize(JSON.parse(content), storeEmails);
	});


	function authorize(credentials, callback) {

	  client_id = credentials.installed.client_id;
	  client_secret = credentials.installed.client_secret;
	  redirect_uris = credentials.installed.redirect_uris[0];

	  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris);

	  fs.readFile(path.join(__dirname, 'emails', TOKEN_PATH), (err, token) => {
	    if (err) return getNewToken(oAuth2Client, callback);
	    clientToken = JSON.parse(token)
	    oAuth2Client.setCredentials(clientToken);
	    callback(oAuth2Client);
	  });
	}

	function getNewToken(oAuth2Client, callback) {
	  const authUrl = oAuth2Client.generateAuthUrl({
	    access_type: 'offline',
	    scope: SCOPES,
	  });
	  console.log('Authorize this app by visiting this url:', authUrl);
	  const rl = readline.createInterface({
	    input: process.stdin,
	    output: process.stdout,
	  });
	  rl.question('Enter the code from that page here: ', (code) => {
	    rl.close();
	    oAuth2Client.getToken(code, (err, token) => {
	      if (err) return console.error('Error retrieving access token', err);
	      oAuth2Client.setCredentials(token);
	      clientToken = token;
	      fs.writeFile(path.join(__dirname, 'emails', TOKEN_PATH), JSON.stringify(token), (err) => {
	        if (err) return console.error(err);
	        console.log('Token stored to');
	      });
	      callback(oAuth2Client);
	    });
	  });
	}

	function storeEmails(auth) {
	  
	  if (!flag) {
	      process.exit(1);
	  }

	  const gmail = google.gmail({version: 'v1', auth});
	  gmail.users.getProfile({
	    userId: 'me',
	  }, (err, res) => {
	    if (err) return console.log('The API returned an error: ' + err);
	    let info = {
	      'email': res.data.emailAddress,
	      'client_id': client_id,
	      'client_secret': client_secret,
	      'redirect_uris': redirect_uris,
	      'access_token': clientToken.access_token,
	      'refresh_token': clientToken.refresh_token,
	      'scope': clientToken.scope,
	      'token_type': clientToken.token_type,
	      'expiry_date': clientToken.expiry_date,
	    };

	    let workBook = xlsx.readFile(path.join(__dirname, 'emails', 'emails.xlsx'));
	    let workSheet = workBook.SheetNames;
	    let rows = xlsx.utils.sheet_to_json(workBook.Sheets[workSheet[0]]);
	    rows.push(info);
	    let newWB = xlsx.utils.book_new();
	    let xlsxData = xlsx.utils.json_to_sheet(rows);
	    xlsx.utils.book_append_sheet(newWB, xlsxData, 'email');
	    xlsx.writeFile(newWB, path.join(__dirname, 'emails', 'emails.xlsx'));

	    fs.unlink(path.join(__dirname, 'emails', 'credentials.json'), (err) => {
	      if (err) throw err;
	      console.log('credentials.json was deleted');
	    });

	    fs.unlink(path.join(__dirname, 'emails', 'token.json'), (err) => {
	      if (err) throw err;
	      console.log('token.json was deleted');
	    });
	  });
	}
});