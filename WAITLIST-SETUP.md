# Dossier waitlist → Google Sheets

## Why only timestamps showed up

The site was posting **JSON** (`text/plain`). Many waitlist scripts only read **`e.parameter`** (form fields). That writes `new Date()` for Timestamp and leaves Email / Location / Followers blank.

The row with `form@dossier...` worked because that test used form fields. The form now posts form fields again.

## Apps Script — paste this full script

```javascript
function doGet() {
  return ContentService
    .createTextOutput('ok')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var email = '';
  var location = '';
  var followers = '';
  var timestamp = new Date();

  // 1) Form fields (what the website sends)
  if (e && e.parameter) {
    email = e.parameter.email || '';
    location = e.parameter.location || '';
    followers = e.parameter.followers || '';
    if (e.parameter.timestamp) timestamp = e.parameter.timestamp;
  }

  // 2) JSON body fallback (terminal / older clients)
  try {
    if (e && e.postData && e.postData.contents) {
      var data = JSON.parse(e.postData.contents);
      email = data.email || email;
      location = data.location || location;
      followers = data.followers || followers;
      if (data.timestamp) timestamp = data.timestamp;
    }
  } catch (err) {}

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp', 'Email', 'Location', 'Followers']);
  }

  sheet.appendRow([timestamp, email, location, followers]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

Then: **Deploy → Manage deployments → Edit → New version → Deploy**.

## Local test

```bash
cd /Users/felisawiley/dossier
python3 -m http.server 8765
```

Open http://127.0.0.1:8765/creators.html and submit.

Terminal check (form fields, same as the site):

```bash
curl -sS -L --post302 -X POST \
  -d "email=check@dossier.test&location=Philly&followers=1200&timestamp=now" \
  "YOUR_EXEC_URL"
```
