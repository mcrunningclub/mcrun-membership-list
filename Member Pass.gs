function createMemberPass(passInfo) {
  const TEMPLATE_ID = '14NG31db-g-bFX1OUHeRByTKN6S2QuMAkDuANOAtwF6o';
  const FOLDER_ID = '1_NVOD_HbXfzPl26lC_-jjytzaWXqLxTn';

  // Get the template presentation
  const template = DriveApp.getFileById(TEMPLATE_ID);
  const passFolder = DriveApp.getFolderById(FOLDER_ID);

  // Use information to create custom file name
  const memberName = `${passInfo.firstName}-${passInfo.lastName}`;

  // Add formatted name to memberInfo
  const today = new Date();
  passInfo['name'] = `${passInfo.firstName} ${(passInfo.lastName).charAt(0)}.`;
  passInfo['generatedDate'] = Utilities.formatDate(today, TIMEZONE, 'MMM-dd-yyyy');
  passInfo['cYear'] = Utilities.formatDate(today, TIMEZONE, 'yyyy');

  // Make a copy to edit
  const copyRef = template.makeCopy(`${memberName}-pass-copy`, passFolder);
  const copyID = copyRef.getId();
  const copyFilePtr = SlidesApp.openById(copyID);

  // Replace placeholders with member data
  for (const [key, value] of Object.entries(passInfo)) {
    let placeHolder = `{{${key}}}`;
    copyFilePtr.replaceAllText(placeHolder, value);
  }
  
  // Open the presentation and get the first slide
  const slide = copyFilePtr.getSlides()[0];

  // Create QR code
  const qrCodeUrl = generateQrUrl(passInfo.aMemberId);
  const qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();

  // Find shape with placeholder alt text and replace with qr code
  const qrPlaceholder = '{{qrCodeImage}}';
  const images = slide.getImages();

  for (let image of images) {
    if (image.getDescription() === qrPlaceholder) {
      // Replace the placeholder image with the QR code image
      image.replace(qrCodeBlob);
      Logger.log(`QR Code placeholder replaced with the generated QR Code ${passInfo.memberID}.`);
      break;
    }
  }

  /*
  const shapes = slide.getShapes();

  for (let shape of shapes) {
    if (shape.getText().asString().includes(placeholder)) {
      // Generate the QR code image as a blob
      const qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();

      // Get the position and size of the placeholder
      const position = shape.getLeft();
      const top = shape.getTop();
      const width = shape.getWidth();
      const height = shape.getHeight();

      // Remove the text placeholder
      shape.getText().setText('');

      // Insert the QR code image in place of the placeholder
      slide.insertImage(qrCodeBlob, position, top, width, height);

      Logger.log('QR Code added to the slide.');
      break;
    }
  }
  */

  // Save anc close copy template
  copyFilePtr.saveAndClose();

  // Export the copy presentation as PNG
  const exportUrl = `https://docs.google.com/presentation/d/${copyID}/export/png`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
      muteHttpExceptions: true,
    },
  });

  // Save the PNG file to the folder
  const blob = response.getBlob();
  const fileDate = Utilities.formatDate(today, TIMEZONE, 'yyyyMMdd');
  passFolder.createFile(blob).setName(`${memberName}-McRun-Pass-${fileDate}.png`);

  // Moves the file to the trash
  copyRef.setTrashed(true);
}


function generateQrUrl(memberID) {
  const baseUrl = 'https://quickchart.io/qr?';
  const params = `text=${encodeURIComponent(memberID)}&margin=1&size=200`

  return baseUrl + params;
}


function start() {
  const sheet = MASTER_SHEET;
  const memberEmail = 'adela.cerna@mail.mcgill.ca';
  const memberRow = findMemberByEmail(memberEmail, sheet);
  
  const testInfo = {
    firstName : 'Benjamin',
    lastName : 'Higgins',
    memberID : '21a15ee377a022d7',
    memberStatus : 'Active',
    feeStatus : 'Paid',
    expiry : 'Feb 2025',
  }

  createMemberPass(testInfo);
}


//https://docs.google.com/presentation/d/10eZkny4yeuafoGnnkWG2VV4umKZ66CN4D9xV2A9J9-Y/export?format=png
//const slidesBlob = (tempDoc, 'application/vnd.google-apps.presentation');
//const pngBlob = slidesBlob.getAs('image/png');
//const tempCopy = template.makeCopy(tempName, passFolder);
//const tempDoc = SlidesApp.openById(tempCopy.getId());

function generateMemberIDCards_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members'); // Replace with your sheet name
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const qrCodeUrlColumn = header.indexOf('QR Code URL') + 1; // Adjust this if your sheet already has a 'QR Code URL' column
  
  // Create a folder for the member cards
  const cardsFolder = DriveApp.createFolder('Member ID Cards');

  // Iterate through member rows
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const firstName = row[header.indexOf('First Name')];
    const lastName = row[header.indexOf('Last Name')];
    const memberId = row[header.indexOf('Member ID')];
    const runPoints = row[header.indexOf('Run Points')];
    const tier = row[header.indexOf('Tier')];
    const validUntil = row[header.indexOf('Valid Until')];

    // Generate QR Code URL
    const qrCodeUrl = generateQRURL(memberId);
    //sheet.getRange(i + 1, qrCodeUrlColumn).setValue(qrCodeUrl); // Write QR Code URL to the sheet

    // Load HTML template and replace placeholders
    const template = HtmlService.createTemplateFromFile('IDCardTemplate');
    template.firstName = firstName;
    template.lastName = lastName;
    template.runPoints = runPoints;
    template.tier = tier;
    template.qrCodeUrl = qrCodeUrl;
    template.validUntil = validUntil;

    const htmlContent = template.evaluate().getContent();

    // Convert HTML to Blob and create PDF
    const blob = Utilities.newBlob(htmlContent, 'text/html', `${firstName}_${lastName}_IDCard.html`);
    const pdf = blob.getAs('application/pdf');
    cardsFolder.createFile(pdf).setName(`${firstName}_${lastName}_IDCard.pdf`);
  }

  //SpreadsheetApp.getUi().alert(`ID cards and QR codes have been created.`);
}



function generateMemberIDCards2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Members');
  const data = sheet.getDataRange().getValues();
  const header = data[0]; // Assuming the first row contains headers

  // Gets a folder in Google Drive
  const cardsFolder = DriveApp.getFolderById('1_NVOD_HbXfzPl26lC_-jjytzaWXqLxTn');

  // Mapping header indices
  const firstNameIndex = header.indexOf('First Name');
  const lastNameIndex = header.indexOf('Last Name');
  const memberIdIndex = header.indexOf('Member ID');
  const runPointsIndex = header.indexOf('Run Points');
  const tierIndex = header.indexOf('Tier');
  const validUntilIndex = header.indexOf('Valid Until');


  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const firstName = row[firstNameIndex];
    const lastName = row[lastNameIndex];
    const memberId = row[memberIdIndex];
    const runPoints = row[runPointsIndex];
    const tier = row[tierIndex];
    const validUntil = row[validUntilIndex];

    // Generate QR Code URL
    const fileId = '1z1bsjoVoIQfQzTj1h5FNSTt84GS3rtXn' 
    //'https://drive.google.com/uc?export=download&id=1z1bsjoVoIQfQzTj1h5FNSTt84GS3rtXn' //generateQRCode2(memberId);

    const qrCodeUrl = loadImageBytes(fileId);
    //sheet.getRange(i + 1, qrCodeUrlColumn).setValue(qrCodeUrl); // Write QR Code URL to the sheet
    
    // Load HTML template and replace placeholders
    const template = HtmlService.createTemplateFromFile('memberIDTemplate');
    template.name = `${firstName} ${lastName}`;
    template.runPoints = runPoints;
    template.tier = tier;
    //template.qrCodeUrl = 'data:image/png;base64,' + `${qrCodeUrl}`;
    template.validUntil = validUntil;

    const htmlContent = template.evaluate().getContent();
    Logger.log(htmlContent);
    
    // Convert HTML to a Blob and create a PDF
    const blob = Utilities.newBlob(htmlContent, 'text/html', `${firstName}_${lastName}_IDCard.html`);
    const pdf = blob.getAs('application/pdf');
    cardsFolder.createFile(pdf).setName(`${firstName}_${lastName}_IDCard_6.pdf`);
  }
  
  //SpreadsheetApp.getUi().alert(`ID cards have been created in the folder: ${cardsFolder.getName()}`);
}

function generateQRCode2(memberID) {
  const baseUrl = 'https://quickchart.io/qr?';
  const bottomText = "Valid until 2023"
  const params = `text=${encodeURIComponent(memberID)}&margin=1&size=200`

  //const params = `text=${encodeURIComponent(memberID)}&margin=1&size=200&caption=${encodeURIComponent(`${bottomText}`)}&captionFontFamily=mono&captionFontSize=11`

  return baseUrl + params;
}

function test() {
  const fileId = '1zIQfQzTj1h5FNSTttXn';
  const encoded = loadImageBytes(fileId);
  const res = 'data:image/png;base64,' + encoded;

  Logger.log(res) ;

  //"https://quickchart.io/qr?text=hello$margin=1$size=300"
}


function getImage(url) {
  var image = UrlFetchApp.fetch(url).getAs('image/png')
  var response = UrlFetchApp.fetch(url).getResponseCode();
  if (response === 200) {
    var img = UrlFetchApp.fetch(url).getAs('image/png');
  }
  return img;
}


function loadImageBytes(id){
  var bytes = DriveApp.getFileById(id).getBlob().getBytes();
  return Utilities.base64Encode(bytes);
}


