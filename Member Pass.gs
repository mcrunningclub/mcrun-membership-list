function createPassFile(passInfo) {
  const TEMPLATE_ID = '14NG31db-g-bFX1OUHeRByTKN6S2QuMAkDuANOAtwF6o';   // Is not confidential
  const FOLDER_ID = '1_NVOD_HbXfzPl26lC_-jjytzaWXqLxTn';    // Is not confidential either

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
  const fileDate = Utilities.formatDate(today, TIMEZONE, 'yyyyMMdd');
  const copyRef = template.makeCopy(`${memberName}-McRun-Pass-${fileDate}`, passFolder);
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
  const qrCodeUrl = generateQrUrl(passInfo.memberID);
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

  // Save anc close copy template
  copyFilePtr.saveAndClose();

  // Set permissions to general to allow downloading
  copyRef.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);

  // Create download link for member
  return `https://docs.google.com/presentation/d/${copyID}/export/png`;   // Download link for user


  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
      muteHttpExceptions: true,
    },
  });

  // Save the PNG file to the folder
  const blob = response.getBlob();
  //const fileDate = Utilities.formatDate(today, TIMEZONE, 'yyyyMMdd');
  const file = passFolder.createFile(blob).setName(`${memberName}-McRun-Pass-${fileDate}.png`);
  
  // Set permissions to general to allow downloading
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  

  // Moves the file to the trash
  copyRef.setTrashed(true);
  return file.getDownloadUrl();
}


function testPass() {
  /**
   * 3:46:30 AM	Info	https://docs.google.com/presentation/d/1zYTdKOvjqb90KwjhmcZxiqS-UlnWFO61EDZIqxhm-VQ/export/png
   * 3:46:35 AM	Info	https://drive.google.com/uc?id=1aLeRsQ25FK3IDBWuUn2ym4TGSGrMNpMA&export=download
   * 
   * 4:13:29 AM	Info	https://docs.google.com/presentation/d/1SxrZtzp_LaqBjVdcfGjoxrt7t-1Hggdwl0gewWpC_Mw/export/png
   * 4:13:30 AM	Info	Direct Download Link: https://drive.google.com/uc?export=download&id=1SxrZtzp_LaqBjVdcfGjoxrt7t-1Hggdwl0gewWpC_Mw
   */
}


function testRuntime() {
  const memberEmail = 'emerson.darling@mail.mcgill.ca';
  const startTime = new Date().getTime();

  // Runtime if creating and sharing as png: 10351 ms
  // If only creating Slides file and sharing download link: 5906 ms
  const url = generateMemberPass(memberEmail);
  console.log(url);
  
  // Record the end time
  const endTime = new Date().getTime();
  
  // Calculate the runtime in milliseconds
  const runtime = endTime - startTime;
  
  // Log the runtime
  Logger.log(`Function runtime: ${runtime} ms`);
}


function generateMemberPass(memberEmail) {
  const sheet = MASTER_SHEET;
  const row = findMemberByEmail(memberEmail, sheet);

  // Error management
  if (isNaN(row)) {
    throw Error(`Member email (${memberEmail}) could not be found in MASTER`);
  }

  // Get member data to populate pass template
  const endCol = MASTER_MEMBER_ID_COL;   // From first name to id col
  const memberData = sheet.getRange(row, 1, 1, endCol).getValues()[0];

  // Add entry to beginning to allow 1-indexed data access like for GSheet
  memberData.unshift('');

  // Get membership expiration date
  const membershipExpiration = getExpirationDate(memberData[MASTER_LAST_REG_SEM]);

  // Map member info to pass info
  const passInfo = {
    firstName: memberData[MASTER_FIRST_NAME_COL],
    lastName: memberData[MASTER_LAST_NAME_COL],
    memberID: memberData[MASTER_MEMBER_ID_COL],
    memberStatus: 'Active',    // If email not found, then membership expired
    feeStatus: memberData[MASTER_FEE_STATUS],
    expiry: membershipExpiration,
  }

  return createPassFile(passInfo);    // Get download url for member
}


function generateQrUrl(memberID) {
  const baseUrl = 'https://quickchart.io/qr?';
  const params = `text=${encodeURIComponent(memberID)}&margin=1&size=200`

  return baseUrl + params;
}


function getExpirationDate(semCode) {
  const validDuration = 1;    // 1 year

  const semester = semCode.charAt(0);
  const expirationYear = '20' + (parseInt(semCode.slice(-2)) + validDuration)

  switch (semester) {
    case ('F'): return `Sep ${expirationYear}`;
    case ('W'): return `Jan ${expirationYear}`;
    case ('S'): return `Jun ${expirationYear}`;
    default: return null;
  };

}


function getImage(url) {
  var response = UrlFetchApp.fetch(url).getResponseCode();
  if (response === 200) {
    var img = UrlFetchApp.fetch(url).getAs('image/png');
  }
  return img;
}


function loadImageBytes(id) {
  var bytes = DriveApp.getFileById(id).getBlob().getBytes();
  return Utilities.base64Encode(bytes);
}


function testQRGenerator() {
  const FOLDER_ID = '1_NVOD_HbXfzPl26lC_-jjytzaWXqLxTn';
  const passFolder = DriveApp.getFolderById(FOLDER_ID);

  const memberId = '1zIQfQzTj1h5FNSTttXn';

  const qrCodeUrl = generateQrUrl(memberId);
  const qrCodeBlob = UrlFetchApp.fetch(qrCodeUrl).getBlob();

  passFolder.createFile(qrCodeBlob).setName(`QRCode-test.png`);
}


