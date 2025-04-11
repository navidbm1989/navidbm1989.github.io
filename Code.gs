/**
 * Sacred Needles Consent Form Web App - Server-Side Script
 *
 * Overview:
 * This Google Apps Script serves as the backend for a tattoo consent form web app for Sacred Needles LLC.
 * It collects client personal info, tattoo details, and consent signatures through a multi-step process,
 * generating a PDF, saving it to Google Drive, logging submissions in a Google Sheet, and sending emails.
 *
 * Structure:
 * - Welcome Page: Introduces the process.
 * - Personal Info Page: Collects name, email, phone, gender, address, DOB, ID type, and photo ID uploads.
 * - Tattoo Description Page: Gathers placement, size, and ink selection.
 * - Signature Preference Page: Offers font adoption or manual signature options.
 * - Consent Form Page: Displays consent terms and collects initials/signatures.
 * - Thank You Page: Confirms submission with social media links.
 *
 * Inputs:
 * - Personal: Full Name, Email, Phone, Gender, Address (Street, City, State, Zip), DOB, ID Type, Photo IDs
 * - Tattoo: Placement, Size, Ink Selection
 * - Consent: Signature Method (Adopt/Manual), Initials (14 statements), Signature, Date
 *
 * Intentions:
 * 1. Legal Compliance: Ensures clients acknowledge risks and waive liability.
 * 2. Data Collection: Records client and tattoo info for legal/business use.
 * 3. User Experience: Guides clients through a validated, step-by-step process.
 * 4. Automation: Creates PDFs, saves to Drive, logs in Sheets, and emails automatically.
 *
 * Functions:
 * - doGet(): Serves the HTML page.
 * - getConfigData(): Fetches config from a Google Doc (e.g., logo, emails).
 * - submitForm(data): Processes submissions, generates PDFs, saves/logs data, sends emails.
 * - Helper Functions: extractGoogleDriveFileId(), extractSheetIdFromUrl(), getSocialMediaLinks(),
 *   getUniqueFileName(), getOrdinalSuffix(), dataURLtoBlob(), scaleImage()
 *
 * Customization Instructions:
 * 1. Update Config: Replace Google Doc ID in getConfigData() with new client’s config doc ID.
 * 2. Branding: Update logo URL and business name in config doc.
 * 3. Consent Terms: Modify consentStatements array in submitForm() for client-specific terms.
 * 4. Social Media: Adjust links in config doc.
 * 5. Test: Verify functionality with new client data.
 */
// Serve the HTML page as a web app
function doGet() {
  // Fetch config data (assumes getConfigData() retrieves data from a Google Doc or similar source)
  var config = getConfigData();
  
  // Create an HTML template from the 'index' file
  var template = HtmlService.createTemplateFromFile('index');
  
  // Assign the logo URL from config to the template
  template.logoUrl = config.LogoImage;
  
  // Evaluate and serve the template with additional settings
  return template.evaluate()
    .setTitle('Sacred Needles Consent Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fetch configuration data from a Google Doc
function getConfigData() {
  Logger.log('before Opening config data')
  //https://docs.google.com/document/d/1VOBavb7IRE7yr2uCu2Tx4QwyKNact5uwDbo03HTxWUw/edit?tab=t.0
  var docId = '1VOBavb7IRE7yr2uCu2Tx4QwyKNact5uwDbo03HTxWUw'; // Replace with your Google Doc ID
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var tables = body.getTables();
  var config = {};
Logger.log('Opened config data')

      if (tables.length > 0) {
        var table = tables[0];
        for (var i = 0; i < table.getNumRows(); i++) {
          var row = table.getRow(i);
          var key = row.getCell(0).getText();
          var value = row.getCell(1).getText();
          
          /*// Convert Google Drive view links to direct download links
          if (value.includes('drive.google.com/file/d/')) {
            var fileId = extractGoogleDriveFileId(value);
            if (fileId) {
              value = `https://drive.google.com/uc?export=download&id=${fileId}`;
              Logger.log(`Converted ${key} URL to: ${value}`);
            }
          }*/

          // Convert Google Drive view links to thumbnail URLs
          if (value.includes('drive.google.com/file/d/')) {
            var fileId = extractGoogleDriveFileId(value);
            if (fileId) {
              value = `https://drive.google.com/thumbnail?id=${fileId}`;
              Logger.log(`Converted ${key} URL to thumbnail: ${value}`);
            }
          }
          
          config[key] = value;
          if (key === 'ArtistSignature') {
            var textElement = row.getCell(1).editAsText();
            var font = textElement.getFontFamily(0) || 'Times New Roman';
            config['ArtistSignatureFont'] = font;
          }
        }
      }
      Logger.log('Config Data Fetched: ' + JSON.stringify(config));
      return config;
    }
    // Helper function to extract Google Drive file ID from a view link
      function extractGoogleDriveFileId(url) {
        var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          return match[1];
        }
        Logger.log('Could not extract file ID from URL: ' + url);
        return null;
      }

  // New helper function to extract Google Sheet ID from URL
    function extractSheetIdFromUrl(url) {
      try {
        Logger.log('Extracting Sheet ID from URL: ' + url); // Log the input URL
        if (!url || typeof url !== 'string') {
          throw new Error('Invalid or missing Google Sheet URL');
        }
        var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
          Logger.log('Extracted Sheet ID: ' + match[1]);
          return match[1];
        }
        throw new Error('Could not extract Sheet ID from URL: ' + url);
      } catch (error) {
        Logger.log('Error extracting Sheet ID: ' + error.message);
        throw error;
      }
    }

        // Function to fetch social media links for the Thank You page
        function getSocialMediaLinks() {
          var config = getConfigData();
          return {
            instagramHandle: config.InstagramHandle || '@sacred.needles',
            instagramLink: `https://www.instagram.com/${(config.InstagramHandle || '@sacred.needles').replace('@', '')}/`,
            facebookPage: config.FacebookPage || 'sacred.needlesfb',
            facebookLink: `https://www.facebook.com/${(config.FacebookPage || 'sacred.needlesfb').replace(/\s+/g, '')}/`
          };
        }
// Process form submission, generate PDF, save to Drive, update Sheet, send emails
function submitForm(data) {
      try {
        // Check email quota before proceeding
        var quotaRemaining = MailApp.getRemainingDailyQuota();
        Logger.log('Email quota remaining before submission: ' + quotaRemaining);
        if (quotaRemaining < 5) {
          Logger.log('Quota too low: ' + quotaRemaining + ' emails remaining. Submission aborted.');
          return 'quota_exceeded'; // Signal to client
        }

        // Fetch configuration data
        var config = getConfigData();

        // Extract and sanitize form data
        var fullName = data.fullName || 'N/A';
        var email = data.email || 'N/A';
        var phone = data.phone || 'N/A';
        var gender = data.gender || 'N/A';
        var address = data.address ? `${data.address.street}, ${data.address.city}, ${data.address.state} ${data.address.zipcode}` : 'N/A';
        var dob = data.dob || 'N/A';
        var age = data.age || 'N/A';
        var idType = data.idType || 'N/A';
        var tattooDetails = data.tattooDetails || {};
        var placement = tattooDetails.placement || 'N/A';
        var size = tattooDetails.size || 'N/A';
        var inkSelection = tattooDetails.inkSelection || 'N/A';
        var signatureMethod = data.signatureMethod || 'N/A';
        var selectedFont = data.selectedFont || '';
        var date = data.date || 'N/A';
        var submissionDate = new Date().toLocaleString();
        var initials = data.initials || Array(14).fill('N/A');
        var clientSignature = data.signature || 'N/A';
        var frontIdBlob = data.frontId ? dataURLtoBlob(data.frontId) : null;
        var backIdBlob = data.backId ? dataURLtoBlob(data.backId) : null;

        // Extract the client's first name from fullName
        var firstName = fullName.split(' ')[0] || 'Dear Recipient'; // Fallback to 'Client' if fullName is 'N/A' or empty

         // File upload validation
        var maxFileSize = 10 * 1024 * 1024; // 10MB in bytes
        var validImageTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/bmp', 'image/heic', 'image/heif'];

        // Validate front ID
        if (!frontIdBlob) {
          throw new Error('Front ID is required.');
        }
        // Check file size
        if (frontIdBlob.getBytes().length > maxFileSize) {
          throw new Error('Front ID file size exceeds 10MB limit.');
        }
        // Check MIME type
        var frontMimeType = frontIdBlob.getContentType();
        if (!validImageTypes.includes(frontMimeType)) {
          throw new Error('Front ID must be a valid image (JPEG, PNG, GIF, BMP, HEIC).');
        }
        // Check file signature
        if (!isValidImage(frontIdBlob, frontMimeType)) {
          throw new Error('Front ID file is not a valid image based on content.');
        }

        // Validate back ID
        if (!backIdBlob) {
          throw new Error('Back ID is required.');
        }
        // Check file size
        if (backIdBlob.getBytes().length > maxFileSize) {
          throw new Error('Back ID file size exceeds 10MB limit.');
        }
        // Check MIME type
        var backMimeType = backIdBlob.getContentType();
        if (!validImageTypes.includes(backMimeType)) {
          throw new Error('Back ID must be a valid image (JPEG, PNG, GIF, BMP, HEIC).');
        }
        // Check file signature
        if (!isValidImage(backIdBlob, backMimeType)) {
          throw new Error('Back ID file is not a valid image based on content.');
        }

        // Determine ink value from config based on selection
        var inkValue;
        if (inkSelection === 'Black & Gray Ink') {
          inkValue = config.InkOptionBlackGray || 'N/A';
        } else if (inkSelection === 'Colorwork Ink') {
          inkValue = config.InkOptionColorwork || 'N/A';
        } else if (inkSelection === 'Both') {
          var blackGray = config.InkOptionBlackGray || 'N/A';
          var colorwork = config.InkOptionColorwork || 'N/A';
          inkSelection = 'Black&Gray and Colorwork Ink'
          inkValue = `${blackGray}, ${colorwork}`; // Concatenate with a comma
        } else {
          inkValue = 'N/A';
        }

        // Format date for display
        var formattedDate = date ? Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'N/A';
        var formattedDOB = date ? Utilities.formatDate(new Date(dob), Session.getScriptTimeZone(), 'MM/dd/yyyy') : 'N/A';
        // Log submitted data for debugging
        Logger.log('Submitted Data: ' + JSON.stringify({
          fullName: fullName,
          email: email,
          phone: phone,
          gender: gender,
          address: address,
          dob: formattedDOB,
          age: age,
          idType: idType,
          placement: placement,
          size: size,
          inkSelection: inkSelection,
          inkValue: inkValue,
          signatureMethod: signatureMethod,
          date: date,
          formattedDate: formattedDate,
          submissionDate: submissionDate,
          initials: initials,
          signature: clientSignature
        }));
       Logger.log('After Logging Submited Data');

        // Create a new Google Doc for the consent form
        var baseFileName = `Consent-Form-${fullName}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd-yyyy')}`;
        var doc = DocumentApp.create(baseFileName);
        var body = doc.getBody();
        body.editAsText().setFontFamily('Times New Roman');
         Logger.log('Base File Name: ' +baseFileName);

        // Add logo at the top
        var logoResponse = UrlFetchApp.fetch(config.LogoImage);
        var logoBlob = logoResponse.getBlob();
        var logoParagraph = body.appendParagraph('');
        var logoImage = logoParagraph.appendInlineImage(logoBlob);
        scaleImage(logoImage, 100);
        logoParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        Logger.log('Logo Was added to Top');

        // Add title and subtitle
        var title = body.appendParagraph('Sacred Needles LLC');
        title.setHeading(DocumentApp.ParagraphHeading.TITLE);
        title.editAsText().setFontFamily('Times New Roman').setBold(true).setFontSize(36);
        title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        var subtitle = body.appendParagraph('TATTOO CONSENT FORM');
        subtitle.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
        subtitle.editAsText().setFontFamily('Georgia').setItalic(true).setFontSize(24);
        subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        body.appendParagraph('');
        Logger.log('Title and SubTtile added');
        
        // Add introductory paragraph
        var introParagraph = body.appendParagraph('In consideration of receiving a tattoo from SACRED NEEDLES LLC, including its artists, associates, apprentices, agents, or any employees (hereinafter referred to as the “Tattoo Studio”), I agree to the following:');
        introParagraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        introParagraph.editAsText().setFontFamily('Times New Roman').setItalic(false).setFontSize(12);
        body.appendParagraph('');

        // Add consent statements with initials
        var consentStatements = [
          `I, ${fullName}, have been fully informed of the inherent risks associated with getting a tattoo. Therefore, I fully understand that these risks, known and unknown, can lead to injury including but not limited to: infection, scarring, difficulties in the detection of melanoma and allergic reactions to tattoo pigment, latex gloves and/or soap. Having been informed of the potential risks associated with getting a tattoo I wish to proceed with the tattoo procedure and application and freely accept and expressly assume any and all risks that may arise from tattooing.`,
          `I WAIVE AND RELEASE to the fullest extent permitted by law any person of the Tattoo Studio from all liability whatsoever, including but not limited to, any and all claims or causes of action that I, my estate, heirs, executors or assigns may have for personal injury or otherwise, including any direct and/or consequential damages, which result or arise from the procedure and application of my tattoo, whether caused by the negligence or fault of either the Tattoo Studio, or otherwise.`,
          `The Tattoo Studio has given me the full opportunity to ask any question about the procedure and application of my tattoo and all of my questions, if any, have been answered to my total satisfaction.`,
          `The Tattoo Studio has given me instructions on the care of my tattoo while it's healing. I understand and will follow them. I acknowledge that it is possible that the tattoo can become infected, particularly if I do not follow the instructions given to me.`,
          `I am not under the influence of alcohol or drugs, and I am voluntarily submitting to be tattooed by the Tattoo Studio without duress or coercion.`,
          `I do not suffer from diabetes, epilepsy, hemophilia, heart condition(s), nor do I take blood thinning medication. I do not have any other medical or skin condition that may interfere with the procedure, application or healing of the tattoo. I am not the recipient of an organ or bone marrow transplant or, if I am, I have taken the prescribed preventative regimen of antibiotics that is required by my doctor in advance of any invasive procedure such as tattooing. I am not pregnant or nursing. I do not have a mental impairment that may affect my judgement in getting the tattoo.`,
          `The Tattoo Studio is not responsible for the meaning or spelling of the symbol or text that I have provided to them or chosen from the flash (design) sheets.`,
          `Variations in color and design may exist between the tattoo art I have selected and the actual tattoo when it is applied to my body. I also understand that over time, the colors and the clarity of my tattoo will fade due to unprotected exposure to the sun and the naturally occurring dispersion of pigment under the skin.`,
          `A tattoo is a permanent change to my appearance and can only be removed by laser or surgical means, which can be disfiguring and/or costly and which in all likelihood will not result in the restoration of my skin to its exact appearance before being tattooed.`,
          `I release the right to any photographs taken of me and the tattoo and give consent in advance to their reproduction in print or electronic form. (For assurance, if you do not initial this provision, please inform the Tattoo Studio NOT to take any pictures of you and your completed tattoo).`,
          `I agree that the Tattoo Studio has a NO REFUND policy on tattoos, piercing and/or retail sales and I will not ask for a refund for any reason whatsoever.`,
          `I agree to reimburse the Tattoo Studio for any attorneys' fees and costs incurred in any legal action I bring against the Tattoo Studio and in which either the Artist of the Tattoo Studio is the prevailing party. I agree that the courts located in the County of Harris within the State of Texas shall have jurisdiction and venue over me and shall have exclusive jurisdiction for the purposes of litigating any dispute arising out of or related to this agreement.`,
          `I acknowledge that I have been given adequate opportunity to read and understand this document that it was not presented to me at the last minute and grasp that I am signing a legal contract waiving certain rights to recover damages against the Tattoo Studio. If any provision, section, subsection, clause or phrase of this release is found to be unenforceable or invalid, that portion shall be severed from this contract. The remainder of this contract will then be construed as though the unenforceable portion had never been contained in this document. I hereby declare that I am of legal age (and have provided valid proof of age and identification) and am competent to sign this Agreement.`,
          `I consent to the use of electronic signatures and records.`
        ];

        for (var i = 0; i < consentStatements.length; i++) {
          var initialValue = initials[i] || 'N/A';
          var paragraph = body.appendParagraph(`[${initialValue}] - ${consentStatements[i]}`);
          paragraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
          paragraph.editAsText().setFontFamily('Times New Roman').setItalic(false).setFontSize(12);
          if (signatureMethod === 'adopt' && initialValue !== 'N/A') {
            var text = paragraph.editAsText();
            text.setFontFamily(0, `[${initialValue}]`.length - 1, selectedFont);
          }
        }
        body.appendParagraph('');
        Logger.log('Consent Items Added');
        // Add additional statements
        var careInstructions = body.appendParagraph('I have received a copy of applicable written care instructions and I have read and understand such written care instructions.');
        careInstructions.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        careInstructions.editAsText().setFontFamily('Times New Roman').setItalic(false).setFontSize(12);
        body.appendParagraph('');
        var agreementStatement = body.appendParagraph('I HAVE READ THE AGREEMENT, I UNDERSTAND IT, AND I AGREE TO BE BOUND BY IT.');
        agreementStatement.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        agreementStatement.editAsText().setFontFamily('Times New Roman').setItalic(false).setFontSize(12);
        body.appendParagraph('');
         Logger.log('additional statements Added');
        // Client section
        var clientHeader = body.appendParagraph('Recipient');
        clientHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        clientHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        clientHeader.editAsText().setFontFamily('Times New Roman').setBold(true);
        // Client Name
        var clientNameParagraph = body.appendParagraph('Full Name: ');
        clientNameParagraph.appendText(fullName || 'N/A');
        clientNameParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        clientNameParagraph.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);
        Logger.log('Client Name: '+fullName+' Added');
        // Client Signature Top
        var clientSigParagraph = body.appendParagraph('Signature: ');
        clientSigParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        clientSigParagraph.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);

        // Client Signature
        var clientSignatureParagraph = body.appendParagraph('');
        Logger.log('Processing client signature - Method: ' + signatureMethod + ', Signature: ' + clientSignature);
        if (signatureMethod === 'manual' && clientSignature && clientSignature.startsWith('data:image/')) {
          try {
            Logger.log('Inserting manual signature image');
            var clientImage = clientSignatureParagraph.appendInlineImage(dataURLtoBlob(clientSignature));
            scaleImage(clientImage, 600);
            Logger.log('Manual signature image inserted successfully');
          } catch (e) {
            Logger.log('Error inserting manual signature image: ' + e.message);
            clientSignatureParagraph.appendText('Signature Image Failed to Load');
          }
        } else if (signatureMethod === 'adopt' && clientSignature !== 'N/A') {
          Logger.log('Inserting adopted signature text with font: ' + selectedFont);
          var clientSignatureLabel = clientSignatureParagraph.appendText(clientSignature);
          clientSignatureLabel.setBold(false);
          clientSignatureParagraph.editAsText().setFontSize(30).setFontFamily(selectedFont || 'Times New Roman'); // Fallback font
          Logger.log('Adopted signature text inserted');
        } else {
          Logger.log('No valid signature provided');
          clientSignatureParagraph.appendText('N/A');
        }
        clientSignatureParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);

        // Client Date
        var clientDateParagraph = body.appendParagraph('Date: ');
        clientDateParagraph.appendText(formattedDate);
        clientDateParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        clientDateParagraph.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);
        body.appendParagraph('');

        // Artist section
        var artistHeader = body.appendParagraph('Artist');
        artistHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        artistHeader.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        artistHeader.editAsText().setFontFamily('Times New Roman').setBold(true);

        // Artist Name
        var artistNameParagraph = body.appendParagraph('Full Name: ');
        artistNameParagraph.appendText(config.ArtistName || 'N/A');
        artistNameParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        artistNameParagraph.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);

        // Artist Signature Top
        var artistSigParagraph = body.appendParagraph('Signature: ');
        artistSigParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        artistSigParagraph.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);

        // Artist Signature
        var artistSignatureParagraph = body.appendParagraph('');
        var artistSignatureLabel = artistSignatureParagraph.appendText(config.ArtistSignature || 'N/A');
        artistSignatureLabel.setBold(false);
        artistSignatureParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        artistSignatureParagraph.editAsText().setFontSize(30);
        artistSignatureParagraph.editAsText().setFontFamily(config.ArtistSignatureFont);

        // Artist Date
        var artistDateParagraph = body.appendParagraph('Date: ');
        artistDateParagraph.appendText(formattedDate);
        artistDateParagraph.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
        artistDateParagraph.editAsText().setFontFamily('Times New Roman').setFontSize(12).setBold(false);
        body.appendParagraph('');

        // Personal Information section (new page)
        body.appendPageBreak();
        var personalInfoHeader = body.appendParagraph('Recipient Information:');
        personalInfoHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        personalInfoHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        var personalInfoFullName = body.appendParagraph(`Full Name: ${fullName}`);
        personalInfoFullName.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoFullName.editAsText().setFontFamily('Times New Roman');
        var personalInfoEmail = body.appendParagraph(`Email: ${email}`);
        personalInfoEmail.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoEmail.editAsText().setFontFamily('Times New Roman');
        var personalInfoPhone = body.appendParagraph(`Phone: ${phone}`);
        personalInfoPhone.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoPhone.editAsText().setFontFamily('Times New Roman');
        var personalInfoGender = body.appendParagraph(`Gender: ${gender}`);
        personalInfoGender.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoGender.editAsText().setFontFamily('Times New Roman');
        var personalInfoAddress = body.appendParagraph(`Address: ${address}`);
        personalInfoAddress.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoAddress.editAsText().setFontFamily('Times New Roman');
        var personalInfoDob = body.appendParagraph(`Date of Birth: ${formattedDOB}`);
        personalInfoDob.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoDob.editAsText().setFontFamily('Times New Roman');
        var personalInfoAge = body.appendParagraph(`Age: ${age}`);
        personalInfoAge.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoAge.editAsText().setFontFamily('Times New Roman');
        var personalInfoIdType = body.appendParagraph(`Form of ID: ${idType}`);
        personalInfoIdType.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        personalInfoIdType.editAsText().setFontFamily('Times New Roman');
        body.appendParagraph('');

        // Tattoo Description section
        var tattooDescHeader = body.appendParagraph('Tattoo Description:');
        tattooDescHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        tattooDescHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        var tattooDescPlacement = body.appendParagraph(`Placement: ${placement}`);
        tattooDescPlacement.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        tattooDescPlacement.editAsText().setFontFamily('Times New Roman');
        var tattooDescSize = body.appendParagraph(`Size: ${size}`);
        tattooDescSize.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        tattooDescSize.editAsText().setFontFamily('Times New Roman');

        var tattooDescInk = body.appendParagraph('Ink Selection: ');
        tattooDescInk.appendText(inkSelection);
        tattooDescInk.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        tattooDescInk.editAsText().setFontFamily('Times New Roman');

        var tattooDescInkBrand = body.appendParagraph('Ink Brands/SKU#: ');
        tattooDescInkBrand.appendText(inkValue);
        tattooDescInkBrand.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
        tattooDescInkBrand.editAsText().setFontFamily('Times New Roman');
        body.appendParagraph('');

        // Photo ID section
        var idImagesHeader = body.appendParagraph('Photo ID:');
        idImagesHeader.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        idImagesHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
        if (frontIdBlob) {
          var frontIdLabel = body.appendParagraph('Front of Photo ID:');
          frontIdLabel.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
          frontIdLabel.editAsText().setFontFamily('Times New Roman');
          var frontIdParagraph = body.appendParagraph('');
          try {
            var frontImage = frontIdParagraph.appendInlineImage(frontIdBlob);
            scaleImage(frontImage, 300);
            Logger.log('Front Photo ID image inserted successfully');
          } catch (e) {
            Logger.log('Error inserting front Photo ID image: ' + e.message);
          }
          body.appendParagraph('');
        }
        if (backIdBlob) {
          var backIdLabel = body.appendParagraph('Back of Photo ID:');
          backIdLabel.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
          backIdLabel.editAsText().setFontFamily('Times New Roman');
          var backIdParagraph = body.appendParagraph('');
          try {
            var backImage = backIdParagraph.appendInlineImage(backIdBlob);
            scaleImage(backImage, 300);
            Logger.log('Back Photo ID image inserted successfully');
          } catch (e) {
            Logger.log('Error inserting back Photo ID image: ' + e.message);
          }
        }

        // Save document and convert to PDF
        doc.saveAndClose();
        var pdfBlob = doc.getAs('application/pdf');

        // Save PDF to Google Drive
        var folderName = 'Sacred Needles Consent Forms';
        var folders = DriveApp.getFoldersByName(folderName);
        var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

        // Check for existing files and generate a unique name
        var uniqueFileName = getUniqueFileName(folder, baseFileName, '.pdf');
        var pdfFile = folder.createFile(pdfBlob).setName(uniqueFileName);
        var pdfUrl = pdfFile.getUrl();

        // Update Google Sheet using URL from config
          var sheetUrl = config.GoogleSheetURL; // Changed from GoogleSheetID to GoogleSheetURL
          Logger.log('Sheet URL from config: ' + sheetUrl); // Log the URL before extraction
          var sheetId = extractSheetIdFromUrl(sheetUrl); // Extract the ID from the URL
          Logger.log('Sheet ID: ' + sheetId); // Log Sheet ID
          var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
          if (sheet.getLastRow() === 0) {
            sheet.appendRow(['Submission Date', 'Full Name', 'Email', 'Phone', 'Gender', 'Address', 'Date of Birth', 'Age', 'Form of ID', 'Placement', 'Size', 'Ink Selection', 'Ink Brand and SKU#', 'PDF URL']);
          }
          sheet.appendRow([submissionDate, fullName, email, phone, gender, address, formattedDOB, age, idType, placement, size, inkSelection, inkValue, pdfUrl]);

        // Send admin email with PDF link
        var adminSubject = `New Consent Form Submission ${fullName}-${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM-dd-yyyy')}`;
        var adminBody = `
          <h3>New Submission</h3>
          <table border="1" cellpadding="5">
            <tr><td>Submission Date</td><td>${submissionDate}</td></tr>
            <tr><td>Full Name</td><td>${fullName}</td></tr>
            <tr><td>Email</td><td>${email}</td></tr>
            <tr><td>Phone</td><td>${phone}</td></tr>
            <tr><td>Gender</td><td>${gender}</td></tr>
            <tr><td>Address</td><td>${address}</td></tr>
            <tr><td>Date of Birth</td><td>${formattedDOB}</td></tr>
            <tr><td>Age</td><td>${age}</td></tr>
            <tr><td>Form of ID</td><td>${idType}</td></tr>
            <tr><td>Placement</td><td>${placement}</td></tr>
            <tr><td>Size</td><td>${size}</td></tr>
            <tr><td>Ink Selection</td><td>${inkSelection}</td></tr>
            <tr><td>Ink Brand and SKU#</td><td>${inkValue}</td></tr>
            <tr><td>PDF URL</td><td><a href="${pdfUrl}">View PDF</a></td></tr>
          </table>
        `;
          // Get or create the Gmail folder (label) from config
          var gmailFolderName = config.GmailFolderName || 'ConsentForms'; // Default to 'ConsentForms' if not specified
          var label = GmailApp.getUserLabelByName(gmailFolderName);
          if (!label) {
            label = GmailApp.createLabel(gmailFolderName);
            Logger.log('Created new Gmail folder: ' + gmailFolderName);
          }

        // Send the email
        MailApp.sendEmail({
          to: config.AdminEmail,
          replyTo: email,
          subject: adminSubject,
          htmlBody: adminBody
        });

        // Wait for Gmail to process it
        Utilities.sleep(3000); // 3-second delay

        // Search for the email
        var threads = GmailApp.search(`to:${config.AdminEmail} ${adminSubject}`, 0, 1);
        if (threads.length > 0) {
          threads[0].addLabel(label);
          Logger.log('Admin email labeled successfully');
        } else {
          Logger.log('Error: Could not find the sent email');
        }
        
    // Send client email with PDF attachment and social media links
    var clientSubject = 'Your Sacred Needles Consent Form';
    var instagramLink = `https://www.instagram.com/${config.InstagramHandle.replace('@', '')}/`;
    var facebookLink = `https://www.facebook.com/${config.FacebookPage.replace(/\s+/g, '')}/`;
    var clientBody = `
      <p>${firstName},</p>
      <p>Thank you for choosing Sacred Needles. Attached is your completed consent form for your records.</p>
      <p>We’d love to stay connected with you! Please follow us on social media to see our latest work, updates, and promotions:</p>
      <ul>
        <li>
          <a href="${instagramLink}" target="_blank">
            <img src="https://img.icons8.com/color/24/000000/instagram-new.png" 
                alt="Instagram" 
                style="width: 24px; height: 24px; vertical-align: middle; margin-right: 5px;">
            Follow us on Instagram: ${config.InstagramHandle}
          </a>
        </li>
        <li>
          <a href="${facebookLink}" target="_blank">
            <img src="https://img.icons8.com/color/24/000000/facebook-new.png" 
                alt="Facebook" 
                style="width: 24px; height: 24px; vertical-align: middle; margin-right: 5px;">
            Like our Facebook page: ${config.FacebookPage}
          </a>
        </li>
      </ul>
      <p>We look forward to seeing you soon!</p>
      <p>Best regards,<br>The Sacred Needles Team</p>
    `;
    MailApp.sendEmail({
      to: email || 'default@example.com',
      subject: clientSubject,
      htmlBody: clientBody,
      attachments: [pdfBlob]
    });


        // Log remaining quota after emails
        var quotaAfter = MailApp.getRemainingDailyQuota();
        Logger.log('Email quota remaining after submission: ' + quotaAfter);

        // Return success with firstName for the Thank You page
          return { status: 'success', firstName: firstName };
      } catch (error) {
        Logger.log('Submission Error: ' + error.message);
        throw new Error('Submission Error: ' + error.message);
      }
}
// Helper function to generate a unique filename
function getUniqueFileName(folder, baseName, extension) {
  var fileName = baseName + extension;
  var files = folder.getFilesByName(fileName);
  var counter = 1;

  while (files.hasNext()) {
    counter++;
    fileName = `${baseName}-${getOrdinalSuffix(counter)}${extension}`;
    files = folder.getFilesByName(fileName); // Check again with the new name
  }

  return fileName;
}

// Helper function to get ordinal suffix (e.g., 2nd, 3rd, 4th)
function getOrdinalSuffix(number) {
  var suffixes = ['th', 'st', 'nd', 'rd'];
  var value = number % 100;
  var suffix = suffixes[(value - 20) % 10] || suffixes[value] || suffixes[0];
  return number + suffix;
}
// Convert data URL to blob for image insertion
function dataURLtoBlob(dataURL) {
  try {
    if (!dataURL || typeof dataURL !== 'string' || !dataURL.startsWith('data:')) {
      throw new Error('Invalid or missing data URL');
    }
    var parts = dataURL.split(';base64,');
    if (parts.length < 2) throw new Error('Invalid data URL format: missing base64 part');
    var contentType = parts[0].split(':')[1];
    var raw = Utilities.base64Decode(parts[1]);
    var blob = Utilities.newBlob(raw, contentType);
    Logger.log('Blob created with size: ' + blob.getBytes().length + ', type: ' + contentType);
    return blob;
  } catch (error) {
    Logger.log('Error in dataURLtoBlob: ' + error.message);
    throw new Error('Blob conversion failed: ' + error.message);
  }
}

// Scale image to fit within max width while maintaining aspect ratio
function scaleImage(image, maxWidth) {
  var width = image.getWidth();
  var height = image.getHeight();
  if (width > maxWidth) {
    var newWidth = maxWidth;
    var newHeight = height * (maxWidth / width);
    image.setWidth(newWidth).setHeight(newHeight);
    Logger.log('Image scaled to width: ' + newWidth + ', height: ' + newHeight);
  }
}
// Helper function to validate image content based on file signature
function isValidImage(blob, mimeType) {
  try {
    var bytes = blob.getBytes();
    if (bytes.length < 12) {
      Logger.log('File too small to validate image signature');
      return false;
    }

    // Convert first few bytes to hex for comparison
    var hexSignature = '';
    for (var i = 0; i < 12; i++) {
      var byte = bytes[i] & 0xff; // Convert to unsigned byte
      hexSignature += ('0' + byte.toString(16)).slice(-2).toUpperCase() + ' ';
    }
    hexSignature = hexSignature.trim();
    Logger.log('File signature (hex): ' + hexSignature);

    // Check signatures based on MIME type
    if (mimeType === 'image/jpeg') {
      // JPEG: FF D8
      return hexSignature.startsWith('FF D8');
    } else if (mimeType === 'image/png') {
      // PNG: 89 50 4E 47
      return hexSignature.startsWith('89 50 4E 47');
    } else if (mimeType === 'image/gif') {
      // GIF: 47 49 46
      return hexSignature.startsWith('47 49 46');
    } else if (mimeType === 'image/bmp') {
      // BMP: 42 4D
      return hexSignature.startsWith('42 4D');
    } else if (mimeType === 'image/heic' || mimeType === 'image/heif') {
      // HEIC/HEIF: Look for 'ftypheic' or 'ftypheif' at offset 4
      var ftyp = '';
      for (var i = 4; i < 12; i++) {
        ftyp += String.fromCharCode(bytes[i]);
      }
      Logger.log('HEIC/HEIF ftyp: ' + ftyp);
      return ftyp === 'ftypheic' || ftyp === 'ftypheif';
    }

    return false;
  } catch (error) {
    Logger.log('Error validating image signature: ' + error.message);
    return false;
  }
}