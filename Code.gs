const parentFolderId = '1STzJ2r91zi05A3VLZVqIM5GJYX8OI7wp'; // Change as necessary.//change
const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1') //change
const data = ss.getDataRange().getDisplayValues()
// Make sure we have this at the top of Code.gs
var WHATSAPP_GROUP_ID = "120363404166006005@g.us";
var WHATSAPP_API_URL = 'https://7103.api.greenapi.com/waInstance7103843365/sendMessage/0a5de9e314b3412a847b6a7de3df105fcf4a1039db56429cb3';


function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
}

function include(file){
  return HtmlService.createHtmlOutputFromFile(file).getContent()
}


function getDropdownValues() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database Form"); // Change "Sheet1" to your sheet's name
  var columnValuesPM = sheet.getRange("A2:A").getValues(); // Change "A2:A" to the range of your specific column
  var filteredValues = columnValuesPM.filter(String); // Removes any empty strings from the array
  return filteredValues;
}


function getData() {
 return data.slice(1)
}

function readId(id) {
  let rowID = data.find(r => {
    return r[1] == id
  })
  return rowID
}

function generateRandomAlphanumeric() {
  result = Math.floor(1000 + Math.random() * 9000);
  return result;
}

function getTodaysDateFormatted() {
  var today = new Date();
  var dd = String(today.getDate()).padStart(2, '0');
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var yyyy = today.getFullYear();

  var formattedDate = dd + '/' + mm + '/' + yyyy;
  Logger.log(formattedDate);
  return formattedDate;
}

function getDropdownValues() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Database Form');
    
    // Get all values from respective columns, starting from row 2
    const roomValues = sheet.getRange('H2:H' + sheet.getLastRow()).getValues().flat();
    const snameValues = sheet.getRange('I2:I' + sheet.getLastRow()).getValues().flat();
    const classSTDValues = sheet.getRange('J2:J' + sheet.getLastRow()).getValues().flat();
    const birthdayValues = sheet.getRange('K2:K' + sheet.getLastRow()).getValues().flat();
    const fatherValues = sheet.getRange('L2:L' + sheet.getLastRow()).getValues().flat();
    const motherValues = sheet.getRange('M2:M' + sheet.getLastRow()).getValues().flat();
    const plantValues = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();

    // Remove duplicates and empty values
    return {
        roomValues: [...new Set(roomValues.filter(Boolean))],
        snameValues: [...new Set(snameValues.filter(Boolean))],
        classSTDValues: [...new Set(classSTDValues.filter(Boolean))],
        birthdayValues: [...new Set(birthdayValues.filter(Boolean))],
        fatherValues: [...new Set(fatherValues.filter(Boolean))],
        motherValues: [...new Set(motherValues.filter(Boolean))],
        plantValues: [...new Set(plantValues.filter(Boolean))]

    };
}

function getFormattedTimestamp() {
  var now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function getUsername() {
    try {
        var email = Session.getActiveUser().getEmail();
        return email ? email.split('@')[0] : 'Guest';
    } catch (e) {
        console.error('Error getting username:', e);
        return 'Guest';
    }
}

// Function to send email on new entry
function sendEmailOnNewEntry(obj, sID, myPic) {
    var recipientEmail = "bittumonkbpdy@gmail.com";
    var ccEmail = "b2kbpdy@gmail.com";
    var subject = "Aperto TCK Manutenzione No: " + sID + " " + obj.sID13 + " il " + getTodaysDateFormatted();
    
    var replyBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="font-family: Arial, sans-serif; margin: 0; padding: 20px; color: #333333;">
        <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
          <!-- Header -->
          <div style="background-color: #2c3e50; color: #ffffff; text-align: center; padding: 20px; margin-bottom: 20px; border-radius: 6px;">
            <h2 style="margin: 0; font-size: 24px;">Aperto Ticket Manutenzione No: ${sID} ${obj.sID13} il ${getTodaysDateFormatted()}</h2>
          </div>

          <!-- Content Table -->
          <table cellpadding="10" cellspacing="0" style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <th style="width: 40%; background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">TICKET</th>
              <td style="width: 60%; background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${sID}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Impianto</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.sID13}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Priorità</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.classSTD}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Responsabile Interno</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.room}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Programmato il</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.gender1}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Dispositivi</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.sname}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Problemi</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.fname}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Ultimo aggiornamento</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${getTodaysDateFormatted()}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Tecnico</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.birthday}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Stato Impianto</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.father}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Stato Ticket</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.mother}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Istoria</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.weight}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Aggiornamento</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.height}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Link drive</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">
                <a href="${myPic}" style="color: #3498db; text-decoration: underline;">Visualizza</a>
              </td>
            </tr>
          </table>

          <!-- Footer -->
          <div style="border-top: 2px solid #3498db; padding-top: 20px; margin-top: 20px; text-align: center; color: #666666; font-size: 12px;">
            <p style="margin: 0;">Email generata automaticamente da TCK Manutenzione</p>
            <p style="margin: 5px 0;">Creato da: ${getUsername()}</p>
            <p style="margin: 5px 0;">Data: ${getTodaysDateFormatted()}</p>
          </div>
        </div>
      </body>
      </html>
    `;

    GmailApp.sendEmail(recipientEmail, subject, "", {
        htmlBody: replyBody
    });
  
}



function formatWhatsAppMessage(obj, sID, myPic) {
    return `*TCK Manutenzione - Nuovo Ticket*\n\n` +
        `*Creato il:* \`\`\`${getFormattedTimestamp()}\`\`\`\n\n` +
        `*TCK Manutenzione No:* \`\`\`${sID} ${obj.sID13}\`\`\`\n\n` +
        `*TCK:* \`\`\`${sID}\`\`\`\n\n` +
        `*Impianto:* \`\`\`${obj.sID13}\`\`\`\n\n` +
        `*Priorità:* \`\`\`${obj.classSTD}\`\`\`\n\n` +
        `*Responsabile Interno:* \`\`\`${obj.room}\`\`\`\n\n` +
        `*Programmato il:* \`\`\`${obj.gender1}\`\`\`\n\n` +
        `*Dispositivi:* \`\`\`${obj.sname}\`\`\`\n\n` +
        `*Problemi:* \`\`\`${obj.fname}\`\`\`\n\n` +
        `*Tecnico:* \`\`\`${obj.birthday}\`\`\`\n\n` +
        `*Stato Impianto:* \`\`\`${obj.father}\`\`\`\n\n` +
        `*Stato Ticket:* \`\`\`${obj.mother}\`\`\`\n\n` +
        `*Link Drive:* \`\`\`${myPic}\`\`\`\n\n` +
        `_Creato da: ${getUsername()}_`;
}


// Function to send email on new entry
function sendEmailOnButton(obj, tempo, cap) {
    var recipientEmail = cap;
    var ccEmail = "bittumonkbpdy@gmail.com";
    var subject = "Estrato creato il " + tempo + " per TCK Manutenzione No: " + obj[1] + " Impianto " + obj[2] + " aperto il " + obj[0];
    
    var replyBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="font-family: Arial, sans-serif; margin: 0; padding: 20px; color: #333333;">
        <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
          <!-- Header -->
          <div style="background-color: #2c3e50; color: #ffffff; text-align: center; padding: 20px; margin-bottom: 20px; border-radius: 6px;">
            <h2 style="margin: 0; font-size: 24px;">Aperto Ticket Manutenzione No: ${obj[1]} ${obj[2]} il ${obj[0]}</h2>
          </div>

          <!-- Content Table -->
          <table cellpadding="10" cellspacing="0" style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <th style="width: 40%; background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">TICKET</th>
              <td style="width: 60%; background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[1]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Impianto</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[2]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Priorità</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[3]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Responsabile Interno</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[4]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Programmato il</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[5]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Dispositivi</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[6]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Problemi</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[7]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Ultimo aggiornamento</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[8]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Tecnico</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[9]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Stato Impianto</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[10]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Stato Ticket</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[11]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Istoria</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[12]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Aggiornamento</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj[13]}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Link drive</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">
                <a href="${obj[14]}" style="color: #3498db; text-decoration: underline;">Visualizza</a>
              </td>
            </tr>
          </table>

          <!-- Footer -->
          <div style="border-top: 2px solid #3498db; padding-top: 20px; margin-top: 20px; text-align: center; color: #666666; font-size: 12px;">
            <p style="margin: 0;">Email generata automaticamente da TCK Manutenzione</p>
            <p style="margin: 5px 0;">Data: ${tempo}</p>
          </div>
        </div>
      </body>
      </html>
    `;

    GmailApp.sendEmail(recipientEmail, subject, "", {
        htmlBody: replyBody,
        cc: ccEmail
    });
}

// Function to search for email by subject and reply within the same conversation
function searchForEmailAndReply(obj, sID, led, myPic) {
    // Search for emails with the specified ticket number in both inbox and sent items
    var searchQuery = `subject:"TCK Manutenzione No: ${sID} ${obj.sID13}"`;
    Logger.log("Search Query: " + searchQuery);
    
    var threads = GmailApp.search(searchQuery);
    Logger.log("Found threads: " + threads.length);

    // Set the header based on ticket status
    var head = obj.classSTD === "Chiuso"
        ? `Chiuso Ticket Manutenzione No: ${sID} ${obj.sID13} il ${getTodaysDateFormatted()}`
        : `Aggiornamento Ticket Manutenzione No: ${sID} ${obj.sID13} il ${getTodaysDateFormatted()}`;

    // Construct the email body with inline styles
    var replyBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="font-family: Arial, sans-serif; margin: 0; padding: 20px; color: #333333;">
        <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff;">
          <!-- Header -->
          <div style="background-color: #2c3e50; color: #ffffff; text-align: center; padding: 20px; margin-bottom: 20px; border-radius: 6px;">
            <h2 style="margin: 0; font-size: 24px;">${head}</h2>
          </div>

          <!-- Content Table -->
          <table cellpadding="10" cellspacing="0" style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
            <tr>
              <th style="width: 40%; background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">TICKET</th>
              <td style="width: 60%; background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${sID}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Impianto</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.sID13}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Priorità</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.classSTD}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Responsabile Interno</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.room}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Programmato il</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.gender1}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Dispositivi</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.sname}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Problemi</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.fname}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Ultimo aggiornamento</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${getTodaysDateFormatted()}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Tecnico</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.birthday}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Stato Impianto</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.father}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Stato Ticket</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.mother}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Istoria</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.weight}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Aggiornamento</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">${obj.height}</td>
            </tr>
            <tr>
              <th style="background-color: #3498db; color: #ffffff; text-align: left; padding: 10px; border: 1px solid #cccccc;">Link drive</th>
              <td style="background-color: #f9f9f9; padding: 10px; border: 1px solid #cccccc;">
                <a href="${myPic}" style="color: #3498db; text-decoration: underline;">Visualizza</a>
              </td>
            </tr>
          </table>

          <!-- Footer -->
          <div style="border-top: 2px solid #3498db; padding-top: 20px; margin-top: 20px; text-align: center; color: #666666; font-size: 12px;">
            <p style="margin: 0;">Aggiornamento generato automaticamente da TCK Manutenzione</p>
            <p style="margin: 5px 0;">Aggiornato da: ${getUsername()}</p>
            <p style="margin: 5px 0;">Data: ${getTodaysDateFormatted()}</p>
          </div>
        </div>
      </body>
      </html>`;

    try {
        if (threads && threads.length > 0) {
            Logger.log("Found existing thread, replying...");
            var thread = threads[0];
            thread.replyAll("", {
                htmlBody: replyBody,
                cc: "bittumonkbpdy@gmail.com"
            });
        } else {
            Logger.log("No existing thread found, sending new email...");
            // If no thread is found, send a new email
            GmailApp.sendEmail(
                "bittumonkbpdy@gmail.com",
                `TCK Manutenzione No: ${sID} ${obj.sID13} - ${obj.classSTD === "Chiuso" ? "Chiuso" : "Aggiornamento"}`,
                "",
                {
                    htmlBody: replyBody
                }
            );
        }
    } catch (error) {
        Logger.log("Error in searchForEmailAndReply: " + error.toString());
        // Send as new email if there's an error
        GmailApp.sendEmail(
            "bittumonkbpdy@gmail.com",
            `TCK Manutenzione No: ${sID} ${obj.sID13} - ${obj.classSTD === "Chiuso" ? "Chiuso" : "Aggiornamento"}`,
            "",
            {
                htmlBody: replyBody
            }
        );
    }
}



function saveData(obj){
  Logger.log('Starting saveData with obj:', obj);
  let myPic
  let file
  let rowID = data.findIndex (r=> r[1]== obj.sID)+1
    
    try{
      if(rowID > 1){
        if(obj.myFile.length == 0){
                myPic = ss.getRange(rowID,15).getValue()
          }else{
                file = folder.createFile(obj.myFile).getId()
                chumma = myPic      
          }   
          ss.getRange(rowID,2).setValue(obj.sID); 
          ss.getRange(rowID,3).setValue(obj.sID13);  
          ss.getRange(rowID,4).setValue(obj.classSTD); 
          ss.getRange(rowID,5).setValue(obj.room); 
          ss.getRange(rowID,6).setValue(obj.gender1); 
          ss.getRange(rowID,7).setValue(obj.sname); 
          ss.getRange(rowID,8).setValue(obj.fname); 
          ss.getRange(rowID,9).setValue(getTodaysDateFormatted()); 
          ss.getRange(rowID,10).setValue(obj.birthday); 
          ss.getRange(rowID,11).setValue(obj.father); 
          ss.getRange(rowID,12).setValue(obj.mother); 

          // Get the current text of 'weight' and 'height' fields.
          let currentWeight = ss.getRange(rowID, 13).getValue();
          let currentHeight = ss.getRange(rowID, 14).getValue();

          // Prepare the additional text to be appended.
          let updateText = `
          <div style="border: 1px solid #ddd; border-radius: 5px; padding: 2px; margin-bottom: 5px; background-color: #f9f9f9;">
              <p style="font-weight: bold; color: #555;">Aggiornato il ${getFormattedTimestamp()}</p>
              <p style="font-style: italic; color: #777;">By --> ${getUsername()}</p>
              <p>* ${obj.height}</p>
              <p style="font-style: italic; color: #777;">Stato: ${obj.mother}</p>
          </div>`;

          let height = "dd"
          // Append the update text to the 'weight' field.
          let updatedWeight = currentWeight + updateText;

          // Set the new value of the 'weight' field.
          ss.getRange(rowID, 13).setValue(updatedWeight);
          ss.getRange(rowID,14).setValue(obj.height); 
          ss.getRange(rowID,15).setValue(myPic);      
          searchForEmailAndReply(obj,obj.sID,obj.numSTD,myPic)
        
      }else{
          // Assume we are adding a new entry here.
          Logger.log('Creating new entry');
            folder = DriveApp.getFolderById(parentFolderId);

            sID = generateRandomAlphanumeric();
            Logger.log('Generated new sID:', sID);

            // Create folders structure
            let classFolders = folder.searchFolders('title="' + obj.sID13 + '"');
            if (classFolders.hasNext()) {
                classFolder = classFolders.next();
            } else {
                classFolder = folder.createFolder(obj.sID13);
            }

            let studentFolderName = 'TCK -' + sID + '- Creato il ' + getFormattedTimestamp();
            studentFolder = classFolder.createFolder(studentFolderName);
            studentFolder.createFolder("Primo");
            studentFolder.createFolder("Dopo");
            myPic = studentFolder.getUrl();

            // Create initial weight content
            let initialWeight = `<div style="border: 1px solid #ddd; border-radius: 5px; padding: 2px; margin-bottom: 5px; background-color: #f9f9f9;">
                <p style="font-weight: bold; color: #555;">Creato il ${getFormattedTimestamp()}</p>
                <p style="font-style: italic; color: #777;">by --> ${getUsername()}</p>
                <p style="font-style: italic; color: #777;">Stato: ${obj.mother}</p>
            </div>`;

            // Append new entry
            ss.appendRow([
                getTodaysDateFormatted(),
                sID,
                "'" + obj.sID13,
                obj.classSTD,
                obj.room,
                obj.gender1,
                obj.sname,
                obj.fname,
                getTodaysDateFormatted(),
                obj.birthday,
                obj.father,
                obj.mother,
                initialWeight,
                obj.height,
                myPic
            ]);

            Logger.log('New entry created with ID:', sID);

            // Send email notification
            sendEmailOnNewEntry(obj, sID, myPic);

            // Send WhatsApp notification
            try {
                const whatsappMessage = `*TCK Manutenzione - Nuovo Ticket*\n\n` +
                    `*Creato il:* \`\`\`${getFormattedTimestamp()}\`\`\`\n\n` +
                    `*TCK Manutenzione No:* \`\`\`${sID} ${obj.sID13}\`\`\`\n\n` +
                    `*TCK:* \`\`\`${sID}\`\`\`\n\n` +
                    `*Impianto:* \`\`\`${obj.sID13}\`\`\`\n\n` +
                    `*Priorità:* \`\`\`${obj.classSTD}\`\`\`\n\n` +
                    `*Responsabile Interno:* \`\`\`${obj.room}\`\`\`\n\n` +
                    `*Programmato il:* \`\`\`${obj.gender1}\`\`\`\n\n` +
                    `*Dispositivi:* \`\`\`${obj.sname}\`\`\`\n\n` +
                    `*Problemi:* \`\`\`${obj.fname}\`\`\`\n\n` +
                    `*Tecnico:* \`\`\`${obj.birthday}\`\`\`\n\n` +
                    `*Stato Impianto:* \`\`\`${obj.father}\`\`\`\n\n` +
                    `*Stato Ticket:* \`\`\`${obj.mother}\`\`\`\n\n` +
                    `*Link Drive:* \`\`\`${myPic}\`\`\`\n\n` +
                    `_Creato da: ${getUsername()}_`;

                const payload = {
                    chatId: "120363404166006005@g.us",
                    message: whatsappMessage
                };

                const options = {
                    'method': 'post',
                    'headers': {
                        'User-Agent': 'GREEN-API_POSTMAN/1.0',
                        'Content-Type': 'application/json'
                    },
                    'payload': JSON.stringify(payload),
                    'muteHttpExceptions': true
                };

                const url = 'https://7103.api.greenapi.com/waInstance7103843365/sendMessage/0a5de9e314b3412a847b6a7de3df105fcf4a1039db56429cb3';
                
                Logger.log('Sending WhatsApp notification...');
                const response = UrlFetchApp.fetch(url, options);
                const responseCode = response.getResponseCode();
                const responseText = response.getContentText();

                Logger.log('WhatsApp API Response Code:', responseCode);
                Logger.log('WhatsApp API Response:', responseText);

            } catch (whatsappError) {
                Logger.log('WhatsApp notification failed:', whatsappError);
            }

            Logger.log('New entry process completed successfully');
        }
    } catch (error) {
        Logger.log('Error in saveData:', error);
        throw error;
    }
}


function deleteData(id) {
  Logger.log("id"+id)
  let rowID = data.findIndex(r => r[1] == id) + 1
  Logger.log("row"+rowID)
  if (rowID > 1) {
    ss.deleteRow(rowID);
  }
  return true;
}

function testSetup() {
    testWhatsAppWithEmailData();
}

function testWhatsAppWithEmailData() {
    try {
        const testObj = {
            sID: "TEST123",
            sID13: "Test Location",
            classSTD: "High",
            room: "Test Room",
            gender1: "2025-02-20",
            sname: "Test Device",
            fname: "Test Problem",
            birthday: "Test Tech",
            father: "Test Status",
            mother: "Open",
        };
        
        const testMyPic = "https://drive.google.com/test-url";
        
        const message = formatWhatsAppMessage(testObj, testObj.sID, testMyPic);
        Logger.log('Test message:', message);
        
        const result = sendWhatsAppNotification(message);
        Logger.log('Test result:', result);
        
        return result;
    } catch (error) {
        Logger.log('Test failed:', error);
        throw error;
    }
}


function testWhatsAppSetup() {
    Logger.log('Testing WhatsApp setup...');
    
    // Test OAuth scopes
    if (!checkOAuthScopes()) {
        Logger.log('OAuth scopes not properly set up!');
        return;
    }

    // Test WhatsApp API with a simple message
    try {
        const testMessage = "*Test Message*\nSent at: " + new Date().toISOString();
        sendWhatsAppNotification(testMessage);
        Logger.log('Test message sent successfully!');
    } catch (error) {
        Logger.log('Test message failed:', error);
    }
}

function getNamesByCategory(category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database Form");
  const data = sheet.getDataRange().getValues();
  return data
    .filter(row => row[2] === category) // Filter rows by column C (index 2)
    .map(row => row[3]) // Map to names in column D (index 3)
    .filter(name => name); // Remove empty names
}

function getEmailByName(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database Form");
  const data = sheet.getDataRange().getValues();
  const row = data.find(row => row[3] === name); // Find row by name in column D (index 3)
  return row ? row[4] : ''; // Return email from column E (index 4)
}

function getAllContacts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database Form");
  const data = sheet.getRange("C2:E" + sheet.getLastRow()).getValues();
  
  return data
    .filter(row => row[0] && row[1] && row[2]) // Filter out empty rows
    .map(row => ({
      type: row[0],    // Column C - TIPO
      name: row[1],    // Column D - NOME
      email: row[2],   // Column E - Email
      phone: row[3]    // Column F - PHONE NUMBER
    }));
}




function getWhatsAppRecipients() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Database Form');
        
        // Get data from columns C, D, F
        const range = sheet.getRange('C2:F' + sheet.getLastRow());
        const values = range.getValues();
        
        return values
            .filter(row => row[0] && row[1] && row[3]) // Check if category, name, and number exist
            .map(row => {
                const isGroup = row[0] === 'Internal';
                let number = String(row[3]).trim();

                // Don't modify the group ID from the sheet
                if (!isGroup) {
                    // Only format personal numbers
                    number = number.replace(/[^0-9]/g, '');
                    if (!number.startsWith('39')) {
                        number = '39' + number;
                    }
                }

                return {
                    category: row[0],    // Column C (Client/Tecnico/Internal)
                    name: row[1],        // Column D (Name)
                    number: number,      // Raw number/group ID from sheet
                    type: isGroup ? 'group' : 'person'
                };
            });

    } catch (error) {
        Logger.log('Error in getWhatsAppRecipients:', error);
        return [];
    }
}

function sendWhatsAppMessage(requestData) {
    try {
        // Format the chat ID properly
        let chatId;
        if (requestData.messageType === 'group') {
            // For groups, use the ID directly from the sheet
            chatId = `${requestData.recipientNumber}@g.us`;
        } else {
            // For individual numbers
            let cleanNumber = requestData.recipientNumber.replace(/[^0-9]/g, '');
            if (!cleanNumber.startsWith('39')) {
                cleanNumber = '39' + cleanNumber;
            }
            chatId = `${cleanNumber}@c.us`;
        }

        // Debug logging
        Logger.log('Request Data:', requestData);
        Logger.log('Formatted chatId:', chatId);

        const url = 'https://7103.api.greenapi.com/waInstance7103843365/sendMessage/0a5de9e314b3412a847b6a7de3df105fcf4a1039db56429cb3';
        const payload = {
            chatId: chatId,
            message: requestData.message
        };

        Logger.log('Sending payload:', payload);

        const options = {
            'method': 'post',
            'headers': {
                'User-Agent': 'GREEN-API_POSTMAN/1.0',
                'Content-Type': 'application/json'
            },
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        };
        
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        Logger.log('API Response Code:', responseCode);
        Logger.log('API Response Text:', responseText);
        
        if (responseCode === 200) {
            return { success: true };
        } else {
            return { 
                success: false, 
                error: `API Error: ${responseText}\nAttempted chatId: ${chatId}`
            };
        }
    } catch (error) {
        Logger.log('Error in sendWhatsAppMessage:', error.toString());
        return { 
            success: false, 
            error: error.toString() 
        };
    }
}


function formatNewEntryWhatsAppMessage(obj, sID, myPic) {
    try {
        Logger.log('Formatting WhatsApp message for new entry:');
        Logger.log('sID:', sID);
        Logger.log('obj:', JSON.stringify(obj));
        Logger.log('myPic:', myPic);
        
        const message = `*TCK Manutenzione - Nuovo Ticket*\n\n` +
            `*Creato il:* \`\`\`${getFormattedTimestamp()}\`\`\`\n\n` +
            `*TCK Manutenzione No:* \`\`\`${sID} ${obj.sID13}\`\`\` , aperto il \`\`\`${getTodaysDateFormatted()}\`\`\`\n\n` +
            `*TCK:* \`\`\`${sID}\`\`\`\n\n` +
            `*Impianto:* \`\`\`${obj.sID13}\`\`\`\n\n` +
            `*Priorità:* \`\`\`${obj.classSTD}\`\`\`\n\n` +
            `*Responsabile Interno:* \`\`\`${obj.room}\`\`\`\n\n` +
            `*Programmato il:* \`\`\`${obj.gender1}\`\`\`\n\n` +
            `*Dispositivi:* \`\`\`${obj.sname}\`\`\`\n\n` +
            `*Problemi:* \`\`\`${obj.fname}\`\`\`\n\n` +
            `*Ultimo Aggiornamento:* \`\`\`${getTodaysDateFormatted()}\`\`\`\n\n` +
            `*Tecnico:* \`\`\`${obj.birthday}\`\`\`\n\n` +
            `*Stato Impianto:* \`\`\`${obj.father}\`\`\`\n\n` +
            `*Stato Ticket:* \`\`\`${obj.mother}\`\`\`\n\n` +
            `*Link Drive:* \`\`\`${myPic}\`\`\`\n\n` +
            `_Creato da: ${getUsername()}_`;

        Logger.log('Formatted message:', message);
        return message;
    } catch (error) {
        Logger.log('Error formatting WhatsApp message:', error);
        throw error;
    }
}

function sendWhatsAppNotification(message) {
    try {
        Logger.log('Sending WhatsApp notification...');
        
        const payload = {
            chatId: "393334451888@c.us",
            message: message
        };

        const options = {
            'method': 'post',
            'headers': {
                'User-Agent': 'GREEN-API_POSTMAN/1.0',
                'Content-Type': 'application/json'
            },
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        };

        const url = 'https://7103.api.greenapi.com/waInstance7103843365/sendMessage/0a5de9e314b3412a847b6a7de3df105fcf4a1039db56429cb3';
        
        Logger.log('Making WhatsApp API request...');
        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        Logger.log('WhatsApp API Response Code:', responseCode);
        Logger.log('WhatsApp API Response:', responseText);

        if (responseCode !== 200) {
            throw new Error(`WhatsApp API error: ${responseCode} - ${responseText}`);
        }

        return true;
    } catch (error) {
        Logger.log('Error in sendWhatsAppNotification:', error);
        throw error;
    }
}


function testWhatsAppConnection() {
    try {
        const testMessage = "*Test Message*\nThis is a test message sent at: " + new Date().toISOString();
        Logger.log('Sending test message:', testMessage);
        
        const result = sendWhatsAppNotification(testMessage);
        Logger.log('Test result:', JSON.stringify(result));
        
        return result;
    } catch (error) {
        Logger.log('Test failed:', error);
        Logger.log('Error stack:', error.stack);
        throw error;
    }
}

function runWhatsAppTest() {
    Logger.log('=== Starting WhatsApp Connection Test ===');
    try {
        const testResult = testWhatsAppConnection();
        Logger.log('Test completed:', JSON.stringify(testResult));
    } catch (error) {
        Logger.log('Test failed:', error);
    }
    Logger.log('=== Test Complete ===');
}

function checkOAuthScopes() {
    try {
        const response = UrlFetchApp.fetch('https://www.google.com');
        Logger.log('UrlFetchApp is properly authorized');
        return true;
    } catch (error) {
        Logger.log('UrlFetchApp authorization error:', error);
        return false;
    }
}


function sendWhatsAppForNewTicket(ticketData) {
    try {
        Logger.log('Formatting WhatsApp message with ticket data:', ticketData);

        const message = `*TCK Manutenzione - Nuovo Ticket*\n\n` +
            `*Creato il:* \`\`\`${ticketData.createdAt}\`\`\`\n\n` +
            `*TCK Manutenzione No:* \`\`\`${ticketData.sID} ${ticketData.sID13}\`\`\`\n\n` +
            `*TCK:* \`\`\`${ticketData.sID}\`\`\`\n\n` +
            `*Impianto:* \`\`\`${ticketData.sID13}\`\`\`\n\n` +
            `*Priorità:* \`\`\`${ticketData.classSTD}\`\`\`\n\n` +
            `*Responsabile Interno:* \`\`\`${ticketData.room}\`\`\`\n\n` +
            `*Programmato il:* \`\`\`${ticketData.gender1}\`\`\`\n\n` +
            `*Dispositivi:* \`\`\`${ticketData.sname}\`\`\`\n\n` +
            `*Problemi:* \`\`\`${ticketData.fname}\`\`\`\n\n` +
            `*Tecnico:* \`\`\`${ticketData.birthday}\`\`\`\n\n` +
            `*Stato Impianto:* \`\`\`${ticketData.father}\`\`\`\n\n` +
            `*Stato Ticket:* \`\`\`${ticketData.mother}\`\`\`\n\n` +
            `*Link Drive:* \`\`\`${ticketData.folderUrl}\`\`\`\n\n` +
            `_Creato da: ${ticketData.createdBy}_`;

        Logger.log('Message formatted:', message);

        const payload = {
            chatId: WHATSAPP_GROUP_ID,
            message: message
        };

        Logger.log('Sending payload:', payload);

        const options = {
            'method': 'post',
            'headers': {
                'User-Agent': 'GREEN-API_POSTMAN/1.0',
                'Content-Type': 'application/json'
            },
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
        };

        Logger.log('Making API request...');
        const response = UrlFetchApp.fetch(WHATSAPP_API_URL, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();

        Logger.log('API Response Code:', responseCode);
        Logger.log('API Response:', responseText);

        if (responseCode !== 200) {
            throw new Error(`WhatsApp API error: ${responseCode} - ${responseText}`);
        }

        return true;
    } catch (error) {
        Logger.log('Error in sendWhatsAppForNewTicket:', error);
        throw error;
    }
}
