let advisorData = {};
let user;

function onEdit(e) {
  if (e.oldValue == "false") {
    user = e.user;
    let a1 = e.range.getA1Notation();
    let sheet = SpreadsheetApp.getActiveSheet();
    let numArray = a1.split("");
    numArray.shift();
    let row = numArray.join("");
    let range = sheet.getRange(row + ":" + row);
    let data = range.getValues();
    getAdvisorData(data[0][2].toLowerCase());
    draftEmail(data);
  }  
}

function getFromEmail() {
  let aliases = GmailApp.getAliases();
  for (let i = 0; i < aliases.length; i++) {
    if (aliases[i].includes("incoll")){
      return aliases[i];
    }
  }
}

// function findPrompt(link) {
//   let sheet = SpreadsheetApp.getActiveSheet();
//   let data = sheet.getDataRange().getValues();
//   for (let i = 0; i < data.length; i++) {
//     if (data[i][8] != "" && link == data[i][11]) {
//       return data[i][8];
//     }
//   }
// }

function draftEmail(data) {
    let template = HtmlService.createTemplateFromFile('email');
    template.name = data[0][1];
    template.advisor = advisorData;
    template.prompt = data[0][8];
    let message = template.evaluate().getContent();
    GmailApp.sendEmail(advisorData.advisorEmail + "," + data[0][4], "Congrats on Completing Your Essay", message, { htmlBody: message, from: getFromEmail() });
}


function getAdvisorData(advisorName) {
  Logger.log(advisorName)
  switch(advisorName) {
    case "yoshi": 
      advisorData.fullName = "Yoshi Akutsu";
      advisorData.position = "Lead Advisor";
      advisorData.phoneExtension = "709";
      advisorData.advisorEmail = "yoshi@incollegeplanning.com";
      break;
    case "juleanna": 
      advisorData.fullName = "Juleanna Smith";
      advisorData.position = "Lead Advisor";
      advisorData.phoneExtension = "710";
      advisorData.advisorEmail = "juleanna@incollegeplanning.com";
      break;
    case "emma": 
      advisorData.fullName = "Emma Mote";
      advisorData.position = "Lead Advisor";
      advisorData.phoneExtension = "708";
      advisorData.advisorEmail = "emma@incollegeplanning.com";
      break;
    case "sara":
      advisorData.fullName = "Sara Kapaj";
      advisorData.position = "Lead Advisor";
      advisorData.phoneExtension = "711";
      advisorData.advisorEmail = "sara@incollegeplanning.com";
      break;
    case "hannah":
      advisorData.fullName = "Hannah Laubach";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "712";
      advisorData.advisorEmail = "hannah@incollegeplanning.com";
      break;
    case "sian":
      advisorData.fullName = "Siân Lewis";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "723";
      advisorData.advisorEmail = "sian@incollegeplanning.com";
      break;
    case "eric":
      advisorData.fullName = "Eric Martinez";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "718";
      advisorData.advisorEmail = "eric@incollegeplanning.com";
      break;
    case "alecea":
      advisorData.fullName = "Alecea Howell";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "721";
      advisorData.advisorEmail = "alecea@incollegeplanning.com";
      break;
    case "samantha":
      advisorData.fullName = "Sam Rubinoski";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "720";
      advisorData.advisorEmail = "samantha@incollegeplanning.com";
      break;
    case "reilly":
      advisorData.fullName = "Reilly Grealis";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "719";
      advisorData.advisorEmail = "reilly@incollegeplanning.com";
      break;
    case "alex":
      advisorData.fullName = "Alex Horn";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "711";
      advisorData.advisorEmail = "alex@incollegeplanning.com";
      break;
    case "drake":
      advisorData.fullName = "Drake Hankins";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "702";
      advisorData.advisorEmail = "drake@incollegeplanning.com";
      break;
    case "sarah":
      advisorData.fullName = "Sarah Cook";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "700";
      advisorData.advisorEmail = "sarah@incollegeplanning.com";
      break;
    case "lydia":
      advisorData.fullName = "Lydia Crannell";
      advisorData.position = "College Planning Advisor";
      advisorData.phoneExtension = "713";
      advisorData.advisorEmail = "lydia@incollegeplanning.com";
      break;
    default: 
      break;
  }
}
