// *** DO NOT RUN IF YOU ARE NOT MED/AMED ***

// remember to hit save before running!
// don't run init more than once!!!

// put link to your MED Spreadsheet below here
// make sure its surrounded by quotes ("")
var medUrl = "https://docs.google.com/spreadsheets/d/1fpwqnMtuLNcUHekj58N5kDvCIvfl13e67nQX90cViec/edit#gid=101253311"

// put link to your Candidate Point Book below here
// make sure its surrounded by quotes ("")
var formUrl = "https://docs.google.com/forms/d/1GbHD06lk988YHrZjVxNqRSXa3YUxiQQ2vJQh9nfYkAY/edit"

// Plan!
// if want to, add MED as a collaborator to the form
// pls dont change the order of the first two tabs
// manual: input spreadsheet url, form url, add candidates, run add candidates function

// check to see which candidate names are new
// put stuff in the blue cells, nowhere else
// add candidate names to question in form
// integrate NLP into summary tab

function init() {
  // create form w/ questions
  var candForm = createForm()

  // create spreadsheet from form
  // create spreadsheet
  var medss = SpreadsheetApp.create("MED Spreadsheet")
  SpreadsheetApp.setActiveSpreadsheet(medss);

  // link form
  candForm.setDestination(FormApp.DestinationType.SPREADSHEET, medss.getId())

  // format summary tab
  formatSummary(medss)
}

function addCandidates() {
  var medss = SpreadsheetApp.openByUrl(medUrl);
  SpreadsheetApp.setActiveSpreadsheet(medss);
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  // pull candidate names and emails
  var allCandInfo = getCandInfo(sheets[0].getRange("A4:B"))
  var medEmail = sheets[0].getRange("B1").getCell(1,1). getValue();

  // figuring out which names are new
  var newCandInfo = filterNewCands(allCandInfo, sheets)

  // fill in points/dates columns in "Summary"
  writeSummary(allCandInfo, sheets[0])

  // add candidate names to the dropdown in form
  addNamesToForm(allCandInfo)

  // create tabs for each candidate on main spreadsheet
  writeCandSpace(newCandInfo)

  // create a spreadsheet and share it with each candidate
  createCandSheets(newCandInfo, medEmail)
}

function formatSummary(medss) {
  // move summary sheet to first
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  medss.setActiveSheet(sheets[1])
  medss.moveActiveSheet(1)

  // set text
  var sumSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var range = sumSheet.getRange("A1:D3")
  var vals = [["MED email:", "", "", ""],
              ["", "", "", ""],
              ["Candidate Name", "Candidate Gmail", "Current Points", "Dates"]]

  range.setValues(vals)

  // set background colors
  sumSheet.getRange("B1:B1").setBackground("#c9daf8")
  sumSheet.getRange("A4:B").setBackground("#c9daf8")

  sumSheet.setName("Summary")

  // Slightly fix form responses tab name 
  sumSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]
  sumSheet.setName("Form Responses")
}

function createForm() {
  // create form
  var candForm = FormApp.create('Candidate Point Book')

  candForm.setShowLinkToRespondAgain(true)

  // create questions in form
  // Candidate Name
  var candNameQ = candForm.addListItem()
  candNameQ.setTitle("Candidate Name")
  candNameQ.setRequired(true)

  // Sister Name
  var sisNameQ = candForm.addTextItem()
  sisNameQ.setTitle("Sister Name")
  sisNameQ.setRequired(true)

  // Points Given
  var pointsQ = candForm.addTextItem()
  pointsQ.setTitle("Points Given")
  pointsQ.setRequired(true)

  // Fun Message
  var messageQ = candForm.addTextItem()
  messageQ.setTitle("Fun Message")
  
  // Are these points for a sister date?
  var dateQ = candForm.addMultipleChoiceItem()
  dateQ.setTitle("Are these points for a sister date?")
  dateQ.setChoices([
        dateQ.createChoice('Yes'),
        dateQ.createChoice('No')
     ])
  dateQ.setRequired(true)

  // Do you have any comments, concerns, or feedback regarding this candidate?
  var feedbackQ = candForm.addParagraphTextItem()
  feedbackQ.setTitle("Do you have any comments, concerns, or feedback regarding this candidate?")

  // Sister Password
  var passQ = candForm.addTextItem()
  passQ.setTitle("Sister Password")
  passQ.setRequired(true)

  return candForm
}

function addNamesToForm(cands) {
  // accessing the first question in the form
  var form = FormApp.openByUrl(formUrl)
  var question = form.getItems()[0].asListItem()

  var names = []
  // adding names as options
  for (cand of cands) {
    names.push(question.createChoice(cand[0]))
  }

  question.setChoices(names)
}

function filterNewCands(allCands, sheets) {
  // find list of existing cands from sheet names
  var oldCands = []
  for (sheet of sheets) {
    oldCands.push(sheet.getName())
  }

  // if a name exists in both lists, remove it from new cands
  var newCands = []
  for (cand of allCands) {
    if (!(oldCands.includes(cand[0]))) {
      newCands.push(cand)
    }
  }

  return newCands
}

function writeSummary(cands, sheet) {
  range = sheet.getRange(4, 3, cands.length, 2)
  var vals = []
  
  for (i in cands){
    vals.push(["=\'" + cands[i][0] + "\'!C1", "=\'" + cands[i][0] + "\'!E1"])
  }
  range.setFormulas(vals)
}

function createCandSheets(cands, medEmail) {
  for (i in cands) {
    // create spreadsheet
    var ss = SpreadsheetApp.create(cands[i][0])
    SpreadsheetApp.setActiveSpreadsheet(ss)
    var candSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]

    // fill in values
    var range = candSheet.getRange("A1:C5")
    var vals = [["=\"Candidate Name\"", "=\"Total Points Collected\"", "=\"Sister Dates Attended\""], 
      ["=\"" + cands[i][0] +"\"", "=IMPORTRANGE(\"" + medUrl + "\", \"\'" + cands[i][0] + "\'!C1\")", "=IMPORTRANGE(\"" + medUrl + "\", \"\'" + cands[i][0] + "\'!E1\")"], 
      ["", "", ""],
      ["=\"Sister Name\"", "=\"Fun Message\"", ""], 
      ["=IMPORTRANGE(\"" + medUrl + "\", \"\'" + cands[i][0] + "\'!A2:A\")", "=IMPORTRANGE(\"" + medUrl + "\", \"\'" + cands[i][0] + "\'!C2:C\")", ""]]

    range.setFormulas(vals)

    // add formatting
    candSheet.getRange("A1:C1").setBackground("#c9daf8")
    candSheet.getRange("A4:B4").setBackground("#c9daf8")
    
    // share with candidate + med
    var ssId = ss.getId()
    DriveApp.getFileById(ssId).addViewer(cands[i][1])

    if (medEmail) {
      DriveApp.getFileById(ssId).addEditor(medEmail)
    }
  }
}

function writeCandSpace(cands) {
  for (i in cands) {
    // create new candidate tab
    sheetNum = SpreadsheetApp.getActiveSpreadsheet().getSheets().length
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(cands[i][0], sheetNum + 1)

    // fill in formulas
    var range = sheet.getRange("A1:E1")
    var vals = [["=\"" + cands[i][0]+"\"", "=\"Points\"", "=SUM(C[-1])", "=\"Dates\"", "=COUNTIF(C[-1],\"Yes\")"]]
    range.setFormulasR1C1(vals)

    var range = sheet.getRange("A2:E2")
    vals = [["=FILTER('Form Responses'!$C:$G, 'Form Responses'!$B:$B = A$1, 'Form Responses'!$H:$H = \"yeehaw\")", "", "", "", ""]]
    range.setFormulas(vals)

    // adding border
    range = sheet.getRange("A1:E1")
    range.setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID)
  }
}

function getCandInfo(range) {
  var i = 1
  var result = []
  while (range.getCell(i, 1).getValue()) {
    result.push([range.getCell(i, 1).getValue(), range.getCell(i, 2).getValue()])
    i++;
  }
  return result
}
