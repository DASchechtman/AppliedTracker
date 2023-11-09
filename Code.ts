const JOB_DECLINED = "No"
const JOB_ACCEPTED = "Yes"
const COMPANY_COL = 0
const JOB_TITLE_COL = 1
const STATUS_COL = 6


function GetCol(data: any[][], col: number) {
    const FETCHED_DATA = new Array<any>(data.length-1)

    for (let i = 1; i < data.length; i++) {
        FETCHED_DATA[i-1] = data[i][col]
    }
    
    return FETCHED_DATA
}

function GetData(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const FORMULAS = sheet.getDataRange().getFormulas()
    const DATA = sheet.getDataRange().getValues()

    for (let i = 0; i < FORMULAS.length; i++) {
        for(let j = 0; j < FORMULAS[i].length; j++) {
            if (FORMULAS[i][j] === "") {
                FORMULAS[i][j] = DATA[i][j]
            }
        }
    }

    return FORMULAS
}

function DaysToMS(days: number) {
    return days * 86400000
}

function CheckDeclined() {
  const SHEET_DATA = GetData(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1")!)

  const COMPANIES = GetCol(SHEET_DATA, COMPANY_COL)
  const STATUS = GetCol(SHEET_DATA, STATUS_COL)
  const JOB_TITLES = GetCol(SHEET_DATA, JOB_TITLE_COL)
  const PROPS = PropertiesService.getDocumentProperties()
  
  for(let i = 1; i < COMPANIES.length; i++) {
    const COMPANY_NAME = COMPANIES[i]
    const COMPANY_STATUS = STATUS[i]
    const JOB_TITLE = JOB_TITLES[i]
    const KEY = `${COMPANY_NAME}-${JOB_TITLE}`
    const CACHED_COMPANY = PROPS.getProperty(KEY)

    if (CACHED_COMPANY === null) {
        const MS = Date.now()
        PROPS.setProperty(KEY, MS.toString())
        continue
    }

    if (Date.now() - Number(CACHED_COMPANY) >= DaysToMS(30) && COMPANY_STATUS !== JOB_ACCEPTED && COMPANY_STATUS !== JOB_DECLINED) {
        SHEET_DATA[i][STATUS_COL] = JOB_DECLINED
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1")!.getDataRange().setValues(SHEET_DATA)
}

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
    const RANGE = e.range
    const LINK = RANGE.getValue()
    const NOTATION = RANGE.getA1Notation()
    if (LINK === "link" || LINK === "" || !NOTATION.includes('E') || NOTATION.includes(":")) { return }
    RANGE.setValue(`=hyperlink("${LINK}", "link")`)
}
