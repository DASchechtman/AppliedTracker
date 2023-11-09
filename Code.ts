const JOB_DECLINED = "No"
const JOB_ACCEPTED = "Yes"
const COMPANY_COL = 0
const JOB_TITLE_COL = 1
const STATUS_COL = 6

function GetCompanyCol(data: any[][]) {
    const FETCHED_DATA = new Array<any>()

    for (let i = 1; i < data.length; i++) {
        FETCHED_DATA.push(data[i][COMPANY_COL])
    }

    return FETCHED_DATA
}

function GetJobStatusCol(data: any[][]) {
    const FETCHED_DATA = new Array<any>()
    
    for (let i = 1; i < data.length; i++) {
        FETCHED_DATA.push(data[i][STATUS_COL])
    }

    return FETCHED_DATA
}

function DaysToMS(days: number) {
    return days * 86400000
}

function CheckDeclined() {
  const SHEET_DATA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1")!.getDataRange().getValues()

  const COMPANIES = GetCompanyCol(SHEET_DATA)
  const STATUS = GetJobStatusCol(SHEET_DATA)
  
  for(let i = 1; i < SHEET_DATA.length; i++) {
    const COMPANY_NAME = COMPANIES[i]
    const COMPANY_STATUS = STATUS[i]
    const JOB_TITLE = SHEET_DATA[i][JOB_TITLE_COL]
    const PROPS = PropertiesService.getDocumentProperties()
    const KEY = `${COMPANY_NAME}-${JOB_TITLE}`
    const CACHED_COMPANY = PROPS.getProperty(KEY)
    const MS = Date.now()

    if (CACHED_COMPANY === null) {
        PROPS.setProperty(KEY, MS.toString())
        continue
    }

    if (Date.now() - Number(CACHED_COMPANY) >= DaysToMS(30) && COMPANY_STATUS !== JOB_ACCEPTED && COMPANY_STATUS !== JOB_DECLINED) {
        SHEET_DATA[i][STATUS_COL] = JOB_DECLINED
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1")!.getDataRange().setValues(SHEET_DATA)
}
