const DATA_START_ROW=9
const HEADER_ROW = DATA_START_ROW -1
const LABEL_ROW = HEADER_ROW -1
const ELIGIBLE_DES_COL=1
const ELIGIBLE_SYM_COL=ELIGIBLE_DES_COL+1
const ADDITION_COL=ELIGIBLE_SYM_COL+1
const DELETION_DES_COL=ADDITION_COL+2
const DELETION_SYM_COL=DELETION_DES_COL+1

const HIGHTLIGHT = "#fff2cc"

function hightlightAddition_(sheet){
  for (let current_row = DATA_START_ROW; true; current_row++ ){
    const dataRange = sheet.getRange(current_row,ELIGIBLE_DES_COL,1,2)
    if (dataRange.isBlank()){
      break
    }
    const additionCell = sheet.getRange(current_row, ADDITION_COL)
    if(additionCell.getValue() == "Addition"){
      dataRange.setBackground(HIGHTLIGHT)
    }

  }
}

const SYMBOL_WIDTH=120
const DESCRIPTION_WIDTH=360
function changeColumnWidths_(sheet){
  sheet.hideColumn(sheet.getRange(1,ADDITION_COL));
  sheet.setColumnWidth(ELIGIBLE_DES_COL, DESCRIPTION_WIDTH);
  sheet.setColumnWidth(ELIGIBLE_SYM_COL, SYMBOL_WIDTH);
  sheet.setColumnWidth(DELETION_DES_COL, DESCRIPTION_WIDTH);
  sheet.setColumnWidth(DELETION_SYM_COL, SYMBOL_WIDTH);
}

function crossOutDeletion_(sheet){
  for (let current_row = DATA_START_ROW; true; current_row++ ){
    const dataRange = sheet.getRange(current_row,DELETION_DES_COL,1,2)
    if (dataRange.isBlank()){
      break
    }
    dataRange.setFontLine("line-through")
  }
}


function makeHeadersBold_(sheet){
  const label1 = sheet.getRange(LABEL_ROW,ELIGIBLE_DES_COL)
  const label2 = sheet.getRange(LABEL_ROW,DELETION_DES_COL)
  const header1 = sheet.getRange(HEADER_ROW, ELIGIBLE_DES_COL, 1,3)
  const header2 = sheet.getRange(HEADER_ROW, DELETION_DES_COL,1,2)
  const toBold = [label1, label2, header1, header2]
  const toBorder = [header1, header2]
  for (let b of toBold){
    b.setFontWeight("bold")
  }
  for (let b of toBorder){
    b.setBorder(true, true, true, true, true, true)
  }
}


function main(){
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
  for (const s of sheets){
    changeColumnWidths_(s)
    crossOutDeletion_(s)
    makeHeadersBold_(s)
    hightlightAddition_(s)
  }
}