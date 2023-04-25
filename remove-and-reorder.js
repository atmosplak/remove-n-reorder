/*
Instructions for use: 

1) export CSV from Jira (All Fields)
2) duplicate sheet titled 'reorderMe[template]'
3) rename the duplicated sheet to 'reorderMe' 
4) click on cell 'B2' and import data with the following import location - 'replace data at selected cell'
5) run Macro called 'removeAndReorderColumns'
6) generate links in "link" column by clicking box in bottom right and dragging until you reach the last row
7) profit

Questions: reach out to csantiago
*/

async function removeAndReorderColumns() {
    var sheet = SpreadsheetApp.getActive().getSheetByName('reorderMe')
  
    // sequenced column deletions (not very DRY i know)
    sheet.deleteColumns(7, 10)
    sheet.deleteColumns(14, 106)
    sheet.deleteColumn(4)
    sheet.deleteColumn(4)
    sheet.deleteColumns(6, 3)
    sheet.deleteColumn(9)
    sheet.deleteColumn(10)
  
    // column shifts
    var columnSpec = sheet.getRange("I1:I1")
    sheet.moveColumns(columnSpec, 5)
    
    // add new columns and name them
    sheet.insertColumnAfter(3)
    sheet.insertColumnAfter(10)
  
    sheet.getRange('d2').setValue('Link')
    sheet.getRange('k2').setValue('Comments')
    sheet.getRange('k3').setValue('[Enter your comments here]')
  
    var cell = sheet.getRange('d3')
  
   cell.setFormula('="https://etsy.atlassian.net/browse/"&C3')
   await SpreadsheetApp.flush() // async function to allow preceding processes to finish before formatting cells
  
    sheet.autoResizeColumns(1, 11)
    sheet.getRange('A2').setValue('Action')
    
    var row = sheet.getRange('A2:K2')
    row.setFontWeight('bold')
  }
  