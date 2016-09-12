//Global
var SHEET_NAME_ENTRY = '受注申請';
var SHEET_NAME_CHANGED = '受注変更申請';
var INITIAL_CONTROL_NO = '90000370';

var book = SpreadsheetApp.getActiveSpreadsheet();
var entry_sheet = book.getSheetByName(SHEET_NAME_ENTRY);
var changed_sheet = book.getSheetByName(SHEET_NAME_CHANGED);

var trigger_ids = new Array();

var START_ROW = 3;

//Global

function setControlNo(e) {
  var current_row = START_ROW;
  var current_control_no = entry_sheet.getRange(COL_CONTROL_NO + current_row).getValue();
  //c3セルを取得？
  var max_control_no = current_control_no;

  while (current_control_no != '') {
    current_row++;
    current_control_no = entry_sheet.getRange(COL_CONTROL_NO + current_row).getValue();

    if (current_control_no != '') {
      max_control_no = current_control_no; 
    }
  }

  if (max_control_no == '') {
    max_control_no = INITIAL_CONTROL_NO;
  }
  

  var no_id_record = entry_sheet.getRange(COL_FIRST_ANS + current_row).getValue();

  while (no_id_record != '') {
    max_control_no = parseInt(max_control_no) + 1;
    entry_sheet.getRange(COL_CONTROL_NO + current_row).setValue(max_control_no);


    current_row++;
    no_id_record = entry_sheet.getRange(COL_FIRST_ANS + current_row).getValue();

    for(var i=1; i<=85; i++){
	    var currentValue = entry_sheet.getRange(current_row-1,i).getValue();
    	if (typeof currentValue === "string") {
      		currentValue = currentValue.replace(/\s+/g, "");
 	  		currentValue = currentValue.replace(/１/g, "1");
      		currentValue = currentValue.replace(/２/g, "2");
      		currentValue = currentValue.replace(/３/g, "3");
      		currentValue = currentValue.replace(/４/g, "4");
      		currentValue = currentValue.replace(/５/g, "5");
      		currentValue = currentValue.replace(/６/g, "6");
      		currentValue = currentValue.replace(/７/g, "7");
      		currentValue = currentValue.replace(/８/g, "8");
      		currentValue = currentValue.replace(/９/g, "9");
      		currentValue = currentValue.replace(/０/g, "0");
      		entry_sheet.getRange(current_row-1,i).setValue(currentValue);
    	}//if
  	}//for
  }//while
}


function initTrigger(targetSheet) {
  // トリガーは起動時に全消し
  
  var allTriggers = ScriptApp.getScriptTriggers();
  
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getEventType() != ScriptApp.EventType.ON_OPEN) {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }
  
  var onEditTrigger = ScriptApp.newTrigger("onEdit")
  .forSpreadsheet(targetSheet)
  .onEdit()
  .create();
  
  trigger_ids.push(onEditTrigger.getUniqueId());
  
  var onSendFormTrigger = ScriptApp.newTrigger("setControlNo")
  .forSpreadsheet(targetSheet)
  .onFormSubmit()
  .create();
  
  trigger_ids.push(onSendFormTrigger.getUniqueId());
}

function init() {
  initTrigger(SpreadsheetApp.getActiveSpreadsheet());
}

var startEvent = false;
function onEdit(e) {
    
  switch (e.range.getSheet().getName()) {
    case SHEET_NAME_CHANGED:
      if (e.range.getColumn() == 6 && e.value != '') {
        Utilities.sleep(500);

        var ui = SpreadsheetApp.getUi();
        
        if (startEvent) {
          ui.alert("現在処理を実行中です、しばらくお待ち下さい。");
          return;
        }

        startEvent = true;
        
        try {
          switch (ui.alert('確認', '承認してもよろしいですか？', ui.ButtonSet.YES_NO)) {
            case ui.Button.YES:
              updateEntry(e);
              var title = "変更申請が承認されました";
              var body = "下記管理番号の変更申請が承認されました。\n\ngoo.gl/eKPNW5\n\n"
              + "管理番号：" + e.range.getSheet().getRange("C" + e.range.getRow()).getValue();
                var to='medical_hkn_bs@leverages.jp,kangohkn-all@leverages.jp';
                 var options = {from: 'medical_es@leverages.jp'};
 
				MailApp.sendEmail(to, title, body, options);
              
              break;
            case ui.Button.NO:
              e.range.setValue('棄却(' + e.value + ')');
              updateEntrydeny(e);
              var title2 = "変更申請が却下されました";
              var body2 = "下記管理番号の変更申請が却下されました。\n\ngoo.gl/eKPNW5\n\n"
              + "管理番号：" + e.range.getSheet().getRange("C" + e.range.getRow()).getValue();
               var to2='medical_hkn_bs@leverages.jp,kangohkn-all@leverages.jp';
             var options2 = {from: 'medical_es@leverages.jp'};
             MailApp.sendEmail(to2, title2, body2, options2);
              break;
            default:
              ui.alert("処理を中止しました");
              break;
          }
        } catch (e) {
          Logger.log(e);
          ui.alert('エラーが発生しました。'); 
        }
      }
      break;
  }
  
  startEvent = false;
}

function updateEntry(e) {
  var changed_row = e.range.getRow();

  var changed_control_no = changed_sheet.getRange(COL_CONTROL_NO + changed_row).getValue();

  var target_entry_row = 3;
  var target_entry_control_no = entry_sheet.getRange(COL_CONTROL_NO + target_entry_row).getValue();

  while (target_entry_control_no != changed_control_no && target_entry_control_no != "") {
    target_entry_row++;

    target_entry_control_no = entry_sheet.getRange(COL_CONTROL_NO + target_entry_row).getValue();
  }

  if (target_entry_control_no != '') {
    var changed_max = changed_sheet.getMaxColumns();
    for (var i = 3; i <= changed_max; i++) {

      var changed_range = changed_sheet.getRange(changed_row, i);
      var changed_ans = changed_range.getValue();
      if (changed_ans == '') {
        continue;
      }

      var header_text = getHeaderNumber(changed_range);
      if (header_text == '') {
        continue;
      }

      var target_entry_col = searchColumnNumberByHeaderText(entry_sheet, header_text);
      if (target_entry_col === 0) {
        continue;
      }

      entry_sheet.getRange(target_entry_row, target_entry_col).setValue(changed_ans);

    }
  }      changed_sheet.getRange(changed_row, 1).setValue("1");
}

//却下時に１立て
function updateEntrydeny(e) {
  var changed_row = e.range.getRow();

  var changed_control_no = changed_sheet.getRange(COL_CONTROL_NO + changed_row).getValue();

  var target_entry_row = 3;
  var target_entry_control_no = entry_sheet.getRange(COL_CONTROL_NO + target_entry_row).getValue();


  if (target_entry_control_no != '') {
    var changed_max = changed_sheet.getMaxColumns();
    for (var i = 3; i <= changed_max; i++) {

      var changed_range = changed_sheet.getRange(changed_row, i);
      var changed_ans = changed_range.getValue();
      if (changed_ans == '') {
        continue;
      }

    }
  }      changed_sheet.getRange(changed_row, 1).setValue("1");
}


function searchColumnNumberByHeaderText(sheet, headerText) {
  var range;
  var max = sheet.getMaxColumns();
  for (var i = 1; i <= max; i++) {
    range = sheet.getRange(2, i);
    if (range.getValue().split(' ', 1)[0] === headerText) {
      return i;
    }
  }

  return 0;
}

function getHeaderNumber(range) {
  return range.getSheet().getRange(2, range.getColumn()).getValue().split(' ', 1)[0];
}

function clog(text) {
  Logger.log(text);
}

function WFEdit(e) {
  switch (e.range.getSheet().getName()) {
    case SHEET_NAME_ENTRY:
      if (e.range.getColumn() == 1 && e.value != '') {

        var ui = SpreadsheetApp.getUi();
        
        startEvent = true;
        
        try {
          switch (ui.alert('確認', 'WF申請完了しましたか？', ui.ButtonSet.YES_NO)) {
            case ui.Button.YES:
              var title = "WF申請が完了しました";
              var body = "下記管理番号のWF申請が完了しました。\n\ngoo.gl/eKPNW5\n\n"
              + "管理番号：" + e.range.getSheet().getRange("C" + e.range.getRow()).getValue();

              var to='medical_hkn_bs@leverages.jp,kangohkn-all@leverages.jp';            
             var options = {from: 'medical_es@leverages.jp'};
             MailApp.sendEmail(to, title, body, options);
              ui.alert("処理が完了しました");
              break;
            case ui.Button.NO:
//                var entry_sheet = book.getSheetByName(SHEET_NAME_ENTRY);
//                var changed_row = e.range.getRow();
             e.range.setValue('' );
//              entry_sheet.getRange(changed_row + 1).setValue("1");
              break;
            default:
              ui.alert("処理を中止しました");
              break;
          }
        } catch (e) {
          Logger.log(e);
          ui.alert('エラーが発生しました。'); 
        }
      }
      break;
  }
  
  startEvent = false;
}