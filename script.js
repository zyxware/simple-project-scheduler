// Hook onOpen called at the time of opening the spreadsheet
function onOpen() {
  var menuEntries = [];
  menuEntries.push({name: "Sort by Member", functionName: "sortByMember"});
  menuEntries.push({name: "Sort by Project", functionName: "sortByProject"});
  //menuEntries.push({name: "Reload Config", functionName: "reloadConfig"});
  menuEntries.push({name: "Renumber Rows", functionName: "renumberRows"});
  menuEntries.push({name: "Check Names", functionName: "testNames"});
  menuEntries.push({name: "Resync Leave", functionName: "resyncLeave"});
  menuEntries.push({name: "Show Allocation", functionName: "showAlloc"});
  menuEntries.push({name: "About App", functionName: "aboutApp"});
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Schedule Utils", menuEntries);
}

// Menu item added to custom menu to sort the schedules by user
function sortByMember() {
  var sa = new SchedulerApplication();
  var r = sa.getNamedRange('scheduleData');
  r.sort([4, 1]); 
}

// Menu item added to custom menu to sort the schedules by project
function sortByProject() {
  var sa = new SchedulerApplication();
  var r = sa.getNamedRange('scheduleData');
  r.sort(1); 
}

// Menu item added to custom menu to reload the configuration
function reloadConfig() {
  var sa = new SchedulerApplication();
  sa.loadData(true);
}

// Menu item added to custom menu to renumber the first column
function renumberRows() {
  if (Browser.msgBox('Please ensure that the sheet is currently sorted by project before proceeding. ' +
      'Press OK to continue, CANCEL to stop', 
      Browser.Buttons.OK_CANCEL) == 'ok') {
    var sa = new SchedulerApplication();
    //dba("Created");
    sa.renumberRows();
  }
}

// Menu item added to test names
function testNames() {
  var sa = new SchedulerApplication();
  var scheduleRange = sa.getNamedRange('scheduleData');
  for (var curRow = scheduleRange.getRow(); curRow <= scheduleRange.getLastRow(); curRow++) {
    var empName = scheduleRange.getCell(curRow - scheduleRange.getRow() + 1, 4).getValue().replace(/^\s*/, "").replace(/\s*$/, "");
    var empIndex = Number(sa.getEmployeeIndex(empName));
    if (empIndex == 0 && empName != '') {
      Browser.msgBox('Invalid name', 
                     'Please check the spelling of the name - ' + 
                     empName + ' in row ' + curRow + '. If you think the spelling is correct, ' +
                     'please check initials and dots :-)', 
                    Browser.Buttons.OK_CANCEL);
    }  
  }
}

// Menu item added to resync leave for an individual
function resyncLeave() {
  var sa = new SchedulerApplication();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getActiveSheet();
  var r = s.getActiveRange();

  var availability = sa.getNamedRange('employeeAvailability');
  var scheduleRange = sa.getNamedRange('scheduleData');
  var calendarRange = sa.getNamedRange('calendar');
  var curRow = r.getRow();  

  // Check if the edit happened in an availability row
  if (!sa.errors && curRow >= availability.getRow() && curRow <= availability.getLastRow()) {
    curEmpCell = availability.getCell(curRow - availability.getRow() + 1, 1); 
    curEmpName = curEmpCell.getValue().replace(/^\s*/, "").replace(/\s*$/, "");
  }
  //dbc(curEmpCell);
  
  for (var curRow = scheduleRange.getRow(); curRow <= scheduleRange.getLastRow(); curRow++) {
    var empName = scheduleRange.getCell(curRow - scheduleRange.getRow() + 1, 4).getValue().replace(/^\s*/, "").replace(/\s*$/, "");
    if (empName == curEmpName) {
      var empIndex = Number(sa.getEmployeeIndex(empName));
      // Sync leave if employee name is found
      if (empIndex != 0 && empName != '') {
        for (var curDate = calendarRange.getColumn(); curDate <= calendarRange.getLastColumn(); curDate++) {
          // If the current date is a working day then sync leave from the leaves section to the current row
          calendarDateCell = calendarRange.getCell(1, curDate - calendarRange.getColumn() + 1);
          //dbc(calendarDateCell);
          //return false;
          if (calendarDateCell.getBackgroundColor() == '#ffffff') {
            curLeaveCell = availability.getCell(empIndex, availability.getColumn() + curDate - calendarRange.getColumn() - 1);
            curScheduleCell = scheduleRange.getCell(curRow - scheduleRange.getRow() + 1, curDate - calendarRange.getColumn() + 6);
            curLeave = (curLeaveCell.getValue() + " ").replace(/\s/, "");
            if (curLeave != '') {
              curLeaveCell.setBackgroundColor('#ff9900');
              curLeaveCell.setFontColor('#000000');
            }
            if (curLeave == '8') {
              //dbc(curLeaveCell);
              //dbc(curScheduleCell);
              //return false;
              //dbc(curScheduleCell);
              //return false;
              curScheduleCell.setValue('L');
              curScheduleCell.setBackgroundColor('#d9d9d9');
            }
            else {
              // If an L is marked incorrectly in the schedule cells then set it to blank
              if (curScheduleCell.getValue() == 'L') {
                curScheduleCell.setValue('');
                curScheduleCell.setBackgroundColor('#ffffff');
              }
            }
          }  
        }
      }
    }
  }
}

// Menu item added to custom menu to reload the configuration
function aboutApp() {
  Browser.msgBox('Simple Spreadsheet Scheduler', 'Application developed and maintained by Zyxware Technologies. You can get support and the latest version from http://www.github.com/zyxware/simple-spreadsheet-scheduler', Browser.Buttons.OK);
}

// Hook onEdit called whenever a cell is edited
function onEdit(event) {
  // commented until the eval error is fixed
  // http://code.google.com/p/google-apps-script-issues/issues/detail?id=897
  return;
  //sa.store();
  showAlloc(event);
}

function showAlloc() {
  var sa = new SchedulerApplication();
  //sa.load();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!(typeof loaded == 'undefined')) {
    var s = event.source.getActiveSheet();
    var r = event.source.getActiveRange();
    //Browser.msgBox(sa.errors + ' before edit in global');
    sa.onEdit(event);
    //Browser.msgBox(schedulerApp.errors + ' after edit in global');
  } 
  else {
    var s = ss.getActiveSheet();
    var r = s.getActiveRange();
  }
  
  var scheduleRange = sa.getNamedRange('scheduleData');
  var curRow = r.getRow();
  //Browser.msgBox('current row ' + curRow);
  //Browser.msgBox('current col ' + r.getColumn());
  //Browser.msgBox('scheduleRange.getRow() ' + scheduleRange.getRow());
  //Browser.msgBox('scheduleRange.getLastRow ' + scheduleRange.getLastRow());
  // Check if the edit happened in a schedule row
  if (!sa.errors && curRow >= scheduleRange.getRow() && curRow <= scheduleRange.getLastRow()) {
    var curCol = r.getColumn();
    if (curCol >= 4) {
      // trim and get the name
      var empName = scheduleRange.getCell(curRow - scheduleRange.getRow() + 1, 4).getValue().replace(/^\s*/, "").replace(/\s*$/, "");
      var empIndex = Number(sa.getEmployeeIndex(empName));
      //var empIndex = sa.getEmployeeIndex(empName);
      //var rc = sa.getNamedRange('currentEmployee');
      //Browser.msgBox(empName);
      //Browser.msgBox('Emp Index - ' + empIndex);
      if (curCol > 4 && empName != '') {
        var curEmp = sa.getNamedRange('currentEmployee');
        var fullSchedule = sa.getNamedRange('fullSchedule');
        //Browser.msgBox('Copying from ' + (fullSchedule.getRowIndex()+empIndex) + ' : ' +  fullSchedule.getColumnIndex());
        var empData = sa.s.getRange(fullSchedule.getRowIndex()+empIndex-1, fullSchedule.getColumnIndex(), 1, curEmp.getNumColumns());
        //Browser.msgBox(empData.getValues());
        empData.copyTo(curEmp, {contentsOnly:true});
        var curEmpAvail = sa.getNamedRange('curEmpAvail');
        var availability = sa.getNamedRange('employeeAvailability');
        var empAvl = sa.s.getRange(availability.getRowIndex()+empIndex-1, availability.getColumnIndex(), 1, curEmpAvail.getNumColumns());
        //Browser.msgBox(empAvl.getValues());
        empAvl.copyTo(curEmpAvail);
      }
      else {
        if (empIndex == 0 && empName != '') {
          Browser.msgBox('Invalid name', 
                         'Please check the spelling of the name - ' + 
                         empName + '. If you think the spelling is correct, ' +
                         'please check initials and dots :-)', 
                        Browser.Buttons.OK_CANCEL);
        }
      } 
    }
  }
}
 
function SchedulerApplication() {
  
  var that = this;
  var rangeNames = [];
  var config = [];
  var employees = [];
  var employeeIndex = [];
  
  // Helper globals
  this.ss = SpreadsheetApp.getActiveSpreadsheet();
  this.s = SpreadsheetApp.getActiveSheet();
 
  this.errors = false;
  this.initialized = false;
  
  // Refresh the class variables  
  this.onEdit = function (e) {
    // load config data
    //this.loadData();
    //dba(rangeNames);
  };
  
  // Load all the data for the current sheet
  // If this has already been loaded to sheet state then load from that.
  this.loadData = function(force) {
    // Browser.msgBox('loading');  
    var code = this.s.getRange('A1').getComment();
    // Eval throwing reference error hence setting code to '' to reload on every request
    // http://code.google.com/p/google-apps-script-issues/issues/detail?id=897
    code = '';
    if (!force && code != '') {
      result = eval(code);
    }
    if (force || (typeof loaded == 'undefined' && code == '')) {
      code = this.reloadData();
      // Browser.msgBox(code);
      if (!this.errors) {
        this.s.getRange('A1').setComment(code);
      }  
    }
  }  
  

  // Load all the data for the current sheet
  this.reloadData = function() {
    
    // Get markers from the first column
    // Browser.msgBox('Loading schedule index');
    var index = this.s.getRange("A1:A10000").getValues();
    // Browser.msgBox(index);
    var num = index.length;
    var j = 0;
    for (var i = 0; i < num; i++) {
      // Start of data section
      if (index[i][0] == '#') {
        config['schedule_start'] = i+1+1;
        j++;
      }
      // End of data section
      if (index[i][0] == '##') {
        config['schedule_end'] = i-1+1;
        j++;
      }
      // Start of Leaves and Holidays section
      if (index[i][0] == 'LH') {
        config['leave_start'] = i+1;
        j++;
      }
      // Start of combined schedule section
      if (index[i][0] == 'FS') {
        config['project_start'] = i+1;
        j++;
      }
    }
    if (j != 4) {
      Browser.msgBox('Invalid index column(A)', 
                     'Please ensure that the first column contains #, ##, LH, and FS ' +
                     'as demarcators. Read documentation to see format. Once you take ' + 
                     'care of this, reload config from the menu', 
                     Browser.Buttons.OK_CANCEL);
      this.errors = true;  
      return false;  
    }
    // Get number of employees from the leave section
    var names = this.s.getRange("D" + config['leave_start'] + ":D10000").getValues();
    var num = index.length;
    for (var i = 0; i < num; i++) {
      if (names[i][0].replace(/^\s*/, "").replace(/\s*$/, "") == '') {
        config['leave_end'] = config['leave_start'] + i - 1;
        config['project_end'] = config['project_start'] + i - 1 ;
        config['num_employees'] = i;
        break;
      }
      else {
        employees[i] = names[i][0];
        employeeIndex[names[i][0]] = i + 1;
      }
    }
    rangeNames['currentEmployee'] = 'D1:AJ1';
    rangeNames['curEmpAvail'] = 'D2:AJ2';
    rangeNames['calendar'] = 'F5:AJ6';
    rangeNames['employeeAvailability'] = 'D' + config['leave_start'] + ':AJ' + config['leave_end'];
    rangeNames['employeeNames'] = 'D' + config['project_start'] + ':D' + config['project_end'];
    rangeNames['fullSchedule'] = 'D' + config['project_start'] + ':AJ' + config['project_end'];
    rangeNames['scheduleData'] = 'A' + config['schedule_start'] + ':AJ' + config['schedule_end'];
    // dba(rangeNames);
 
    code = '';
    code += generateEval('config', config);  
    code += generateEval('rangeNames', rangeNames);  
    code += generateEval('employeeIndex', employeeIndex);  
    code += generateEval('employees', employees);
    code += 'var loaded=true;';  
     
    // dba(employees);
    // dba(employeeIndex);
    return code;  
  };
  // Generate the code that can be eval'd to regenerate the config 
  // already parsed.
  function generateEval(name, arr) {
    code = 'var ' + name + ' = [];';
    //code = name + '=Array();';
    for (var key in arr) {
      code += name + "['" + key + "']='" + arr[key] + "';";
    }
    return code
  }
    
  // Get the index of a given employee
  this.getEmployeeIndex = function (name) {
    return getArrayVal(employeeIndex, name);
  };

  // Get the named range for the current sheet
  this.getNamedRange = function (name) {
    //Browser.msgBox(rangeNames[name]);
    if (typeof rangeNames[name] == 'undefined')
      return null;
    else
      return this.s.getRange(rangeNames[name]);
  };
  
  this.renumberRows = function () {
    if (!this.errors) {
      var i, imin, imax, j;
      j = 0;
      imin = Number(config['schedule_start']);
      imax = Number(config['schedule_end']);
      for (var i = imin; i <= imax; i++) {
        j++;
        this.s.getRange('A' + i).setValue(j);
      }
      Browser.msgBox('Schedule rows renumbed from 1 to ' + j);
    } 
    else {
      Browser.msgBox('There are errors in this spreadsheet. ' +
        'Please fix those first before trying to use the functionalities');
    } 
  }
    
  // Check for a given key in the array and return the value if the key exists or return null
  // If the value for the key is null it will return null itself  
  function getArrayVal(arr, index) {  
    if (typeof arr[index] == 'undefined')
      return null;
    else
      return arr[index];
  }
  this.loadData();
}

/**
 * debug functions
 */
function dba(obj) {
  var string = '';    
  for (var key in obj) {
    string += key + ':' + obj[key] + "\n\n"; 
  }
  Browser.msgBox(string);
}

function dbc(cell) {
  Browser.msgBox(cell.getRow() + ":" + cell.getColumn() + " = " + cell.getValue() + " (" + cell.getBackgroundColor() + ")");
}

function test() {
  var a = "var config = [];config['schedule_start']='7';var loaded=true;";
  eval(a);
  Browser.msgBox('done');
}
