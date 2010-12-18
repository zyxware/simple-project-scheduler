var TEMPLATE_URL = 'http://bit.ly/gWmypz';

// Hook onOpen called at the time of opening the spreadsheet
function onOpen() {
  var menuEntries = [];
  menuEntries.push({name: "Sort by Member", functionName: "sortByMember"});
  menuEntries.push({name: "Sort by Project", functionName: "sortByProject"});
  menuEntries.push({name: "Add Phase to Project", functionName: "addPhaseToProject"});
  menuEntries.push({name: "Add Member to Project", functionName: "addMemberToProject"});
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Schedule", menuEntries);
  var menuEntries = [];
  menuEntries.push({name: "Add New Project", functionName: "addProject"});
  menuEntries.push({name: "Add New Member", functionName: "addNewMember"});
  menuEntries.push({name: "Add Next Month", functionName: "addNextMonth"});
  menuEntries.push({name: "Reload Config", functionName: "reloadConfig"});
  menuEntries.push({name: "Renumber Rows", functionName: "renumberRows"});
  menuEntries.push({name: "About App", functionName: "aboutApp"});
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Config", menuEntries);
}

function test() {
  var sa = new SchedulerApplication();
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

function addNextMonth() {
  var sa = new SchedulerApplication();
  sa.addNextMonth(true);
}

// Menu item added to custom menu to renumber the first column
function renumberRows() {
  if (Browser.msgBox('Please ensure that you have sorted the rows by project before proceeding. ' +
      'Press OK to continue, CANCEL to stop', 
      Browser.Buttons.OK_CANCEL) == 'ok') {
    var sa = new SchedulerApplication();
    sa.renumberRows();
  }
}

// Menu item added to custom menu to reload the configuration
function aboutApp() {
  Browser.msgBox('Simple Project Scheduler', 'Application developed and maintained by Zyxware Technologies. You can get support and the latest version from http://www.github.com/zyxware/simple-project-scheduler', Browser.Buttons.OK);
}
// Hook onEdit called whenever a cell is edited
function onEdit(event) {

  if (s.getName() == 'readme')
    return;

  var sa = new SchedulerApplication();
  //sa.load();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var s = event.source.getActiveSheet();
  var r = event.source.getActiveRange();
  
  //Browser.msgBox(sa.errors + ' before edit in global');
  sa.onEdit(event);
  //Browser.msgBox(schedulerApp.errors + ' after edit in global');
  
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
  //sa.store();
}
 
function SchedulerApplication() {
  
  var that = this;
  var rangeNames = Array();
  var sheets = Array();
  var sheetInfo = Array();
  var config = Array();
  var employees = Array();
  var employeeIndex = Array();
  var holidays = Array();
  
  this.init = function () {
    // Helper globals
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.s = SpreadsheetApp.getActiveSheet();
    this.errors = false;
    this.loadConfig();
    this.initialized = true;
  }
  
  // Refresh the class variables  
  this.onEdit = function (e) {
    // load config data
    //this.loadData();
    //dba(rangeNames);
  };
  
  // Load all the data for the current sheet
  // If this has already been loaded to sheet state then load from that.
  this.loadData = function(force) {
    //Browser.msgBox('loading');  
    var code = this.ss.getSheetByName('config').getRange('A1').getComment();
    if (code != '') {
      //Browser.msgBox(code);  
      result = eval(code);
      //Browser.msgBox(loaded);  
    }
    if (code == '' || (typeof loaded == 'undefined') || force) {
      code = this.reloadData();
      //Browser.msgBox(code);
      if (!this.errors) {
        this.ss.getSheetByName('config').getRange('A1').setComment(code);
      }  
    }
  }  
  

  // Load all the data for the current sheet
  this.reloadData = function() {
    
    // Get markers from the first column
    //Browser.msgBox('Loading schedule index');
    var index = this.s.getRange("A1:A10000").getValues();
    //Browser.msgBox(index);
    var num = index.length;
    var j = 0;
    for (var i = 0; i < num; i++) {
      // Start of data section
      if (index[i][0] == '#') {
        // +1 to get the spreadsheet index and +1 to get index of next row
        sheetInfo['schedule_start'] = i+1+1;
        j++;
      }
      // End of data section
      if (index[i][0] == '##') {
        // +1 to get the spreadsheet index and -1 to get index of previous row
        sheetInfo['schedule_end'] = i-1+1;
        j++;
      }
      // Start of Leaves and Holidays section
      if (index[i][0] == 'LH') {
        // +1 to get the spreadsheet index
        sheetInfo['leave_start'] = i+1;
        j++;
      }
      // Start of combined schedule section
      if (index[i][0] == 'FS') {
        // +1 to get the spreadsheet index
        sheetInfo['project_start'] = i+1;
        j++;
      }
    }
    if (j != 4) {
      Browser.msgBox('Invalid index column(A)', 
                     'Please ensure that the first column contains #, ##, LH, and FS ' +
                     'as demarcators. Read documentation to see format. Once you take ' + 
                     'care of this, reload sheet info from the menu', 
                     Browser.Buttons.OK_CANCEL);
      this.errors = true;  
      return false;  
    }
    // Get number of employees from the leave section
    var names = this.s.getRange("D" + sheetInfo['leave_start'] + ":D10000").getValues();
    var num = index.length;
    for (var i = 0; i < num; i++) {
      if (names[i][0].replace(/^\s*/, "").replace(/\s*$/, "") == '') {
        sheetInfo['leave_end'] = sheetInfo['leave_start'] + i - 1;
        sheetInfo['project_end'] = sheetInfo['project_start'] + i - 1 ;
        sheetInfo['num_employees'] = i;
        break;
      }
      else {
        employees[i] = names[i][0];
        employeeIndex[names[i][0]] = i + 1;
      }
    }
    rangeNames['currentEmployee'] = 'D1:AJ1';
    rangeNames['curEmpAvail'] = 'D2:AJ2';
    rangeNames['employeeAvailability'] = 'D' + sheetInfo['leave_start'] + ':AJ' + sheetInfo['leave_end'];
    rangeNames['employeeNames'] = 'D' + sheetInfo['project_start'] + ':D' + sheetInfo['project_end'];
    rangeNames['fullSchedule'] = 'D' + sheetInfo['project_start'] + ':AJ' + sheetInfo['project_end'];
    rangeNames['scheduleData'] = 'A' + sheetInfo['schedule_start'] + ':AJ' + sheetInfo['schedule_end'];
    //dba(rangeNames);
 
    code = '';
    code += generateEval('sheetInfo', sheetInfo);  
    code += generateEval('rangeNames', rangeNames);  
    code += generateEval('employeeIndex', employeeIndex);  
    code += generateEval('employees', employees);
    code += 'var loaded=true;';  
    return code;  
     
    //dba(employees);
    //dba(employeeIndex);
  };
    
  this.loadConfig = function(force) {
    //Browser.msgBox('loading');  
    var cs = this.ss.getSheetByName('config');
    // If the config sheet does not exist in the application set error 
    if (!cs) {
      Browser.msgBox('Missing config sheet', 
                     'You do not seem to have a config sheet in your application. ' + 
                     'This application requires a well defined config sheet. ' + 
                     'Please copy the latest template from ' + TEMPLATE_URL + ' and copy the config sheet to this spreadsheet.', 
                     Browser.Buttons.OK_CANCEL);
      this.errors = true;  
      return false;  
    }
    //Browser.msgBox('loading');  
    var code = this.ss.getSheetByName('config').getRange('A1').getComment();
    if (code != '') {
      //Browser.msgBox(code);  
      eval(code);
      //Browser.msgBox(loaded);  
    }
    // Parse config only on explicit force load requests
    if (code == '' || force) {
      code = this.parseConfig(cs);
      //Browser.msgBox(code);
      if (!this.errors) {
        this.saveConfig();
      }  
    }
  } 

  // Load all the data for the current sheet
  // If this has already been loaded to sheet state then load from that.
  this.parseConfig = function(cs, force) {
    //Browser.msgBox('loading');  
    // Get markers from the first column
    //Browser.msgBox('Loading config index');
    var index = cs.getRange("A1:A1000").getValues();
    //Browser.msgBox(index);
    var num = index.length;
    var j = 0;
    for (var i = 0; i < num; i++) {
      // Start of members list
      if (index[i][0] == 'Members') {
        // +1 to get the spreadsheet index, +1 to get next row after title row
        config['members_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['members_list_end'] = getNextEmptyCell(cs.getRange('B' + config['members_list_start']))-1;
        if (config['members_list_end'] < 0) 
          j--;
        j++;
      }
      // Start of projects list
      if (index[i][0] == 'Projects') {
        // +1 to get the spreadsheet index, +1 to get next row after title row
        config['projects_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['projects_list_end'] = getNextEmptyCell(cs.getRange('B' + config['projects_list_start']))-1;
        if (config['projects_list_end'] < 0) 
          j--;
        j++;
      }
      // Start of phases list, +1 to get next row after title row
      if (index[i][0] == 'Phases') {
        config['phases_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['phases_list_end'] = getNextEmptyCell(cs.getRange('B' + config['phases_list_start']))-1;
        if (config['phases_list_end'] < 0) 
          j--;
        j++;
      }
      // Start of holidays list, +1 to get next row after title row
      if (index[i][0] == 'Holidays') {
        config['holidays_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['holidays_list_end'] = getNextEmptyCell(cs.getRange('B' + config['holidays_list_start']))-1;
        if (config['holidays_list_end'] < 0) 
          j--;
        j++;
      }
      // Start of config list, +1 to get next row after title row
      if (index[i][0] == 'Params') {
        config['param_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['param_list_end'] = getNextEmptyCell(cs.getRange('B' + config['param_list_start']))-1;
        if (config['param_list_end'] < 0) 
          j--;
        j++;
      }
    }
    if (j != 5) {
      Browser.msgBox('Invalid config sheet', 
                     'Your config sheet does not look like it has everything it is supposed to have. ' + 
                     'This application requires a well defined config sheet. ' + 
                     'Please copy the latest template from ' + TEMPLATE_URL + ' and copy the config sheet to this spreadsheet.', 
                     Browser.Buttons.OK_CANCEL);
      this.errors = true;  
      return false;  
    }
    //dba(config);
  }
    
  this.saveConfig = function () {
    code = '';
    code += generateEval('config', config);
    this.ss.getSheetByName('config').getRange('A1').setComment(code);
    //Browser.msgBox(code);  
    return code;  
  } 
    
  // Create a sheet for the next month to the schedule    
  this.addNextMonth = function (){
    return;
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
    
  // Generate the code that can be eval'd to regenerate the config 
  // already parsed.
  function generateEval(name, arr) {
    code = '';
    //code = name + '=Array();';
    for (var key in arr) {
      if (typeof(arr[key]) == 'number') {
        code += name + "['" + key + "'] = " + arr[key] + ";\n";
      }
      else if (typeof(arr[key]) == 'string') {
        code += name + "['" + key + "'] = '" + arr[key] + "';\n";
      } 
      else if (typeof(arr[key]) == 'object') {
        code += name + "['" + key + "'] = Array();\n";
        // Recursively generate code for the array
        code += generateEval(name + "['" + key + "']", arr[key]); 
      } 
      else {
        throw "Invalid data type("+ typeof(arr[key]) +") passed";
      }
    }
    return code;
  }
  
  // Given a cell range the function iterates vertically until the 
  // next empty cell is found. Returns the row index of the next empty cell
  // or 0 if no empty cells are found  
  function getNextEmptyCell(range) {
    var cell = range;
    while (cell = cell.offset(1, 0, 1, 1)) {
      var value = cell.getValue();
      //Browser.msgBox(value + ':' + cell.getRowIndex());
      if (value.toString().replace(/^\s*/, "").replace(/\s*$/, "") == '') {
        return cell.getRowIndex();
      }
    }
    return 0;
  }

  // Check for a given key in the array and return the value if the key exists or return null
  // If the value for the key is null it will return null itself  
  function getArrayVal(arr, index) {  
    if (typeof arr[index] == 'undefined')
      return null;
    else
      return arr[index];
  }
  this.init();  
}

/**
 * debug functions
 */
function dba(obj) {
  Browser.msgBox(obj2string(obj));
}

function obj2string(obj) {
  var string = '';    
  for (var key in obj) {
    if (typeof(obj[key]) == 'object') {
      string += key + ':<' + obj2string(obj[key]) + ">"; 
    }
    else
      string += key + ':' + obj[key] + "    "; 
  }
  return string;
} 
  
function showGui() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
// create UiApp object named app
  var app = UiApp.createApplication().setTitle('my title');
//  .... populate app with ui objecs ...
// ..and display the UiApp object from the current spreadsheet
  var panel = app.createSimplePanel();

  var submit = app.createButton("Add User");

  c.addClickHandler('d');
  b.add(c);
  app.add(b);
  
  doc.show(app);
}
function d() {
  
}
/*
On Load
  Set Menus
  Reload Config

Config Per Sheet
  Per Sheet
  Range for Schedule data
  Range for Availability data
  Range for All Schedule data

Config For Spreadsheet  
  List of Holidays
  List of Team Members
  List of Projects
  Misc Config

On Edit
  If in a schedule sheet
    If within Schedule Region
  
    If within Availability Region
  
    If number of rows have changed
      Reload info for Sheet
      Save info for Sheet
  Else
    If in config sheet
      Reparse config
      Save config

On Add Phase to Project
  Add Phase row (Will this trigger on edit?)
  Reload Info for sheet
  Save Info for sheet

On Add Member to Project
  Add Member row (Will this trigger on edit?)
  Reload Info for sheet
  Save Info for sheet
*/
