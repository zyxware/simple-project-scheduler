var TEMPLATE_URL = 'http://bit.ly/gWmypz';
var PROJECT_URL = 'http://www.github.com/zyxware/simple-project-scheduler';

var DEBUG = false;
// debug level 1 least verbose, bigger values more verbosity
var DEBUG_LEVEL = 1;

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
  alert("Hello");
  var sa = new SchedulerApplication();
  sa.loadConfig(true);
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
  if (confirm('Please ensure that you have sorted the rows by project before proceeding. ' +
      'Do you want to continue?') == 'yes') {
    var sa = new SchedulerApplication();
    sa.renumberRows();
  }
}

// Menu item added to custom menu to reload the configuration
function aboutApp() {
  alert('Application developed and maintained by Zyxware Technologies. You can get support and the latest version from ' + PROJECT_URL);
}
// Hook onEdit called whenever a cell is edited
function onEdit(event) {
  return;
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
      var empName = trim(scheduleRange.getCell(curRow - scheduleRange.getRow() + 1, 4).getValue());
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
          alert('Invalid name. Please check the spelling of the name - ' +
                empName + '. If you think the spelling is correct, ' +
                'please check initials and dots :-)');
        }
      }
    }
  }
  //sa.store();
}

function SchedulerApplication() {

  var rangeNames = Array();
  var sheetInfo = Array();
  var config = Array();
  var employees = Array();
  var employeeIndex = Array();

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
    this.loadSheetData();
    //dba(rangeNames);
  };

  // Load all the data for the current sheet
  // If this has already been loaded to sheet state then load from that.
  this.loadSheetData = function(force) {
    var code;
    if (!force && this.isSheetDataParsed()) {
      debug('loading sheet data', 1);
      code = this.getSheetDataCode();
    }
    if (!this.isSheetDataParsed() || force || code == '') {
      code = this.parseConfig();
      if (!this.errors) {
        this.saveSheetDataCode(code);
      }
    }
    // Load the data
    if (code != '') {
      //Browser.msgBox(code);
      eval(code);
    }
    return !this.errors;
  }
 
  this.isSheetDataParsed = function () {
    if (typeof(config['sheets'][this.s.getName()]) != 'undefined') {
      return true;
    }
    return false;
  }
  
  this.getSheetDataCode = function () {
    return this.ss.getSheetByName('config').getRange(config['sheets'][this.s.getName()]).getComment();
  }
  
  this.saveSheetDataCode = function (code) {
    return this.ss.getSheetByName('config').getRange(config['sheets'][this.s.getName()]).setComment(code);
  }
  
  // Load all the data for the current sheet
  this.parseSheetData = function() {
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
      alert('Please ensure that the first column contains #, ##, LH, and FS ' +
            'as demarcators. Read documentation to see format. Once you take ' +
            'care of this, reload sheet info from the menu');
      this.errors = true;
      return false;
    }
    // Get number of employees from the leave section
    var names = this.s.getRange("D" + sheetInfo['leave_start'] + ":D10000").getValues();
    var num = index.length;
    for (var i = 0; i < num; i++) {
      if (trim(names[i][0]) == '') {
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
    var code = '';
    debug('loading config');
    var cs = this.ss.getSheetByName('config');
    // If the config sheet does not exist in the application set error
    if (!cs) {
      alert('You do not seem to have a config sheet in your application. ' +
            'This application requires a well defined config sheet. ' +
            'Please copy the latest template from ' + TEMPLATE_URL + ' and copy the config sheet to this spreadsheet.');
      this.errors = true;
      return false;
    }
    if (!force) {
      debug('loading from file', 1);
      code = this.getConfigCode();
    }
    else {
      debug('loadConfig called with force', 1);
    }
    // If forced or if somebody had deleted the code manually reload it
    if (force || code == '') {
      debug('parsing config sheet', 1);
      code = this.parseConfig(cs);
      debug('got code: ' + code);
      if (!this.errors) {
        this.saveConfigCode(code);
      }
    }
    if (code != '') {
      //Browser.msgBox(code);
      eval(code);
    }
    //Browser.msgBox(code);
    // Also load the information about the sheets
    this.loadSheetInfo(cs, force);
    return !this.errors;
  }
  
  this.getConfigCode = function () {
    return this.ss.getSheetByName('config').getRange('A1').getComment();
  }

  this.saveConfigCode = function (code) {
    debug('saving config', 1);
    dba(config);
    if (code == '') {
      code = generateEval('config', config);
      //Browser.msgBox(code);
    }  
    debug(code, 2);
    this.ss.getSheetByName('config').getRange('A1').setComment(code);
    return code;
  }

  // Load config for the spreadsheet
  this.parseConfig = function(cs, force) {
    debug('start parsing config');
    // Get markers from the first column
    //Browser.msgBox('Loading config index');
    var index = cs.getRange("A1:A1000").getValues();
    //Browser.msgBox(index);
    var num = index.length;
    var i = 0, j = 0, k = 0, l = 0;
    var list = null;
    var cell;
    for (i = 0; i < num; i++) {
      // Start of members list
      if (index[i][0] == 'Members') {
        debug('parsing member list', 2);
        // +1 to get the spreadsheet index, +1 to get next row after title row
        config['members_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['members_list_end'] = getNextEmptyCell(cs.getRange('B' + config['members_list_start']))-1;
        if (config['members_list_end'] < 0)
          j--;
        else {
          config['members'] = Array();
          list = cs.getRange('B' + config['members_list_start'] + ":C" + config['members_list_end']).getValues();
          for (k = 0; k < list.length; k++) {
            if (list[k][1] == 'Y') {
              config['members'][k] = list[k][0];
            }
          }
        }
        j++;
      }
      // Start of projects list
      if (index[i][0] == 'Projects' && i > config['members_list_end']) {
        debug('parsing project list', 2);
        // +1 to get the spreadsheet index, +1 to get next row after title row
        config['projects_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['projects_list_end'] = getNextEmptyCell(cs.getRange('B' + config['projects_list_start']))-1;
        if (config['projects_list_end'] < 0)
          j--;
        else {
          config['projects'] = Array();
          config['projectInfo'] = Array();
          list = cs.getRange('B' + config['projects_list_start'] + ":E" + config['projects_list_end']).getValues();
          for (k = 0; k < list.length; k++) {
            if (list[k][3] == 'Y') {
              config['projects'][k] = list[k][0];
              config['projectInfo'][list[k][0]] = {'start':'', 'end':''};
              if (typeof(list[k][1]) == 'object') {
                config['projectInfo'][list[k][0]]['start'] = list[k][1].toUTCString();
              }
              if (typeof(list[k][2]) == 'object') {
                config['projectInfo'][list[k][0]]['end'] = list[k][2].toUTCString();
              }
            }
          }
        }
        j++;
      }
      // Start of phases list, +1 to get next row after title row
      if (index[i][0] == 'Phases' && i > config['projects_list_end']) {
        debug('parsing phases list', 2);
        config['phases_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['phases_list_end'] = getNextEmptyCell(cs.getRange('B' + config['phases_list_start']))-1;
        if (config['phases_list_end'] < 0)
          j--;
        else {
          config['phases'] = Array();
          list = cs.getRange('B' + config['phases_list_start'] + ":B" + config['phases_list_end']).getValues();
          for (k = 0; k < list.length; k++) {
            config['phases'][k] = list[k][0];
          }
        }
        j++;
      }
      // Start of holidays list, +1 to get next row after title row
      if (index[i][0] == 'Holidays' && i > config['phases_list_end']) {
        debug('parsing holidays list', 2);
        config['holidays_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['holidays_list_end'] = getNextEmptyCell(cs.getRange('B' + config['holidays_list_start']))-1;
        if (config['holidays_list_end'] < 0)
          j--;
        else {
          config['holidays'] = Array();
          config['holidaysInfo'] = Array();
          list = cs.getRange('B' + config['holidays_list_start'] + ":C" + config['holidays_list_end']).getValues();
          for (k = 0; k < list.length; k++) {
            if (typeof(list[k][0]) == 'object') {
              // Month is 0 based
              var month = list[k][0].getMonth();
              if (typeof(config['holidays'][month]) != 'object') {
                 config['holidays'][month] = Array();
              }
              // Date seems to be 0 based as well
              config['holidays'][month][list[k][0].getDate()] = list[k][1];
            }
          }
        }
        j++;
      }
      // Start of config list, +1 to get next row after title row
      if (index[i][0] == 'Params' && i > config['holidays_list_end']) {
        debug('parsing params list', 2);
        config['params_list_start'] = i+1+1;
        //-1 to get last non empty row
        config['params_list_end'] = getNextEmptyCell(cs.getRange('B' + config['params_list_start']))-1;
        if (config['params_list_end'] < 0)
          j--;
        else {
          config['params'] = Array();
          list = cs.getRange('B' + config['params_list_start'] + ":C" + config['params_list_end']).getValues();
          for (k = 0; k < list.length; k++) {
            cell = cs.getRange('C' + (config['params_list_start'] + k));
            switch (list[k][0]) {
              case 'hoursPerDay':
                if (typeof(list[k][1]) == 'number') {
                  config['params']['hoursPerDay'] = list[k][1];
                }
                else {
                  this.errors = true;
                  alert('Invalid config value: Invalid value for parameter - ' + list[k][0] + 
                       'Please read the instructions for filling in the parameters.');
                  return false;      
                }
                break;
              case 'offDays':
                var offDays = list[k][1].replace(/\s*/g, "").toLowerCase().split(",");
                for (l in offDays) {
                  if ("|sun|mon|tue|wed|thu|fri|sat|".search('\\|'+ offDays[l] +'\\|') < 0) {
                    this.errors = true;
                    alert('Invalid config value: Invalid value (' + offDays[l] + ') for parameter - ' + list[k][0] + '. ' +                        
                          'Please read the instructions for filling in the parameters.');
                    return false;      
                  }
                }
                config['params']['offDays'] = offDays;
                break;
              case 'offDayBg':
                config['params']['offDayBg'] = cell.getBackgroundColor();
                break;
              case 'leaveMarker':
                config['params']['leaveMarker'] = Array();
                config['params']['leaveMarker']['text'] = list[k][1];
                config['params']['leaveMarker']['textColor'] = cell.getFontColor();
                config['params']['leaveMarker']['bgColor'] = cell.getBackgroundColor();
                break;
              case 'releaseMarker':
                config['params']['releaseMarker'] = Array();
                config['params']['releaseMarker']['text'] = list[k][1];
                config['params']['releaseMarker']['textColor'] = cell.getFontColor();
                config['params']['releaseMarker']['bgColor'] = cell.getBackgroundColor();
                break;
              case 'curMemberBg':
                config['params']['curMemberBg'] = cell.getBackgroundColor();
                break;
              case 'skipSheets':
                var skipSheets = list[k][1].replace(/\s*,\s*/g, ",").toLowerCase().split(",");
                config['params']['skipSheets'] = '\\|'+ skipSheets.join('|') + '\\|';
                break;
              default:
                this.errors = true;
                alert('Unknown parameter - ' + list[k][0] + '. ' + 
                     'This application requires a well defined config sheet. ' +
                     'Please open ' + TEMPLATE_URL + ' and check the config sheet to see the list of configurable parameters.');
                return false;      
            }
          }
        }
        j++;
      }
    }
    if (j != 5) {
      alert('Your config sheet does not look like it has everything it is supposed to have or it is formatted incorrectly. ' +
            'This application requires a well defined config sheet. ' +
            'Please copy the latest template from ' + TEMPLATE_URL + ' and copy the config sheet to this spreadsheet.');
      this.errors = true;
      return false;
    }
    dba(config);
    debug('done parsing config');
    return generateEval('config', config);
  }
  
  // Function to load the sheetinfo data
  this.loadSheetInfo = function (cs, force) {
    var code;
    // Always load current sheetInfo code. Only update this value if forced
    code = this.getSheetInfoCode();
    if (force || code == '') {
      code = this.parseSheetInfo(cs);
      if (!this.errors) {
        this.saveSheetInfoCode(code);
      }
    }
    // Load the data
    if (code != '') {
      //Browser.msgBox(code);
      eval(code);
    }
    return !this.errors;
  }
  
  // Get the code for the sheetinfo data
  this.getSheetInfoCode = function () {
    return this.ss.getSheetByName('config').getRange('B1').getComment();
  }
  
  // Save the code for the sheetinfo data
  this.saveSheetInfoCode = function (code) {
    debug('saving sheetinfo', 1);
    dba(sheetInfo);
    if (code == '') {
      code = generateEval('sheetInfo', sheetInfo);
      //Browser.msgBox(code);
    }
    debug(code, 2);
    this.ss.getSheetByName('config').getRange('B1').setComment(code);
    return code;
  }
  
  // Parse sheets and build sheetinfo data
  this.parseSheetInfo = function (cs) {
    debug('start parsing sheetInfo');
    var sheets = this.ss.getSheets();
    var isValidSheet = Array();
    var stores = Array();
    var i, j, k, l;
    // Do not recreate the object if it already exists.
    if (typeof(sheetInfo) != 'object') {
      sheetInfo = Array();
    }
    // Invalidate current sheetInfo records
    for (j in sheetInfo) {
      isValidSheet[j] = false;
    }
    var skipSheets = config['params']['skipSheets'];
    debug('Will skip ' + skipSheets, 2);
    for (i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      debug('Processing ' + sheetName, 1);
      // Skip template sheets
      if (('|readme|template|config|').search('\\|'+ sheetName +'\\|') >= 0) {
        debug('Skipping ' + '\\|'+ sheetName.toString() +'\\|', 2);
        debug('skipSheets.search ' + ('|readme|template|config|').search('\\|'+ sheetName.toString() +'\\|'), 3);
        continue;
      }  
      // Skip sheets that are configured to be skipped  
      if (trim(skipSheets) != '' && skipSheets.search('\\|'+ sheetName.toLowerCase() +'\\|') >= 0) {
        debug('Skipping ' + sheetName, 2);
        debug('trim(skipSheets) ' + trim(skipSheets), 3);
        debug('skipSheets.search ' + skipSheets.search('\\|'+ sheetName +'\\|'), 3);
        continue;  
      }
      t = sheetName.toLowerCase().split(" ");
      var month = t[0]
      var year = parseInt(t[1]);
      // Check if the sheetnames are valid
      if (("|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|".search('\\|'+ month +'\\|') < 0)
            || (year < 2010) || (year > 2020)
      ) {
        this.errors = true;
        alert('Invalid sheet name: ' + sheetName + '. ' +                        
              'Sheets should be named in the format "Mmm YYYY". Please read the README provided at the project page - ' + PROJECT_URL);
        return false;      
      }
      // Create only if not initialized earlier
      if (typeof(sheetInfo[sheetName]) != 'object') {
        sheetInfo[sheetName] = Array();
      }
      sheetInfo[sheetName]['month'] = month;
      sheetInfo[sheetName]['year'] = year;
      isValidSheet[sheetName] = true;
    }
    var isAvailable;
    debug('Processing store numbers', 2);
    // Iterate through all sheets and assign stores or continue using same stores
    for (sheetName in sheetInfo) {
      if (isValidSheet[sheetName]) {
        // Allocate stores for sheets that don't already have one.
        if (typeof(sheetInfo[sheetName]['store']) != 'number') {
          // Assign stores from A3 onwards to the right in the config sheet
          for (k = 3; k < 1000; k++) {
            isAvailable = true;
            for (j in sheetInfo) {
              // Check among valid sheets if store is used
              if (sheetInfo[j]['store'] && isValidSheet[j] && sheetInfo[j]['store'] == k) {
                isAvailable = false;
                break;
              }
            }
            if (isAvailable) {
              sheetInfo[sheetName]['store'] = k;
              break;
            }
          }
          if (k == 1000) {
            this.errors = true;
            alert('Errors in assigning store. Please report this issue at - ' + PROJECT_URL);
            return false;      
          }
        }
      }
      // If the sheet is no longer valid delete the item
      else {
        delete sheetInfo[sheetName];
      }
    }
    dba(sheetInfo);
    debug('done parsing sheetInfo');
    return generateEval('sheetInfo', sheetInfo);
  }

  // Create a sheet for the next month to the schedule
  this.addNextMonth = function () {
    // Find last month

    // Find template
    // Copy template
    // Populate data in new template
      // Add active team members
      // Add active projects
      // Set month data
      // Set holidays
    // Increment last month
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
      alert('Schedule rows renumbed from 1 to ' + j);
    }
    else {
      alert('There are errors in this spreadsheet. ' +
        'Please fix those first before trying to use the functionalities');
    }
  }

  // Generate the code that can be eval'd to regenerate the config
  // already parsed.
  function generateEval(name, arr) {
    var code = '';
    debug('generating eval for ' + name, 1);
    //code = name + '=Array();';
    for (var key in arr) {
      debug(code, 5);
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
        throw "Invalid data type("+ typeof(arr[key]) +") passed for " + name + " with key " + key;
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
      if (trim(value.toString()) == '') {
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

// Utility function to trim a string
function trim(s) {
  return s.replace(/^\s*|\s*$/g, "");
}

/**
 * debug functions
 */
function dba(obj) {
  if (DEBUG)‌
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
  
function alert(prompt, buttons) {
  title = "Simple Scheduler";
  if (typeof(buttons) == 'undefined') 
    buttons = Browser.Buttons.OK_CANCEL;
  return Browser.msgBox(title, prompt, buttons);
}

function confirm(prompt, buttons) {
  if (typeof(buttons) == 'undefined') 
    buttons = Browser.Buttons.YES_NO_CANCEL;
  return alert(prompt, buttons);
}

// Debug function. Print message if debug level is <= current debug level
function debug(message, level) {
  if (typeof(level) == 'undefined')
    level = 0;
  if (DEBUG && (level <= DEBUG_LEVEL)) {
    alert(message, Browser.Buttons.OK);
  }
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
  debug('a', 5);
  debug('b', 6);
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
      Update currently edited member rows
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

​
