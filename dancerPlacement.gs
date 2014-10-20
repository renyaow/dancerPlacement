function Person(name, sex, minD, maxD, preferences) {
  this.name = name;
  this.sex = sex;
  this.minD = parseInt(minD);
  this.maxD = parseInt(maxD);
  this.preferences = preferences;
  this.basicReq = false;
  this.numberOfDance = 0;
  this.dances = [];
}

function Dance(name, male, female, choreographers) {
  this.name = name;
  this.male = parseInt(male);
  this.female = parseInt(female);
  this.choreographers = choreographers;
  this.mDancers = [];
  this.fDancers = [];
  
  this.addDancer = function(dancer) {
    dancer.numberOfDance ++;
    if (dancer.sex == 'M' && !(inArray(dancer,this.mDancers))) {
      this.mDancers.push(dancer);
      dancer.dances.push(this);
    }
    else if (dancer.sex== 'F' && !(inArray(dancer,this.fDancers))) {
      this.fDancers.push(dancer);
      dancer.dances.push(this);
    }
  }
  
  this.removeDancer = function(dancer) {
    dancer.numberOfDance --;
    if (dancer.sex == 'M') {
      this.mDancers.splice(this.mDancers.indexOf(dancer), 1);
      dancer.dances.splice(dances.indexOf(this), 1);
    } 
    else if (dancer.sex == 'F') {
      this.fDancers.splice(this.fDancers.indexOf(dancer), 1);
      dancer.dances.splice(dances.indexOf(this), 1);
    }
  }
  
  this.isOverSized = function(gender) {
    if (gender == 'M') {
      return this.mDancers.length > this.male;
    } else {
      return this.fDancers.length > this.female;
    }
  }
}

var dancers = [];
var dances = [];
var danceMap = {};
var numberOfMale = 0;
var numberOfFemale = 0;
var numberOfDance = 0;
var spreadsheet;

/**
  * Check if this dance is full for gender of this dancer
  */
function danceNotFull(dance, dancer) {
  var dance = danceMap[dance];
  if (dance == null) {return;}
  if (dance.mDancers.length < dance.male && dancer.sex == 'M') {
    return true;
  }
  if (dance.fDancers.length < dance.female && dancer.sex == 'F') {
    return true;
  }
  return false;
}

function dancesNotFull(dances) {
  for (var i = 0; i <= dances.length - 1; i++) {
    var dance = dances[i];
    if (dance.mDancers.length < dance.male || dance.fDancers.length < dance.female) {
      return true;
    }
  }
  return false;
}

function inArray(item, array) {
  return array.indexOf(item) > -1
}


/**
 */
function readDancers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dancer Selection");
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows(); 
  var values = rows.getValues();

  for (var i = 1; i <= numRows - 1; i++) {
    var row = values[i];
    var preferences = [];
    for (var j = 5; j <= row.length -1; j++) {
      preferences.push(row[j]);
    }
    var person = new Person(row[1],row[2],row[3],row[4],preferences);
    dancers.push(person);
    
    if (row[2] == 'M') {
      numberOfMale ++;
    }
    else if (row[2] == 'F') {
      numberOfFemale ++;
    }
  }
};


function readDance() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dance Info');
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows(); 
  var values = rows.getValues();
  numberOfDance = numRows - 1;
  
  for (var i = 1; i <= numRows - 1; i++) {
    var row = values[i];
    var choreographers = [];
     for (var j = 3; j <= row.length -1; j++) {
      choreographers.push(row[j]);
    }
    var dance = new Dance(row[0], row[1],row[2], choreographers);
    dances.push(dance);
    danceMap[dance.name] = dance;
  }
}  



//start assignment
function assignment() {
  
  //add each person into their first choice
  for (var i = 0; i <= dancers.length - 1; i++) {
    //add choreographers into the dance
    var person = dancers[i];
    for (var j = 0; j <= dances.length - 1; j++) {
      var dance = dances[j];
      if (inArray(person.name,dance.choreographers)) {
        dance.addDancer(person);
      }
    }
    
    //if first choice is not choreographing piece 
    if (!inArray(person, danceMap[person.preferences[0]].mDancers) && !inArray(person, danceMap[person.preferences[0]].fDancers)) {
      danceMap[person.preferences[0]].addDancer(person);
    }
    person.basicReq = true; 
  }
  
  //check if any dance is over limit
   for (var i = 0; i <= dances.length - 1; i++) {
     var dance = dances[i];
     while (dance.mDancers.length > dance.male) {
       var dancer = dance.mDancers.pop();
       dancer.basicReq = false;
       dancer.numberOfDance --;
     }
     while (dance.fDancers.length > dance.female) {
       var dancer = dance.fDancers.pop();
       dancer.basicReq = false;
       dancer.numberOfDance --;
     }
   }
  
   //do second choice to make sure everyone get first/second
  for (var i = 0; i <= dancers.length - 1; i++) {
    var dancer = dancers[i];
    if (dancer.basicReq == false && dancer.numberOfDance < dancer.maxD) {
      if (danceMap[dancer.preferences[1]].isOverSized(dancer.sex)) {
        
        //2nd choice is unavailable,check 1st choice again
        var first = danceMap[dancer.preferences[0]];
        if (dancer.sex == 'M') {
          var swapDancer = first.mDancers[0];
        } else {
          var swapDancer = first.fDancers[0];
        }
        j = 0;
        while (danceMap[swapDancer.preferences[1]].isOverSized(swapDancer.sex) || 
                inArray(swapDancer.name, first.choreographers) && 
                 i < (first.mDancers.length+first.fDancers.length)) {
          i++;
          if (dancer.sex == 'M') {
            swapDancer = first.mDancers[i];
          } else {
            swapDancer = first.fDancers[i];
          }
        }
        
        //swap swapDancer and dancer
        first.removeDancer(swapDancer);
        danceMap[swapDancer.preferences[1]].addDancer(swapDancer);
        first.addDancer(dancer);
      }  else {
        danceMap[dancer.preferences[1]].addDancer(dancer);
      }
      dancer.basicReq = true;
    }
  }
  
  //fill in 2nd choice dances
  for (var i = 0; i <= dancers.length - 1; i++) {
    var dancer = dancers[i]; 
    if (dancer.numberOfDance < dancer.maxD) {
      if (danceNotFull(dancer.preferences[1], dancer)) {
        danceMap[dancer.preferences[1]].addDancer(dancer);
      }
    }
  }
  
  //fill in the rest
  var choice = 2;
  while (dancesNotFull(dances) && choice < numberOfDance) {
    for (var i = 0; i <= dancers.length - 1; i++) {
      var dancer = dancers[i]; 
      if (dancer.numberOfDance < dancer.maxD && danceNotFull(dancer.preferences[choice], dancer)) {
        danceMap[dancer.preferences[choice]].addDancer(dancer);
      }
    } 
    choice ++;
  }

}


function write() {
  var year = new Date().getFullYear();
  spreadsheet = SpreadsheetApp.create('Dancer Selection Result ' + year);
  
  var sheet = spreadsheet.insertSheet('By Dancer', 0);
  for (var i = 0; i <= dancers.length - 1; i++) {
    var dancer = dancers[i]; 
    var danceList = [];
    danceList.push(dancer.name);
    danceList.push(dancer.numberOfDance);
    for (var j = 0; j <= dancer.dances.length - 1; j++) {
      danceList.push(dancer.dances[j].name);
    }
    sheet.appendRow(danceList);
  }
  
  sheet = spreadsheet.insertSheet('By Dance', 1);
  for (var i = 0; i <= dances.length - 1; i++) {
    var dance = dances[i];
    var dancerList = [];
    dancerList.push(dance.name);
    for (var j = 0; j <= dance.fDancers.length - 1; j++) {
     dancerList.push(dance.fDancers[j].name);
    }
    for (var j = 0; j <= dance.mDancers.length - 1; j++) {
     dancerList.push(dance.mDancers[j].name);
    }
    sheet.appendRow(dancerList);
  }
}


function main() {
  readDancers();
  readDance();
  assignment();
  write();
}


//Dialog Box
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Dialog')
      .addItem('Open', 'openDialog')
      .addToUi();
}

function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Instructions');
}


//Email 
//Need Fixing with handlers
function email() {
  //var address = 'panasian-current@googlegroups.com';
  var address = 'weirenyao@gmail.com';
  var message = spreadsheet.getUrl();
  MailApp.sendEmail(address, 'Dancer Selection Result', message);
}