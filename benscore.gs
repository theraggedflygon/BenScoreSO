function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const options = [
    {name: "Setup BenScore", functionName: "setUpResultsSheet"},
    {name: "Sort Tournament Results", functionName: "sortResultsSheet"},
    {name: "Setup Event Scoresheet", functionName: "setUpEventScoreSheet"},
    {name: "Rank Event Teams", functionName: "rankEventTeams"}
  ];
  ss.addMenu("Macros", options);
}

function setUpResultsSheet() {
  const teams = getTeams();
  const events = getEvents();
  console.log(teams, events);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let masterSheet = ss.getSheetByName("Scoring")

  if (masterSheet !== null) {
    return;
  }

  masterSheet = ss.insertSheet("Scoring");
  const initCols = masterSheet.getMaxColumns();
  if (initCols > events.length + 6) {
    masterSheet.deleteColumns(1, initCols - (events.length + 6));
  } else if (initCols < events.length + 6) {
    masterSheet.insertColumns(1, events.length + 6 - initCols);
  }
  const headerRange = masterSheet.getRange(1, 1, 1, events.length + 6);
  const headers = ["Rank", "#", "Team Name"].concat(events).concat(["", "Team Score", "Total Points (Includes Drop)"]);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  
  const eventRange = masterSheet.getRange(1, 4, 1, events.length + 3);
  eventRange.setTextRotation(90);
  masterSheet.setColumnWidths(4, events.length + 3, 32);
  masterSheet.setColumnWidths(1, 2, 32);
  masterSheet.setColumnWidth(3, 200);

  const teamVals = []
  for (let i = 0; i < teams.length; i++) {
    teamVals.push([i + 1, i + 1, teams[i]])
  }

  const teamRange = masterSheet.getRange(2, 1, teams.length, 3);
  teamRange.setValues(teamVals);

  const sumRange = masterSheet.getRange(2, masterSheet.getMaxColumns() - 1, teams.length, 1);
  const sumRange2 = masterSheet.getRange(2, masterSheet.getMaxColumns(), teams.length, 1);
  const sumFormulas = new Array(teams.length).fill([`=SUM(R[0]C[-${events.length + 1}]:R[0]C[-1]) - MAX(R[0]C[-${events.length + 1}]:R[0]C[-1])`]);
  const sumFormulas2 = new Array(teams.length).fill([`=SUM(R[0]C[-${events.length + 2}]:R[0]C[-2])`]);
  sumRange.setFormulasR1C1(sumFormulas);
  sumRange2.setFormulasR1C1(sumFormulas2);

  masterSheet.setFrozenRows(1);
  masterSheet.setFrozenColumns(3);

  const totalRange = masterSheet.getRange(1, 1, teams.length + 1, events.length + 6);
  totalRange.setFontSize(9)
  totalRange.setFontFamily("Calibri");

  for (let i = 0; i < events.length; i++) {
    setUpEventSetupSheet(events[i]);
  }
}

function sortResultsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Scoring");
  const teams = getTeams();
  const events = getEvents();
  const dataRange = masterSheet.getRange(2, 2, teams.length, events.length + 5);
  dataRange.sort([events.length + 5, events.length + 6])
}

function setUpEventSetupSheet(eventName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  eventSheet = ss.insertSheet(eventName);
  const setupVals = [["#", "Name", "Points", "Tiebreak"]];
  for (let i = 0; i < 500; i++) {
    setupVals.push([i + 1, "", 0, "None"])
  }
  const setupRange = eventSheet.getRange(1, 1, 501, 4);
  setupRange.setValues(setupVals);

  eventSheet.setColumnWidth(1, 40);
  eventSheet.setColumnWidth(2, 200);
  eventSheet.setColumnWidth(3, 40);

  const instRange = eventSheet.getRange(2, 5, 3, 1);
  const instructions = [["Enter points values for up to 500 questions (Do not leave any rows with zero points if they are supposed to be a question). If desired, you can replace question number     with a nickname, but this is not required. Do not skip any rows."],
    ["Enter tiebreaker priority beginning with '1' in the column. Not all questions need to be marked with tiebreaker priority."],
    ["When you are finished, run the 'Setup Event Scoresheet' macro"]]
  instRange.setValues(instructions);
}

function setUpEventScoreSheet() {
  const eventSheet = SpreadsheetApp.getActiveSheet();
  if (!getEvents().includes(eventSheet.getName())){
    return;
  }
  const pointsRange = eventSheet.getRange(2, 3, 500, 1).getValues();
  const pointValues = [];
  for (let i = 0; i < 500; i++) {
    if (pointsRange[i][0] === 0) {
      break;
    }
    pointValues.push(pointsRange[i][0]);
  }
  
  const nameRange = eventSheet.getRange(2, 2, 500, 1).getValues();
  const questionLabels = [];
  for (let i = 0; i < pointValues.length; i++) {
    if (nameRange[i][0] !== "") {
      questionLabels.push(nameRange[i][0]);
    } else {
      questionLabels.push(`Question ${i + 1}`);
    }
  }

  const tbRange = eventSheet.getRange(2, 4, 500, 1).getValues();
  const tiebreakers = [];
  for (let i = 0; i < pointValues.length; i++) {
    if (tbRange[i][0] !== "None") {
      tiebreakers.push([tbRange[i][0], i]);
    }
  }
  tiebreakers.sort((t1, t2) => (t1[0] > t2[0] ? 1 : -1));
  let tiebreakerData = new Array(pointValues.length + 5);
  tiebreakerData[1] = "Tiebreaks"
  for (let i = 0; i < tiebreakers.length; i++) {
    tiebreakerData[tiebreakers[i][1] + 2] = tiebreakers[i][0];
  }
  
  eventSheet.clear();

  const initCols = eventSheet.getMaxColumns();
  if (initCols > pointValues.length + 5) {
    eventSheet.deleteColumns(1, initCols - (pointValues.length + 5));
  } else if (initCols < pointValues.length + 5) {
    eventSheet.insertColumns(1, pointValues.length + 5 - initCols);
  }

  const teams = getTeams();

  const data = [["#", "Team Name"].concat(questionLabels).concat(["", "Score", "Rank"])];
  data.push(["", "Points"].concat(pointValues).concat("", "", ""));
  data.push(tiebreakerData);
  data.push(new Array(pointValues.length + 5).fill(""));
  for (let i = 0; i < teams.length; i++) {
    let rowData = new Array(pointValues.length + 5).fill(0);
    rowData[0] = i + 1;
    rowData[1] = teams[i];
    rowData[pointValues.length + 2] = "";
    rowData[pointValues.length + 3] = "";
    rowData[pointValues.length + 4] = "N/A";
    data.push(rowData);
  }

  const dataRange = eventSheet.getRange(1, 1, data.length, data[0].length);
  dataRange.setValues(data);
  dataRange.setFontSize(9)
  dataRange.setFontFamily("Calibri");

  const scoreRange = eventSheet.getRange(5, pointValues.length + 4, teams.length, 1);
  const scoreFormulas = new Array(teams.length).fill([`=SUM(R[0]C[-${pointValues.length + 1}]:R[0]C[-2])`]);
  scoreRange.setFormulasR1C1(scoreFormulas);

  const headerRange = eventSheet.getRange(1, 3, 1, pointValues.length + 3);
  headerRange.setTextRotation(90);
  headerRange.setFontWeight("bold");
  eventSheet.setColumnWidths(3, pointValues.length + 3, 32);
  eventSheet.setColumnWidth(1, 32);
  eventSheet.setColumnWidth(2, 200);
  eventSheet.setFrozenRows(3);
  eventSheet.setFrozenColumns(2);

}

function rankEventTeams() {
  const eventSheet = SpreadsheetApp.getActiveSheet();
  if (!getEvents().includes(eventSheet.getName())){
    return;
  }
  const teams = getTeams();
  const qNum = eventSheet.getMaxColumns() - 5;
  
  const sheetScores = eventSheet.getRange(5, qNum + 4, teams.length, 1).getValues();
  const qscores = eventSheet.getRange(5, 3, teams.length, qNum).getValues();

  const tbRow = eventSheet.getRange(3, 3, 1, qNum).getValues()[0];
  const maxTB = Math.max(...tbRow.filter((n) => isNumeric(n)));
  tbIdxs = [];
  for (let i = 1; i <= maxTB; i++) {
    tbIdxs.push(tbRow.indexOf(i));
  }
  
  let scores = new Array(sheetScores.length).fill(0);
  let tbScores = new Array(sheetScores.length).fill(0);
  for (let i = 0; i < scores.length; i++) {
    scores[i] = [i, sheetScores[i][0]];
    tbTeam = [];
    for (let idx of tbIdxs) {
      tbTeam.push(qscores[i][idx]);
    }
    tbScores[i] = tbTeam;
  }

  scores.sort(function(s1, s2) {
    if (s1[1] === s2[1]) {
      const tb1 = tbScores[s1[0]];
      const tb2 = tbScores[s2[0]];
      for (let i = 0; i < tbScores.length; i++) {
        if (tb1[i] === tb2[i]) {
          continue;
        } else if (tb1[i] < tb2[i]) {
          return 1;
        } else {
          return -1;
        }
      }
      return -1;
    } else {
      return s1[1] < s2[1] ? 1 : -1;
    }
  })

  let rank = 1;
  let ranks = new Array(scores.length).fill(0);
  for (let i = 0; i < scores.length; i++) {
    ranks[scores[i][0]] = [rank];
    if (i !== scores.length - 1) {
      if (scores[i][1] !== scores[i + 1][1]) {
        rank = i + 2;
      } else {
          const tb1 = tbScores[scores[i][0]];
          const tb2 = tbScores[scores[i + 1][0]];
          for (let i = 0; i < tbScores.length; i++) {
            if (tb1[i] === tb2[i]) {
              continue;
            } else {
              rank = i + 2;
              break;
            }
        }
      }
    }
  }
  
  const rankRange = eventSheet.getRange(5, qNum + 5, teams.length, 1);
  rankRange.setValues(ranks);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Scoring");
  const events = getEvents();
  const eventIdx = events.indexOf(eventSheet.getName());
  const scoringEventRange = masterSheet.getRange(2, 4 + eventIdx, teams.length, 1);
  scoringEventRange.setValues(ranks);
}

function getTeams() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  const teamData = sheet.getRange("B2:B502").getValues();
  const teams = [];
  for (const row of teamData) {
    if (row[0] === "") {
      break;
    } else {
      teams.push(row[0]);
    }
  }
  return teams;
}

function getEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Data");
  const eventData = sheet.getRange("D2:D52").getValues();
  const events = [];
  for (const row of eventData) {
    if (row[0] === "") {
      break;
    } else {
      events.push(row[0]);
    }
  }
  return events;
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}
