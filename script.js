function ask(){
  var result = SpreadsheetApp.getUi().alert("Retrieve account ranks and levels?", "You're about to refresh the ranks and levels.", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
  return result;
}

function onButtonPress() {
  var confirmation = ask();

  if(confirmation === SpreadsheetApp.getUi().Button.CANCEL || confirmation === SpreadsheetApp.getUi().Button.CLOSE) return;

  getRankAndLevel();

}

function onEdit(e) {
  var row = e.range.getRow();
  var col = e.range.getColumn();

  if (row == 1) return;

  var rankCols = [4, 5]
  var usingCol = 10

  if (rankCols.some(rankCol => rankCol === col)) {
    var timeStamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MMM/dd");
    e.source.getActiveSheet().getRange(row, rankCols[rankCols.length-1]+2).setValue(timeStamp);
  }
  else if (col == usingCol) {
    var timeStamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "M/dd/YYYY HH:mm:ss");
    e.source.getActiveSheet().getRange(row, usingCol+1).setValue(timeStamp);
  }
}

function getRankAndLevel() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var nameColumn = 'A:A'
  var nameData = sheet.getRange(nameColumn).getValues().filter(String)
  nameData = nameData.slice(1, nameData.length).map(nameCell => nameCell[0])
  Logger.log(nameData)

  var hashtagNames = nameData.filter(name => name.includes('#'))
  Logger.log(hashtagNames)

  var rankRequests = []
  var accRequests = []
  var eloRequests = []
  var lastMatchRequests = []

  hashtagNames.forEach(ingameName => {
    var region = 'na'
    var name = ingameName.split('#')[0]
    var tag = ingameName.split('#')[1]
    var rankURL = 'https://api.kyroskoh.xyz/valorant/v1/mmr/'
      + region + '/'
      + encodeURIComponent(name) + '/'
      + encodeURIComponent(tag);

    let request = {
      'url': rankURL,
      'method': 'GET',
      'muteHttpExceptions': true
    }
    rankRequests.push(request);
    
    var accUrl = 'https://api.henrikdev.xyz/valorant/v1/account/'
      + encodeURIComponent(name) + '/'
      + encodeURIComponent(tag);

    request = {
      'url': accUrl,
      'method': 'GET',
      'muteHttpExceptions': true
    }
    accRequests.push(request)

    var eloUrl = 'https://api.kyroskoh.xyz/valorant/v1/mmr/na/'
      + encodeURIComponent(name) + '/'
      + encodeURIComponent(tag)
      + "?show=eloonly&display=0";

    request = {
      'url': eloUrl,
      'method': 'GET',
      'muteHttpExceptions': true
    }

    var lastMatchUrl = 'https://api.henrikdev.xyz/valorant/v3/matches/na/'
      + encodeURIComponent(name) + '/'
      + encodeURIComponent(tag);

    request = {
      'url': lastMatchUrl,
      'method': 'GET',
      'muteHttpExceptions': true
    }

    lastMatchRequests.push(request);

  })

  let ranks = UrlFetchApp.fetchAll(rankRequests);
  let accounts = UrlFetchApp.fetchAll(accRequests);
  let elos = UrlFetchApp.fetchAll(eloRequests);
  let lastMatches = UrlFetchApp.fetchAll(lastMatchRequests);

  var rankCol = 'D'
  var rankRRCol = 'E'
  var lastRankUpdateCol = 'G'
  var eloCol = 'F'
  var accLvlCol = 'I'
  var lastMatchDateCol = 'M'
  
  hashtagNames.forEach( (ingameName, i) => {
    var userRow = nameData.indexOf(ingameName) + 2

    accResponse = accounts[i]
    rankResponse = ranks[i]
    eloResponse = elos[i]
    lastMatchResponse = lastMatches[i]
    Logger.log(lastMatchResponse)

    if(rankResponse.getResponseCode() >= 400) return;

    Logger.log(rankResponse.getResponseCode())
    Logger.log(`${ingameName}: ${rankResponse} : ${accResponse}`)

    var timeStamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd/YY");
    sheet.getRange(`${lastRankUpdateCol}${userRow}`).setValue(timeStamp)

    // Successfully retrieved account level
    let accLvl = processAccountLvl(accResponse)
    sheet.getRange(`${accLvlCol}${userRow}`).setValue(accLvl)

    let rankData = processRank(rankResponse);
    sheet.getRange(`${rankCol}${userRow}`).setValue(rankData[0])
    sheet.getRange(`${rankRRCol}${userRow}`).setValue(rankData[1])

    Logger.log(eloResponse)
    let eloVal = processElo(eloResponse)
    sheet.getRange(`${eloCol}${userRow}`).setValue(eloVal)

    let lastMatchDate = processLastMatch(lastMatchResponse)
    sheet.getRange(`${lastMatchDateCol}${userRow}`).setValue(lastMatchDate)
  })
}

function processAccountLvl(response){
  // Successfully retrieved account level
  let accLvl = '-'
  if (response.getResponseCode() == 200) accLvl = JSON.parse(response.getContentText())['data']['account_level']
  return accLvl;
}

function processRank(response){
  let rankVal = 'Unranked'
  let rrVal = '-'
  let rankData = response.getContentText();
  // If no rank, then they are unrated
  if (response.getResponseCode() !== 200 || rankData.includes('null')) rankVal = 'Unranked';
  // If has ascendant, then set its rank
  else if (rankData.includes('Ascendant')) rankVal = 'Ascendant';
  else if (rankData.includes('Iron')) rankVal = 'Iron';
  // Otherwise, its another rank
  else {
    rankVal = rankData.split('-')[0].trim();
    rrVal = rankData.split('-')[1].trim().replace('.', '');
  }
  return [rankVal, rrVal]
}

function processElo(response){
  if (response == null || response.getResponseCode() != 200 || response.getContentText().includes('null')) eloVal = '-'
  else eloVal = response.getContentText().split(':')[1]

  return eloVal;
}

function processLastMatch(response){
  if (response.getResponseCode() !== 200) return '-';

  let matchData = JSON.parse(response.getContentText())['data']
  if (matchData.length == 0) return '-';

  latestMatchEpoch = matchData[0]['metadata']['game_start']

  return Utilities.formatDate(new Date(latestMatchEpoch*1000), "America/Los_Angeles", "M/dd/YYYY HH:mm:ss");

}
