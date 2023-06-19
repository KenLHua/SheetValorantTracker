

function ask() {
    var result = SpreadsheetApp.getUi().alert("Retrieve account ranks and levels?", "You're about to refresh the ranks and levels.", SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);
    return result;
}

function onButtonPress() {
    var confirmation = ask();

    if (confirmation === SpreadsheetApp.getUi().Button.CANCEL || confirmation === SpreadsheetApp.getUi().Button.CLOSE) return;

    getRankAndLevel();

}

function onEdit(e) {
    var row = e.range.getRow();
    var col = e.range.getColumn();

    if (row == 1) return;

    var rankCols = [4, 5]
    var usingCol = 8

    if (rankCols.some(rankCol => rankCol === col)) {
        var timeStamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MMM/dd");
        e.source.getActiveSheet().getRange(row, rankCols[rankCols.length - 1] + 1).setValue(timeStamp);
    }
    else if (col == usingCol) {
        var timeStamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "M/dd/YYYY HH:mm:ss");
        e.source.getActiveSheet().getRange(row, usingCol + 1).setValue(timeStamp);
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
        accRequests.push(request);

    })

    Logger.log(rankRequests)

    let ranks = UrlFetchApp.fetchAll(rankRequests);
    let accounts = UrlFetchApp.fetchAll(accRequests);

    var accLvlCol = 'G'
    var lastRankUpdateCol = 'F'
    var rankCol = 'D'
    var rankRRCol = 'E'

    hashtagNames.forEach((ingameName, i) => {
        var userRow = nameData.indexOf(ingameName) + 2

        accResponse = accounts[i]
        rankResponse = ranks[i]

        Logger.log(`${ingameName}: ${rankResponse} : ${accResponse}`)


        var timeStamp = Utilities.formatDate(new Date(), "America/Los_Angeles", "MM/dd/YY");
        sheet.getRange(`${lastRankUpdateCol}${userRow}`).setValue(timeStamp)

        // Successfully retrieved account level
        let accLvl = '-'
        if (accResponse.getResponseCode() == 200) accLvl = JSON.parse(accResponse.getContentText())['data']['account_level']
        sheet.getRange(`${accLvlCol}${userRow}`).setValue(accLvl)

        let rankVal = 'Unranked'
        let rrVal = '-'
        let rankData = rankResponse.getContentText();
        // If no rank, then they are unrated
        if (rankResponse.getResponseCode() !== 200 || rankData.includes('null')) rankVal = 'Unranked';
        // If has ascendant, then set its rank
        else if (rankData.includes('Ascendant')) rankVal = 'Ascendant';
        // Otherwise, its another rank
        else {
            rankVal = rankData.split('-')[0].trim();
            rrVal = rankData.split('-')[1].trim().replace('.', '');
        }

        sheet.getRange(`${rankCol}${userRow}`).setValue(rankVal)
        sheet.getRange(`${rankRRCol}${userRow}`).setValue(rrVal)

    })

}