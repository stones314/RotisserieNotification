/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

/**
 * handlePick
 * Triggers when sheet is edited. Checks if cell for next card pick is edited.
 *
 * After you have copied the sheet and this script, you must set up this trigger yourself.
 * Se HowTo tab in sheet for details.
 */
function handlePick(){

  //************************************
  //* Things you can change (and probably should change)
  //************************************

  
  //Link to this sheet and to your cube at cubetutor. Will be included in mail-notifications.
  // If you dont have cubetutor link, just enter "Nothing to see here..." or something.
  var link = "link to this sheet, sent to playes when it is their turn to draft";
  var cubetutor = "https://www.cubetutor.com/viewcube/57513";
  
  //If you draft a cube set this to true. The cards in the cube must be added to the Cards tab for this to work.
  //If you want e.g. Modern Rotisserie, set it to false. In that case the script wont validate card names, only keep track
  //of whos turn it is to draft. That includes no check if a card was already drafed.
  var isCube = true;
  
  //Set number of players here. If this is set to e.g. 6, then only seat 1-6 will be used.
  var numPlayers = 8;

  //Add player emails here for anyone who wants notification when it is their turn to draft.
  //Has to be in the order of the draft (first email is the first player to pick a card, and so on)
  //Leave empty ("") if a player do not want notification.
  var emails = ['player1@somemail.com',
               'player2@somemail.com',
               'player3@somemail.com',
               'player4@somemail.com',
               'player5@somemail.com',
               'player6@somemail.com',
               'player7@somemail.com',
               'player8@somemail.com'];
  
  //For each player, set this to true if that player want notification of all picks made
  // (false means they only will be notified when it is their turn to pick a card)
  var notifyAll = [false,
                  false,
                  false,
                  false,
                  false,
                  false,
                  false,
                  false];
  
  //************************************
  //* Things you might want to change (but it is easier to leave it as is)
  //************************************


  //If you feel like it you can change the background color used for players here.
  //You must manually update the column headers to match
  var playerColor = [ "#7cb6eb" //Blue Dark
                     ,"#b6d7a8" //Green Light
                     ,"#ffd966" //Yellow
                     ,"#f7bcd5" //Pink
                     ,"#81d7eb" //Blue Light
                     ,"#76af95" //Green Dark
                     ,"#f9cb9c" //Orange
                     ,"#af81cf" //Purple
                    ];
  
  //The color used when an erraneous card is entered:
  var errorColor = "#ff0000";
  

  //************************************
  //* You might also want to change these, but these are a bit complicated.
  //* ! Might lead to bugs if not changed correctly !
  //************************************
  

  // These parameters can be used to change the "draft area" in the Draft sheet.
  // E.g. if you want to draft more cards, lastRow should be increased.
  // If you want to start picking two cards at a different time, twoCardRow should be incresed.
  // Don't change the startRow, it will mess things up probably...
  
  //Row in sheet where the first card is drafted. This has to be an odd number! (script uses odd/even to check if we go left or right)
  var startRow = 7;
  
  //Row in sheet where we start picking two cards (the topmost of the two rows)
  //This has to be a row that, if we did not change to two cards, we would go from
  //left to right picking one card.
  //I have not tested what happens if twoCardRow > lastRow, but assume you then get
  //to only do single cards.
  var twoCardRow = 21;
  
  //Row in sheet where the last pick is made.
  var lastRow = 40;
  
  //Column in the sheet where each player is making their picks
  var playerColumn = [3,4,5,6,7,8,9,10];

  //Cell used to give error messages:
  var infoCell = DraftSheet.getRange("C5");


  // These numbes are used to limit the search for card-names in the Cards Tab.
  // If you have a lot of cards in your cube, these might need to be increased.
  // maxCol is the number of columns. If set to 8 it will look in the first 8 columns.
  // maxRow is the number or rows. If set to 150 it will look in the first 150 rows.
  // Ignore these if you are not drafting a cube.
  var maxCol = 8, maxRow = 150;

  //************************************
  //* Dont mess around below here unless you know what you are doing!
  //************************************


  var AllSheets = SpreadsheetApp.getActiveSpreadsheet();
  var CardSheet = AllSheets.getSheetByName("Cards");
  var DraftSheet = AllSheets.getSheetByName("Draft");
  
  var cell = DraftSheet.getActiveCell();
  
  //Get next col/row that shall trigger email notification
  var nextCol = DraftSheet.getSheetValues(51, 3, 1, 1)[0][0];
  var nextRow = DraftSheet.getSheetValues(52, 3, 1, 1)[0][0];
    
  var col = cell.getColumn();
  var row = cell.getRow();
  
  if(nextCol != col || nextRow != row)
  {
    return;//only continue if change was made in the right cell
  }
  
  //Handle multiple picks in a row by the same player (up to 4!):
  var numCards = 1;
  if (row > startRow && row < twoCardRow){
    if (col == playerColumn[0] || col == playerColumn[numPlayers-1]){
      numCards = 2;
    }
  }
  else if (row == twoCardRow+1 && col == playerColumn[0])
  {
    numCards = 3;
  }
  else if (row >= twoCardRow && row <= lastRow)
  {
    if (col==playerColumn[0] || col ==playerColumn[numPlayers-1]){
      numCards = 4;
    }
    else if (col > playerColumn[0] && col < playerColumn[numPlayers-1])
    {
      numCards = 2;
    }
  }
  
  var pickedCards = DraftSheet.getRange(row-numCards+1, col, numCards);
  var pickedCardNames = pickedCards.getValues();
  var infoString = "";
  
  if(pickedCardNames[numCards-1].toString() == "")
  {
    
    return;//only continue if a card has been picked in the "trigger-cell"
  }

  if (isCube)
  {
    //Check Card Sheet if valid card(s) was taken
    var maxCardsPerColor = maxRow;//How many rows to check in the Crads tab while looking for cards being picked
    var error = [0,0,0,0];
    var cardRow = [-1,-1,-1,-1];
    var cardCol = [-1,-1,-1,-1];
    var cardPool = CardSheet.getRange(1, 1, maxRow, maxCol);
    var cardPoolNames = cardPool.getValues();
    var cardPoolBackColor = cardPool.getBackgrounds();
    for(c = 0; c < maxCol; c++)
    {
      if (cardPoolNames[0][c].toString() == "")
      {
        break;//if first cell in column is empty, there are no more cards
      }
      for(r = 0; r < maxRow; r++)
      {
        if (c > maxCardsPerColor)
        {
          break;//we have checked all cards in this collumn
        }
        for (n = numCards-1; n >= 0; n--)
        {
          if(cardRow[n] > 0)//we already found this card...
          {
            continue;
          }
          if (cardPoolNames[r][c].toString() == pickedCardNames[n].toString())
          {
            cardRow[n] = r;
            cardCol[n] = c;
            var isPicked = false
            var pn = 0;
            for (pn = 0; pn < 8; pn++){
              if (cardPoolBackColor[r][c] == playerColor[pn]){
                isPicked = true;
                break;
              }
            }
            if (isPicked)
            {
              error[n] = 10+pn;//Card has been identified, but it has already been picked!
            }else if (cardPoolBackColor[r][c] != "#ffffff")
            {
              error[n] = 2;//This is a heading, not a card...
            }
          }
        }
      }
      var done = 1;
      for(n = numCards-1; n >= 0; n--)
      {
        if (cardRow[n] < 0)
        {
          done = 0;
          break;
        }
      }
      if (done > 0)
      {
        break;
      }
    }
    
    var colors = new Array(numCards);
    for(n = numCards-1; n >= 0; n--)
    {
      colors[n] = new Array(1);
      colors[n][0] = playerColor[col-3];
    }
    
    var allGood = 1;
    var errorInfo = "";
    for(n = numCards-1; n >= 0; n--)
    {
      if(n == numCards-1){
        infoString += pickedCardNames[n].toString();
      }else if(n > 0){
        infoString += ', '+pickedCardNames[n].toString();
      }else {
        infoString += ' og '+pickedCardNames[n].toString();
      }
      if (n == 0)
        infoString += ".";

      if(cardRow[n] < 0)
      {
        colors[n][0] = errorColor;
        errorInfo += pickedCardNames[n].toString() + " was not found in cardlist (spelling/case?). ";
        allGood = 0;
        error[n] = 2;
      }
      else if(error[n] >= 10)
      {
        var pn = error[n] - 10;
        var name = DraftSheet.getSheetValues(6, pn+3, 1, 1)[0][0].toString();
        colors[n][0] = errorColor;
        errorInfo += pickedCardNames[n].toString() + " is already picked by " + name;
        allGood = 0;
      }
      else if(error[n] == 2)
      {
        colors[n][0] = errorColor;
        errorInfo += pickedCardNames[n].toString() + " is a heading, not a card.";
        allGood = 0;
      }
    }
  
    if (allGood == 0)
    {
      if (error[numCards-1] == 0)
      {
        colors[numCards-1][0] = errorColor;
        errorInfo += pickedCardNames[numCards-1].toString()+" was OK, but must be reentered (to trigger script update...)";
        DraftSheet.getRange(row, col).setValue("");
      }
      DraftSheet.getRange(row-numCards+1, col, numCards).setBackgrounds(colors);
      infoCell.setValue(errorInfo);
      infoCell.setBackground(errorColor);
      return;
    }
    DraftSheet.getRange(row-numCards+1, col, numCards).setBackgrounds(colors);
    
    //All possible errors have been checked and are OK! Can notify next player, but first mark the picked cards:
    for(n = numCards-1; n >= 0; n--)
    {
      cardPoolBackColor[cardRow[n]][cardCol[n]] = playerColor[col-3];
      //infoString += ", " + pickedCardNames[n].toString();
    }
    
    CardSheet.getRange(2, 1, maxRow, maxCol).setBackgrounds(cardPoolBackColor);
    
    infoCell.setValue("");
    infoCell.setBackground("white");
  }//end isCube
  else
  {
    //TODO: Check if newly picked cards have been taken earlier?
    //      Error might still happen because of different spelling though
    
    // For now there are no checks, accept whatever has been entered...

    //update background color of cards that was newly picked
    var colors = new Array(numCards);
    for(n = numCards-1; n >= 0; n--)
    {
      if(numCards == 1 || n == numCards-1){
        infoString += pickedCardNames[n].toString();
      }else if(n > 0){
        infoString += ', '+pickedCardNames[n].toString();
      }else {
        infoString += ' og '+pickedCardNames[n].toString() + ".";
      }
      colors[n] = new Array(1);
      colors[n][0] = playerColor[col-3];
    }
    DraftSheet.getRange(row-numCards+1, col, numCards).setBackgrounds(colors);
  }
  
  //Find active player and next player based on which cell has been edited:
  var nextPlayerEmail = '';
  var activePlayer = DraftSheet.getSheetValues(6, col, 1, 1)[0][0].toString();
  var activePlayerEmail = emails[col-3];
  var np = 1337,nr = row;
  if (row >= startRow && row < twoCardRow)//region where we make one pick each
  {
    if (row%2 == 1)//odd numbered row was edited (tells if we are going left or right)
    {
      if (col >= playerColumn[0] && col <= playerColumn[numPlayers-2])
      {
        np = col - 2;
        if (col == playerColumn[numPlayers-2])
          nr = row+1;
      }
    }
    else
    {
      if (col >= playerColumn[1] && col <= playerColumn[numPlayers-1])
      {
        np = col - 4;
        if (col == playerColumn[1]){
          nr = row+1;
          if(row == 26)
            nr = row + 2;
        }
      }
    }
  }
  else if (row >= twoCardRow && row <= lastRow)//we now make two picks each!
  {
    if (row%2 == 0)
    {
      if ((row/2)%2 == 0)
      {
        if (col >= playerColumn[0] && col <= playerColumn[numPlayers-2])
        {
          np = col - 2;
          if (col == playerColumn[numPlayers-2])
            nr = row+2;
        }
      }
      else
      {
        if (col >= playerColumn[1] && col <= playerColumn[numPlayers-1])
        {
          np = col - 4;
          if (col == playerColumn[1])
            nr = row+2;
        }
      }
    }
  }
  
  if (np==1337){
    return;
  }
  
  var nextPlayerName = DraftSheet.getSheetValues(6, np+3, 1, 1)[0][0].toString();
  
  //mark next cell to be filled in a different color:
  //Handle multiple picks in a row by the same player (up to 4!):
  var numCards = 1;
  row = nr;
  col = np+3;
  if (row > startRow && row < twoCardRow){
    if (col==playerColumn[0] || col ==playerColumn[numPlayers-1]){
      numCards = 2;
    }
  }
  else if (row == twoCardRow+1 && col == playerColumn[0])
  {
    numCards = 3;
  }
  else if (row >= twoCardRow && row <= lastRow)
  {
    if (col==playerColumn[0] || col ==playerColumn[numPlayers-1]){
      numCards = 4;
    }
    else if (col >= playerColumn[1] && col <= playerColumn[numPlayers-2])
    {
      numCards = 2;
    }
  }
  var newColor = new Array(numCards);
  for(n = numCards-1; n >= 0; n--)
  {
    newColor[n] = new Array(1);
    newColor[n][0] = "white";
  }
  DraftSheet.getRange(row-numCards+1, col, numCards).setBackgrounds(newColor);

  //Update Next Col/Row info:
  var ncCell = DraftSheet.getRange("C51");
  ncCell.setValue(np+3);
  var nrCell = DraftSheet.getRange("C52");
  nrCell.setValue(nr);
  var npCell = DraftSheet.getRange("C2");
  npCell.setValue(nextPlayerName);
  npCell.setBackground(playerColor[np]);
  
  //Notification stuff:
  nextPlayerEmail = emails[np];
  
  //Mail to next player:
  if(nextPlayerEmail != ""){
    MailApp.sendEmail(nextPlayerEmail,
                      'Din tur i Rotisserie draft',
                      activePlayer+' tok nettopp '+infoString+'\n\nLink:\n'+link+'\n\nCubeTutor:\n'+cubetutor);
  }
  
  //Mail to other players about changes:
  var i = 0;
  for (i = 0; i < 8; i++){
    if (notifyAll[i] && emails[i] != "" && nextPlayerEmail != emails[i] && emails[i] != activePlayerEmail){
      MailApp.sendEmail(emails[i],
                        'Rotisserie draft oppdatering',
                        activePlayer+' tok nettopp '+infoString+'\n\n'+nextPlayerName+' er nestemann.\n\nLink:\n'+link+'\n\nCubeTutor:\n'+cubetutor);
    }
  }  
}
