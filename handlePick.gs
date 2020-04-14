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

  
  //********************************************************************************
  //** SETUP
  //********************************************************************************
  
  
  var AllSheets = SpreadsheetApp.getActiveSpreadsheet();
  var SetupSheet = AllSheets.getSheetByName("Setup");
  var CardSheet = AllSheets.getSheetByName("Cards");
  var DraftSheet = AllSheets.getSheetByName("Draft");
  var infoCell = DraftSheet.getRange("C5");
  
  var numPlayers = Number(SetupSheet.getRange("C2").getValue());
  var isCube = SetupSheet.getRange("C3").getValue();
  
  var names = ['','','','','','','',''];
  var emails = ['','','','','','','',''];
  var notifyTurn  = [false,false,false,false,false,false,false,false];
  var includeAll  = [false,false,false,false,false,false,false,false];
  var notifyAll   = [false,false,false,false,false,false,false,false];
  var notifyFinish= [false,false,false,false,false,false,false,false];
  var playerColor = [ "#7cb6eb","#b6d7a8","#ffd966","#f7bcd5","#81d7eb","#76af95","#f9cb9c","#af81cf"];
  
  var backgrounds = SetupSheet.getRange(9, 4, 8, 1).getBackgrounds();
  var info = SetupSheet.getRange(9, 2, 8, 7).getValues();
  for(n = 0; n < 8; n++){
    names[n] = info[n][0];
    emails[n] = info[n][1];
    notifyTurn[n] = info[n][3];
    includeAll[n] = info[n][4];
    notifyAll[n] = info[n][5];
    notifyFinish[n] = info[n][6];
    playerColor[n] = backgrounds[n][0];
  }
  
  var errorColor = SetupSheet.getRange("B32").getBackground();
  var draftColor = SetupSheet.getRange("B33").getBackground();
  var defaultColor = SetupSheet.getRange("B34").getBackground();
  var unusedColor = SetupSheet.getRange("B35").getBackground();
  var cardHeaderColor = SetupSheet.getRange("B36").getBackground();
  
  var notifyTurnSubject = SetupSheet.getRange("B20").getValue();
  var notifyAllSubject = SetupSheet.getRange("B21").getValue();
  var notifyFinishSubject = SetupSheet.getRange("B22").getValue();
  var sheetLink = '\n\nLink:\n'+SetupSheet.getRange("B23").getValue();
  var cube = SetupSheet.getRange("B24").getValue()
  var cubeLink = '';
  if (cube != '')
    cubeLink = '\n\nCube:\n'+cube;
  var and = SetupSheet.getRange("C26").getValue();
  var sinceLast = SetupSheet.getRange("C27").getValue();
  var isNext = SetupSheet.getRange("C28").getValue();
  var justPicked = SetupSheet.getRange("C29").getValue();
  
  var errorForgot = SetupSheet.getRange("D50").getValue();
  var errorSameCard = SetupSheet.getRange("D51").getValue();
  var errorNotFound = SetupSheet.getRange("D52").getValue();
  var errorAlreadyTaken = SetupSheet.getRange("D53").getValue();
  var errorIsHeading = SetupSheet.getRange("D54").getValue();
  var errorMaybeReenter = SetupSheet.getRange("D55").getValue();
  var errorOKReenter = SetupSheet.getRange("D56").getValue();
  
  var cardsDrafted = Number(SetupSheet.getRange("C4").getValue());
  var twoCardPick = SetupSheet.getRange("C5").getValue();
  
  //Row in sheet where the first card is drafted. This has to be an odd number! (script uses odd/even to check if we go left or right)
  var startRow = 7;
  
  //Row in sheet where we start picking two cards (the topmost of the two rows)
  var twoCardRow = 50;
  if (twoCardPick != "Never")
    twoCardRow = startRow + Number(twoCardPick) - 1;
  
  //Row in sheet where the last pick is made.
  var lastRow = startRow + cardsDrafted - 1;
  
  //Column in the sheet where each player is making their picks
  var playerColumn = [3,4,5,6,7,8,9,10];

  // These numbes are used to limit the search for card-names in the Cards Tab.
  var maxCol = 8;
  var maxRow = Number(SetupSheet.getRange("B40").getValue());

  //***************************************************************************************
  //** Logic for setting up sheet:
  //***************************************************************************************
  

  var resetCells = SetupSheet.getRange(44, 2, 3).getValues();
  if (resetCells[0][0] && resetCells[1][0] && resetCells[2][0])
  {
    //DraftSheet:
    infoCell.setBackground("white");
    infoCell.setValue("");
    var ncCell = DraftSheet.getRange("C51");
    ncCell.setValue(playerColumn[0]);
    var nrCell = DraftSheet.getRange("C52");
    nrCell.setValue(startRow);
    var npCell = DraftSheet.getRange("C2");
    npCell.setValue(names[0]);
    npCell.setBackground(playerColor[0]);
    var pickInfoCell = DraftSheet.getRange("E4");
    if (twoCardRow < lastRow)
      pickInfoCell.setValue("You will make " + cardsDrafted +" picks each. Starting with pick " + twoCardPick + " each player will pick two cards at once.");
    else
      pickInfoCell.setValue("You will make " + cardsDrafted +" picks each.");
    DraftSheet.getRange(startRow, playerColumn[0], lastRow-startRow+1, numPlayers).setBackground(defaultColor);
    DraftSheet.getRange(startRow, playerColumn[0], lastRow-startRow+1, 8).setValue("");
    DraftSheet.getRange(startRow, playerColumn[0]).setBackground(draftColor);
    if (numPlayers < 8)
    {
      DraftSheet.getRange(startRow, playerColumn[numPlayers], lastRow-startRow+1, 8-numPlayers).setBackground(unusedColor);
    }
    if (lastRow < 42)
    {
      DraftSheet.getRange(lastRow+1, playerColumn[0], 48-lastRow, 8).setBackground(unusedColor);
    }
    var headersText = DraftSheet.getRange(startRow-1, playerColumn[0],1,8).getValues();
    var headersBack = DraftSheet.getRange(startRow-1, playerColumn[0],1,8).getBackgrounds();
    for (n = 0; n < 8; n++)
    {
      headersText[0][n] = names[n];
      headersBack[0][n] = playerColor[n];
    }
    DraftSheet.getRange(startRow-1, playerColumn[0],1,8).setValues(headersText);
    DraftSheet.getRange(startRow-1, playerColumn[0],1,8).setBackgrounds(headersBack);
    DraftSheet.getRange(startRow, 13, 32, 2).setValue("");

    
    //CardSheet:
    var CubeCardBack = CardSheet.getRange(1, 1, maxRow, maxCol).getBackgrounds();
    var r=0,c=0,n=0;
    for(r = 0; r< maxRow; r++){
      for (c = 0; c < maxCol; c++){
        for(n = 0; n < 8; n++){
          if(CubeCardBack[r][c] != cardHeaderColor)
          {
            CubeCardBack[r][c] = "white";
          }
        }
      }
    }
    CardSheet.getRange(1, 1, maxRow, maxCol).setBackgrounds(CubeCardBack);
    var playerRef = CardSheet.getRange(19, 10, 8).getValues();
    var playerRefBG = CardSheet.getRange(19, 10, 8).getBackgrounds();
    for(n = 0; n < 8; n++){
      playerRef[n][0] = names[n];
      playerRefBG[n][0] = playerColor[n];
    }
    CardSheet.getRange(19, 10, 8).setValues(playerRef);
    CardSheet.getRange(19, 10, 8).setBackgrounds(playerRefBG);
    
    //SetupSheet:
    resetCells[0][0] = false;
    resetCells[1][0] = false;
    resetCells[2][0] = false;
    SetupSheet.getRange(44, 2, 3).setValues(resetCells);
  }


  //*****************************************************************************
  //* Logic for checking validity of cards drafted
  //*****************************************************************************
  

  //get col/row that was edited now:  
  var cell = DraftSheet.getActiveCell();
  var col = cell.getColumn();
  var row = cell.getRow();

  //Get next col/row that shall trigger script:
  var nextCol = DraftSheet.getSheetValues(51, 3, 1, 1)[0][0];
  var nextRow = DraftSheet.getSheetValues(52, 3, 1, 1)[0][0];
    
  //Check if the correct cell was edited:
  if(nextCol != col || nextRow != row)
  {
    return; //Incorrect cell edited. Script ends with no further actions.
  }
  
  //An edit has been made in the correct cell! Now we must figure out what
  //cards were selected, and check if these cards are OK.
  
  //Handle multiple picks in a row by the same player (up to 4)
  //Make a list of the cards picked:
  var numCardsPicked = 1;
  if (row == lastRow && col == playerColumn[0])
  {
    numCardsPicked = 2;
    if (twoCardRow > lastRow)
      numCardsPicked = 1;
  }
  else if (row > startRow && row < twoCardRow){
    if (col == playerColumn[0] || col == playerColumn[numPlayers-1]){
      numCardsPicked = 2;
    }
  }
  else if (row == twoCardRow+1 && col == playerColumn[0])
  {
    numCardsPicked = 3;
  }
  else if (row >= twoCardRow && row <= lastRow)
  {
    if (col==playerColumn[0] || col ==playerColumn[numPlayers-1]){
      numCardsPicked = 4;
    }
    else if (col > playerColumn[0] && col < playerColumn[numPlayers-1])
    {
      numCardsPicked = 2;
    }
  }
  
  var pickedCards = DraftSheet.getRange(row-numCardsPicked+1, col, numCardsPicked);
  var pickedCardNames = pickedCards.getValues();
  var infoString = "";
  
  //Logic for card validation and updates in the sheet (background colors):
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
    
    for(n = numCardsPicked-1; n >= 0; n--)
    {
      if(pickedCardNames[n].toString() == "")
      {
        error[n] = 4;
      }
      for(m = n-1; m >= 0; m--)
      {
        if(pickedCardNames[n].toString() == pickedCardNames[m].toString())
        {
          error[n] = 5;
        }
      }
    }
    for(c = 0; c < maxCol; c++)
    {
      if (cardPoolNames[0][c].toString() == "")
      {
        break;//if first cell in column is empty, there are no more cards
      }
      for(r = 0; r < maxRow; r++)
      {
        for (n = numCardsPicked-1; n >= 0; n--)
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
      for(n = numCardsPicked-1; n >= 0; n--)
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
    
    var colors = new Array(numCardsPicked);
    for(n = numCardsPicked-1; n >= 0; n--)
    {
      colors[n] = new Array(1);
      colors[n][0] = playerColor[col-3];
    }
    
    var allGood = 1;
    var errorInfo = "";
    for(n = numCardsPicked-1; n >= 0; n--)
    {
      if(n == numCardsPicked-1){
        infoString += pickedCardNames[n].toString();
      }else if(n > 0){
        infoString += ', '+pickedCardNames[n].toString();
      }else {
        infoString += ' '+and+' '+pickedCardNames[n].toString();
      }
      if (n == 0)
        infoString += ".";

      if(error[n] == 4)
      {
        colors[n][0] = errorColor;
        errorInfo += errorForgot +". ";
        allGood = 0;
      }
      else if(error[n] == 5)
      {
        colors[n][0] = errorColor;
        errorInfo += errorSameCard +". ";
        allGood = 0;
      }
      else if(cardRow[n] < 0)
      {
        colors[n][0] = errorColor;
        errorInfo += pickedCardNames[n].toString() + " "+errorNotFound +". ";
        allGood = 0;
        error[n] = 2;
      }
      else if(error[n] >= 10)
      {
        var pn = error[n] - 10;
        var name = DraftSheet.getSheetValues(6, pn+3, 1, 1)[0][0].toString();
        colors[n][0] = errorColor;
        errorInfo += pickedCardNames[n].toString() + " "+ errorAlreadyTaken +" " + name + ". ";
        allGood = 0;
      }
      else if(error[n] == 2)
      {
        colors[n][0] = errorColor;
        errorInfo += pickedCardNames[n].toString() + " "+errorIsHeading +". ";
        allGood = 0;
      }
    }
  
    if (allGood == 0)
    {
      if (error[numCardsPicked-1] == 0)
      {
        colors[numCardsPicked-1][0] = errorColor;
        errorInfo += " " + pickedCardNames[numCardsPicked-1].toString()+" "+errorOKReenter;
        DraftSheet.getRange(row, col).setValue("");
      }
      DraftSheet.getRange(row-numCardsPicked+1, col, numCardsPicked).setBackgrounds(colors);
      infoCell.setValue(errorInfo);
      infoCell.setBackground(errorColor);
      return;
    }
    DraftSheet.getRange(row-numCardsPicked+1, col, numCardsPicked).setBackgrounds(colors);
    
    //All possible errors have been checked and are OK! Can notify next player, but first mark the picked cards:
    for(n = numCardsPicked-1; n >= 0; n--)
    {
      cardPoolBackColor[cardRow[n]][cardCol[n]] = playerColor[col-3];
    }
    
    CardSheet.getRange(1, 1, maxRow, maxCol).setBackgrounds(cardPoolBackColor);
    
    infoCell.setValue("");
    infoCell.setBackground("white");
  }
  else
  { //we go here if isCube = false
    //TODO: Check if newly picked cards have been taken earlier?
    //      There migh be different spelling though, so probably no point checking
    
    //Check that a name was entered in all cells, and that the same name was not entered twize:
    var error = [0,0,0,0];
    for(n = numCardsPicked-1; n >= 0; n--)
    {
      if(pickedCardNames[n].toString() == "")
      {
        error[n] = 4;
      }
      for(m = n-1; m >= 0; m--)
      {
        if(pickedCardNames[n].toString() == pickedCardNames[m].toString())
        {
          error[n] = 5;
        }
      }
    }
    
    //Update color of cells etc:
    var colors = new Array(numCardsPicked);
    var errorText = "";
    for(n = numCardsPicked-1; n >= 0; n--)
    {
      if(numCardsPicked == 1 || n == numCardsPicked-1){
        infoString += pickedCardNames[n].toString();
      }else if(n > 0){
        infoString += ', '+pickedCardNames[n].toString();
      }else {
        infoString += ' '+and+' '+pickedCardNames[n].toString() + ".";
      }
      colors[n] = new Array(1);
      if (error[n] == 4){
        colors[n][0] = errorColor;
        errorText += errorForgot +". ";
      }
      else if (error[n] == 5){
        colors[n][0] = errorColor;
        errorText += errorSameCard +". ";
      }
      else {
        colors[n][0] = playerColor[col-3];
      }
    }
    DraftSheet.getRange(row-numCardsPicked+1, col, numCardsPicked).setBackgrounds(colors);
    if (errorText != ""){
      errorText += "("+ errorMaybeReenter +")";
      infoCell.setBackground(errorColor);
      infoCell.setValue(errorText);
      return;
    }
    else
    {
      infoCell.setBackground("white");
      infoCell.setValue("");
    }
  }
  
  
  //**************************************************************************
  //** Logic for sending notifications and update sheet for next player
  //**************************************************************************
  
  //Find active player and next player based on which cell has been edited:
  var np = 1337,nr = row;
  if (row >= startRow && row < twoCardRow)//region where we make one pick each
  {
    if (row%2 == 1)//odd numbered row was edited (tells us we are going to the right)
    {
      if (col >= playerColumn[0] && col <= playerColumn[numPlayers-2])
      {
        np = col - 2;
        if (col == playerColumn[numPlayers-2])
          nr = row+1;
      }
    }
    else //to the left
    {
      if (col >= playerColumn[1] && col <= playerColumn[numPlayers-1])
      {
        np = col - 4;
        if (row != lastRow && col == playerColumn[1]){
          nr = row+1;
          if(row == twoCardRow-1)
            nr = row + 2;
        }
      }
    }
  }
  else if (row >= twoCardRow && row <= lastRow)//we now make two picks each!
  {
    if (((row-twoCardRow-1)/2)%2 == 0)//moving to the right
    {
      if (col >= playerColumn[0] && col <= playerColumn[numPlayers-2])
        {
          np = col - 2;
          if (col == playerColumn[numPlayers-2])
            nr = row+2;
        }
    }
    else//to the left
    {
      if (col >= playerColumn[1] && col <= playerColumn[numPlayers-1])
      {
        np = col - 4;
        if (row != lastRow && col == playerColumn[1])
          nr = row+2;
      }
    }
  }
  
  //check if draft is over:
  var draftOver = false;
  if(col == playerColumn[0] && row == lastRow){
    draftOver = true;
  }
  else if (np==1337){
    infoCell.setValue("An error occured !!!!!  row = " + row + " col = " + col);
    return;
  }

  var nextPlayerName = DraftSheet.getSheetValues(6, np+3, 1, 1)[0][0].toString();
  var nextPlayerEmail = emails[np];
  var activePlayerName = DraftSheet.getSheetValues(6, col, 1, 1)[0][0].toString();
  var activePlayerEmail = emails[col-3];
  
  if(draftOver){
    //Mail to all players
    var i = 0;
    for (i = 0; i < 8; i++)
    {
      if (notifyFinish[i] && emails[i] != '')
      {
        MailApp.sendEmail(emails[i],
                          notifyFinishSubject,
                          activePlayerName+' '+justPicked+' '+infoString+sheetLink);
      }
    }
    //Update Next Col/Row info:
    var ncCell = DraftSheet.getRange("C51");
    ncCell.setValue("done");
    var nrCell = DraftSheet.getRange("C52");
    nrCell.setValue("done");
    var npCell = DraftSheet.getRange("C2");
    npCell.setValue("Finished");
    npCell.setBackground("white");
  }
  else{
    
    //mark next cell to be filled in a different color:
    //Handle multiple picks in a row by the same player (up to 4!):
    var numCardsToDraft = 1;
    row = nr;
    col = np+3;
    if (row == lastRow && col == playerColumn[0])
    {
      numCardsToDraft = 2;
      if (twoCardRow > lastRow)
        numCardsToDraft = 1;
    }
    else if (row > startRow && row < twoCardRow){
      if (col==playerColumn[0] || col ==playerColumn[numPlayers-1]){
        numCardsToDraft = 2;
      }
    }
    else if (row == twoCardRow+1 && col == playerColumn[0])
    {
      numCardsToDraft = 3;
    }
    else if (row >= twoCardRow && row <= lastRow)
    {
      if (col==playerColumn[0] || col ==playerColumn[numPlayers-1]){
        numCardsToDraft = 4;
      }
      else if (col >= playerColumn[1] && col <= playerColumn[numPlayers-2])
      {
        numCardsToDraft = 2;
      }
    }
    var newColor = new Array(numCardsToDraft);
    for(n = numCardsToDraft-1; n >= 0; n--)
    {
      newColor[n] = new Array(1);
      newColor[n][0] = "white";
    }
    DraftSheet.getRange(row-numCardsToDraft+1, col, numCardsToDraft).setBackgrounds(newColor);
    
    //Update Next Col/Row info:
    var ncCell = DraftSheet.getRange("C51");
    ncCell.setValue(np+3);
    var nrCell = DraftSheet.getRange("C52");
    nrCell.setValue(nr);
    var npCell = DraftSheet.getRange("C2");
    npCell.setValue(nextPlayerName);
    npCell.setBackground(playerColor[np]);
    

    //Logic for handling pick history
    var draftHistory = DraftSheet.getRange(startRow, 13, numPlayers*4, 2).getValues();
    var updatedHistory = DraftSheet.getRange(startRow, 13, numPlayers*4, 2).getValues();
    var prevPickInfo = "\n";
  
    //Find last time next player made a pick (if any)
    //All parts of the history after that point should be added to the previous pick info
    //If this is the first pick of that player, all picks so far shall be added
    //At the same time, set up the updated history (remove oldest picks, add newest).
    var historyName = "Balle";
    var historyLength = 0;
    var firstCards = 0;
    var firstName = draftHistory[0][0];
    var onFirst = true;
    var lastR = 0;
    for(r = 0; r < numPlayers*4; r++)
    {
      if (draftHistory[r][0] == "")
      {
        lastR = r;
        break;//no more historic picks available...
      }

      if (draftHistory[r][0] != historyName)
      {
        if (historyName == firstName)
          onFirst = false;
        historyLength += 1;
        historyName = draftHistory[r][0];
        if(draftHistory[r][0] == nextPlayerName)
        {
          prevPickInfo = "\n";
        }
        else
        {
          prevPickInfo += '\n'+historyName+':\n\t'+draftHistory[r][1];
        }
      }
      else
      {
        if(draftHistory[r][0] == draftHistory[r+1][0])
          prevPickInfo += '\n\t'+draftHistory[r][1];
        else
          prevPickInfo += '\n\t'+draftHistory[r][1];
      }
      
      if (onFirst && draftHistory[r][0] == firstName)
      {
        firstCards += 1
      }
      else
      {
        updatedHistory[r-firstCards] = draftHistory[r];
      }
    }
    prevPickInfo += '\n'+activePlayerName + ":";
    var separator = ['\n\t',
                     '\n\t',
                     '\n\t',
                     '\n\t'];
    if(numCardsPicked > 1)
      separator[numCardsPicked-1] = ' '+and+' ';
    if(historyLength < ((numPlayers*2)-4))
    {
      for (n = 0; n < numCardsPicked; n++)
      {
        draftHistory[lastR + n][0] = activePlayerName;
        draftHistory[lastR + n][1] = pickedCardNames[n];
        prevPickInfo += separator[n]+pickedCardNames[n];
      }
      DraftSheet.getRange(startRow, 13, numPlayers*4, 2).setValues(draftHistory);
    }
    else
    {
      for (n = 0; n < firstCards; n++)
      {
        updatedHistory[lastR + n-firstCards][0] = '';
        updatedHistory[lastR + n-firstCards][1] = '';
      }
      for (n = 0; n < numCardsPicked; n++)
      {
        updatedHistory[lastR + n-firstCards][0] = activePlayerName;
        updatedHistory[lastR + n-firstCards][1] = pickedCardNames[n];
        prevPickInfo += separator[n]+pickedCardNames[n];
      }
      DraftSheet.getRange(startRow, 13, numPlayers*4, 2).setValues(updatedHistory);
    }


    //Mail to next player:
    if(notifyTurn[np] && nextPlayerEmail != "")
    {
      if(includeAll[np])
      {
        MailApp.sendEmail(nextPlayerEmail,
                          notifyTurnSubject,
                          sinceLast+prevPickInfo+sheetLink+cubeLink);
      }
      else
      {
        MailApp.sendEmail(nextPlayerEmail,
                          notifyTurnSubject,
                          activePlayerName+' '+justPicked+' '+infoString+sheetLink+cubeLink);
      }
    }
    
    //Mail to other players about changes:
    var i = 0;
    for (i = 0; i < 8; i++){
      if (notifyAll[i] && emails[i] != "" && nextPlayerEmail != emails[i] && emails[i] != activePlayerEmail)
      {
        MailApp.sendEmail(emails[i],
                          notifyAllSubject,
                          activePlayerName+' '+justPicked+' '+infoString+'\n\n'+nextPlayerName+' '+isNext+sheetLink);
      }
    }  
  }
}
