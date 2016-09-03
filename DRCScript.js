/* 
   Issues
   
   ** Occasionally get "Execution failure a server error occured." **
   ** Seems to just be a Google script server issue. **
   ** Won't be a problem as the triggered functions will continue on during the next triggered event. **

   ** appendRoW() and insertRow() appear to cause a concurrent ghost thread to spawn sporadically. **
   ** Setup private locks to try and stop any ghost threads. Seems to work ** 

   ** Need to modularize redundant code blocks. **

   ** Archived log could get very large and cause problems in the future. Perhaps it should be seperated from ** 
   ** the rest of the spreadsheet during archiving and saved to another google drive spreadsheet. ** 
   
*/


// Constants for sheets.
var DATA_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("uids");
var FORM_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
var SETTINGS_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
var TEMPLATE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Templates");
var LOG_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
var ARCHIVE_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Archive");
var CURRENT_USER = Session.getActiveUser().getEmail();
var START_ROW = SETTINGS_SHEET.getRange('B1').getValue(); //This is the row where the data starts (2 since there is a header row)
var UPDATE_ROW = SETTINGS_SHEET.getRange('B2').getValue();
var TERM_NAME = SETTINGS_SHEET.getRange('B7').getValue();
var ROOT_OWNER = SETTINGS_SHEET.getRange('B3').getValue();
var ROOT_FOLDER = SETTINGS_SHEET.getRange('B4').getValue();
var GROUP_EMAIL = SETTINGS_SHEET.getRange('B5').getValue();
var FORM_ID = SETTINGS_SHEET.getRange('B6').getValue();

// Constants for uids sheet
var STATUS_COL = "E";
var UPDATED_COL = "D";
var ODIN_COL = "A";
var FIRSTNAME_COL = "B";
var LASTNAME_COL = "C";

var lock = LockService.getPrivateLock();
var pubLock = LockService.getPublicLock();


// The main function, creates new user folders and initiates syncronzing of files and folders.
function main() {
  
  // Stop any concurrent script running by each user.
  try 
  {
        lock.waitLock(1000);
  } 
  
  catch(err) 
  {
     Logger.log('Could not obtain lock after 1 seconds.');
  }
  
  
  // Check to make sure there isn't a public lock, from say archiving or
  // removing a user.
  if(pubLock.hasLock() == true)
  {
     var ui = SpreadsheetApp.getUi();
    
     ui.alert("Document is locked for archiving or user removal. Please try again later.");
     return false;
  }
  
  // Since the function will need to be run several times (with multiple users) before completion
  // we need to save our spot in the data.
  var userProperties = PropertiesService.getUserProperties();
  
  // Can't add a trigger lower than 10 minutes or sporadic double script running occurs.
  if(findTriggerByHandler("main") == false)
  {
        ScriptApp.newTrigger("main")
        .timeBased()
        .everyMinutes(5) 
        .create();
        
        userProperties.setProperty('CURRENT_DATA_ROW', 0);
  }

  // Setup variables so we can track time and our spot in the data structure.
  var curr_data_row = parseInt(userProperties.getProperty('CURRENT_DATA_ROW')); 
  var startTime = new Date();
  var currentTime = "";
  var root = getRoot();  
  
  var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
 
  // Create one JavaScript object per row of data.
  objects = getRowsData(DATA_SHEET, dataRange);
       
  for(; curr_data_row < objects.length; ++curr_data_row) 
  {
     try 
     {
        lock.waitLock(1000);
     } 
  
     catch(err) 
     {
        Logger.log('Could not obtain lock after 1 seconds.');
     }
    
    
     // We need to wait 1 second before we try to set a new value
     // to a user property field or we risk hitting the max number
     // of times we are allowed to modify the field in a set time.
     Utilities.sleep(1000);
     userProperties.setProperty('CURRENT_DATA_ROW', curr_data_row);
    
     currentTime = new Date();
 
     // Keep track of the time so we can exit before the script time of 6 minutes.    
     if((currentTime.getTime() - startTime.getTime()) < 240000)
     {
        try 
        {
           lock.waitLock(1000);
        } 
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        }
        
       
        // Get a row object from the spreadsheet
        var rowData = objects[curr_data_row];
       
        // For testing.
        Logger.log("User #: " + curr_data_row);    
        Logger.log("First name: " + rowData.first);
        Logger.log("Last name: " + rowData.last);
        Logger.log("Odin: " + rowData.odin);
        Logger.log("Last Update: " + rowData.updated);
    
        // Find the users folder so we can work on it.
        var user_folder = findUserFolder(rowData);
     
        // As long as we found the user folder, amd their previous transfer
        // didn't fail, attempt to transfer files.
        if(user_folder != null && rowData.status != "Failed")
        {
           try 
           {
              lock.waitLock(1000);
           } 
  
           catch(err) 
           {
              Logger.log('Could not obtain lock after 1 seconds.');
           }
          
           Logger.log("Folder exists");
           xferFiles(user_folder, rowData);   
        }
  
        // No user folder found, we need to create one if we are the admin.
        else if(user_folder === null && CURRENT_USER === ROOT_OWNER)
        {
           Logger.log("Folder does not exist, creating folder.");
           setNewUser(null, rowData);
           sendWelcomeEmail(rowData);
        }
        
        try 
        {
           lock.waitLock(1000);
        } 
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        }
       
        // Update the timestamp for the user.
        var currentTime = (new Date()).toString();
        DATA_SHEET.getRange(UPDATE_ROW + (curr_data_row + 2)).setValue(currentTime);
        SpreadsheetApp.flush();
     
     }
    
     else
     {
         return false;
     }
  }
   
  Logger.log("Made it to the end of rowData, resetting current_data_row for user.");
  
  try 
  {
     lock.waitLock(1000);
  } 
  
  catch(err) 
  {
     Logger.log('Could not obtain lock after 1 seconds.');
  }
  
  Utilities.sleep(1000);
  userProperties.setProperty('CURRENT_DATA_ROW', 0);
  
  SpreadsheetApp.getUi().alert('Manual file transfer finished.');
  
  removeTriggerByHandler("main");
  lock.releaseLock();
  return true;    
};



// Used to setup triggers and script variables for a new staff member.
function newGroupMember()
{
   // Need to set this so that the user can keep track of the data row they were
   // processing if the script fails.
   var userProperties = PropertiesService.getUserProperties();
   
   Utilities.sleep(1000);
   userProperties.setProperty('CURRENT_DATA_ROW', 0);
     
};



// Only administrator should be allowed to run this function.
function removeUser()
{
  
    // Stop any concurrent script running by each user.
    try 
    {
        lock.waitLock(1000);
    } 
  
    catch(err) 
    {
       Logger.log('Could not obtain lock after 1 seconds.');
    }
   
    // Check to make sure there isn't a public lock, from say archiving or
    // removing a user.
    if(pubLock.hasLock() == true)
    {    
      var ui = SpreadsheetApp.getUi();
      
      ui.alert("Document is locked for archiving or user removal. Please try again later.");
      return false;
    }
  
    else
    {
        try 
        {
           pubLock.waitLock(1000);
        } 
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        } 
    }
  

   // Setting up a UI box with a text field for user input.
   var ui = SpreadsheetApp.getUi();

   var result = ui.prompt(
     'Enter ODIN username:',
      ui.ButtonSet.OK_CANCEL);

   // Process the user's response.
   var button = result.getSelectedButton();
   var username = result.getResponseText();
  
   if (button == ui.Button.OK && username != "") 
   {
      // User clicked "OK".
      ui.alert('Removing user ' + username + '.');
   }
  
   else if (button == ui.Button.CANCEL) 
   {
      // User clicked "Cancel".
      ui.alert('Remove user cancelled.');
      return false;
   } 
  
   else if (button == ui.Button.OK && username === "") 
   {
      // User didn't enter a username.
      ui.alert('Please enter a valid ODIN username.');
      return false;
   } 
  
   var scriptProperties = PropertiesService.getScriptProperties();
   scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', 0);
   var curr_archive_row = parseInt(scriptProperties.getProperty("CURRENT_ARCHIVE_ROW")); 
   Logger.log("Current Archive row in removeUser is: " + curr_archive_row);
   
   var root = getRoot();
   var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
 
   // Create one JavaScript object per row of data.
   objects = getRowsData(DATA_SHEET, dataRange);
  
   // Search for the username to remove.
   while(curr_archive_row < objects.length) 
   {   
      // Get a row object from the spreadsheet
      var rowData = objects[curr_archive_row];
      
      // For testing.
      Logger.log("User #: " + curr_archive_row);    
      Logger.log("First name: " + rowData.first);
      Logger.log("Last name: " + rowData.last);
      Logger.log("Odin: " + rowData.odin);
      Logger.log("Last Update: " + rowData.updated);
  
      if(rowData.odin === username)
      {
         var user_folder = findUserFolder(rowData);
         var user_email = rowData.odin + "@pdx.edu";
        
         if(user_folder != null && user_folder.getOwner().getEmail() === CURRENT_USER && rowData.status != "Failed")
         {
           
            scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', curr_archive_row);
            var dog = parseInt(scriptProperties.getProperty("CURRENT_ARCHIVE_ROW")); 
            Logger.log("Folder exists");
           
            Logger.log("Script Current Archive Row is now: " + dog); 
            xferFiles(user_folder, rowData);
            removeEditorWalkDirectoryTree(user_folder, rowData, rowData.files);
            
            // Data for log file and email that will be sent to user.
            var url = user_folder.getUrl();
            var file_name = user_folder.getName();
            var last_updated = user_folder.getLastUpdated();
            var prev_owner = user_folder.getOwner().getEmail();

            Logger.log("File name is: " + file_name);
             
            try
            {
                // Try to transfer ownership.
                user_folder.setOwner(user_email);
                var new_owner = user_folder.getOwner().getEmail();
              
                if(user_folder.getAccess(GROUP_EMAIL) === DriveApp.Permission.EDIT)
                {
                   user_folder.removeEditor(GROUP_EMAIL);
                }
              
                user_folder.removeEditor(CURRENT_USER);
                
                var note = "Successfully transfered root folder ownership.";
              
                updateLog(user_email, user_folder.getName(), new_owner, prev_owner, last_updated, url, note);
               
                // Clear the data sheet.
                DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("");
                DATA_SHEET.getRange(UPDATED_COL + (curr_archive_row + 2)).setValue("");
                DATA_SHEET.deleteRows(curr_archive_row + 2, 1);
                SpreadsheetApp.flush();
              
                // Clear the form response row.
                FORM_SHEET.deleteRows(curr_archive_row + 2, 1);
                SpreadsheetApp.flush();
                
                ui.alert('Successfully removed ' + username + ".");
                scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', 0);
                pubLock.releaseLock();
                return true;            
             }
           
             catch(err)
             {
                // If the ownership tranfer fails, send an email about it, record in the log, and set the
                // user's status as failed so it won't happen again until the failure has been cleared.
                Logger.log("There was an error transfering root user folder ownership!");
        
                var new_owner = user_folder.getOwner().getEmail();
                var note = "Failed to transfer root folder.";        
                updateLog(user_email, user_folder.getName(), new_owner, prev_owner, last_updated, url, note);
                sendErrorEmail(rowData);
                DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Failed");
                SpreadsheetApp.flush();
               
                scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', 0);
                pubLock.releaseLock();
                return false;
             }
         }
  
         // The user has no folder, just delete their spreadsheet entries.
         else if(user_folder === null && CURRENT_USER === ROOT_OWNER)
         {
            Logger.log("Folder does not exist.");
 
            // Clear the data sheet items.
            DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("");
            DATA_SHEET.getRange(UPDATED_COL + (curr_archive_row + 2)).setValue("");
            DATA_SHEET.deleteRows(curr_archive_row + 2, 1);   
            SpreadsheetApp.flush();
              
            // Clear the form response row.
            FORM_SHEET.deleteRows(curr_archive_row + 2, 1);
            SpreadsheetApp.flush();
           
            scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', 0);
            pubLock.releaseLock();
            ui.alert('Successfully removed ' + username + ".");
            return true;           
         }     
      }
     
      ++curr_archive_row;
   }
  
   ui.alert('ODIN username not found, or an error occured during ownership transfer.');
   scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', 0);
   pubLock.releaseLock();
   lock.releaseLock();
   return false;  
};



// Allows syncing of a single user
// Any staff member can run this.
// But administrator will be the only one able to create the users directory.
function syncUser()
{
   // Stop any concurrent script running by each user.
   try 
   {
       lock.waitLock(1000);
   } 
  
   catch(err) 
   {
       Logger.log('Could not obtain lock after 1 seconds.');
   }
  
  
   // Check to make sure there isn't a public lock, from say archiving or
   // removing a user.
   if(pubLock.hasLock() == true)
   {    
      var ui = SpreadsheetApp.getUi();
      
      ui.alert("Document is locked for archiving or user removal. Please try again later.");
      return false;
   }
     
   // Setup a UI for user input.
   var ui = SpreadsheetApp.getUi(); 
   var result = ui.prompt(
     'Enter ODIN username:',
      ui.ButtonSet.OK_CANCEL);

   // Process the user's response.
   var button = result.getSelectedButton();
   var username = result.getResponseText();
  
   if (button == ui.Button.OK && username != "") 
   {
      // User clicked "OK".
     ui.alert('Syncing user ' + username + '.');
   }
  
   else if (button == ui.Button.CANCEL) 
   {
      // User clicked "Cancel".
      ui.alert('Sync user cancelled.');
      return false;
   } 
  
   else if (button == ui.Button.OK && username === "") 
   {
      // User didn't enter a username.
      ui.alert('Please enter a valid ODIN username.');
      return false;
   } 
      
   var userProperties = PropertiesService.getUserProperties();

   Utilities.sleep(2000);
   userProperties.setProperty('CURRENT_DATA_ROW', 0);
   var curr_data_row = parseInt(userProperties.getProperty("CURRENT_DATA_ROW")); 
  
   var root = getRoot();
   var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
 
   // Create one JavaScript object per row of data.
   objects = getRowsData(DATA_SHEET, dataRange);
  
   // Search for the user to sync.
   while(curr_data_row < objects.length) 
   {
    
      // Get a row object from the spreadsheet
      var rowData = objects[curr_data_row];
        
      // For testing.
      Logger.log("User #: " + curr_data_row);    
      Logger.log("First name: " + rowData.first);
      Logger.log("Last name: " + rowData.last);
      Logger.log("Odin: " + rowData.odin);
      Logger.log("Last Update: " + rowData.updated);
  
      if(rowData.odin === username)
      {
         Utilities.sleep(2000);
         userProperties.setProperty('CURRENT_DATA_ROW', curr_data_row);     
        
         var user_folder = findUserFolder(rowData);
     
         // If we found the user, their folder, and their status isn't failed then
         // transfer the files.
         if(user_folder != null && rowData.status != "Failed")
         {
            Logger.log("Folder exists");
            xferFiles(user_folder, rowData); 
         }
  
         // If the user does not have a folder and we are the admin create the folder.
         else if(user_folder === null && CURRENT_USER === ROOT_OWNER)
         {        
            Logger.log("Folder does not exist, creating folder.");
            setNewUser(null, rowData);      
         }
       
         // Update the last updated cell.
         var currentTime = (new Date()).toString();
         DATA_SHEET.getRange(UPDATE_ROW + (curr_data_row + 2)).setValue(currentTime);
         SpreadsheetApp.flush();
        
         ui.alert('Successfully synced ' + username + ".");  
        
         Utilities.sleep(2000);
         userProperties.setProperty('CURRENT_DATA_ROW', 0);
         return true;
      }
      
      ++curr_data_row;
   }
  
   Logger.log("Made it to the end of rowData, resetting current_data_row for user.");
   
   Utilities.sleep(2000);
   userProperties.setProperty('CURRENT_DATA_ROW', 0);

   ui.alert('ODIN username not found.');
   lock.releaseLock();
   return false;  
};



// Copies the log sheet into the archive sheet.
function archiveLog()
{   
   var dataRange = LOG_SHEET.getRange(2, 1, LOG_SHEET.getMaxRows() - 1, 8);  

   // Create one JavaScript object per row of data.
   objects = getRowsData(LOG_SHEET, dataRange);
     
   var obj_remain = objects.length
   var i = 0;
   while(obj_remain > 1) 
   {
        // Get a row object from the spreadsheet
        var rowData = objects[i];
     
        try 
        {
           lock.waitLock(1000);
        }  
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        }
    
        ARCHIVE_SHEET.appendRow([rowData.timestamp, rowData.user, rowData.filename, rowData.currentOwner, rowData.previousOwner, rowData.lastModified, rowData.fileUrl, rowData.notes]);    
        LOG_SHEET.deleteRow(2);
        --obj_remain;
        ++i;
     
        if(obj_remain == 1)
        {
            var rowData = objects[i];
            ARCHIVE_SHEET.appendRow([rowData.timestamp, rowData.user, rowData.filename, rowData.currentOwner, rowData.previousOwner, rowData.lastModified, rowData.fileUrl, rowData.notes]);        
            --obj_remain;
        }
   }
  

   // Row 1 is locked and can't be deleted.
   // If there are only two rows, can't delete any rows, unless we add a blank row first.
   if(LOG_SHEET.getMaxRows() >= 2)
   {
      LOG_SHEET.insertRowBefore(2);
      LOG_SHEET.deleteRows(3, LOG_SHEET.getMaxRows() - 2);
   }

   return true;
};


// Transfers ownership of each users root folder to the user
// and removes the group's access.
function archiveFolders()
{
   try 
   {
      lock.waitLock(1000);
   }  
  
   catch(err) 
   {
       Logger.log('Could not obtain lock after 1 seconds.');
   }
    
   // Gather variables so we can keep track of the time, and where we are in the data.
   var root = getRoot();  
   var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
   var scriptProperties = PropertiesService.getScriptProperties();
   var curr_archive_row = parseInt(scriptProperties.getProperty("CURRENT_ARCHIVE_ROW")); 
   var startTime = new Date();
   var currentTime = new Date();
  
   // Create one JavaScript object per row of data.
   objects = getRowsData(DATA_SHEET, dataRange);
              
   // for (var i = 0; i < objects.length; ++i) 
   while(curr_archive_row < objects.length)
   { 
      currentTime = new Date();
     
      // Get a row object from the spreadsheet
      var rowData = objects[curr_archive_row];
      var item = findUserFolder(rowData);

      // Keep track of the time so we can exit before the script stops at 6 minutes.     
      if((currentTime.getTime() - startTime.getTime()) < 240000 && item != null)
      {
    
         if(rowData.status != "Failed")
         {
            Logger.log("User #: " + curr_archive_row);    
            Logger.log("First name: " + rowData.first);
            Logger.log("Last name: " + rowData.last);
            Logger.log("Odin: " + rowData.odin);
            Logger.log("Last Update: " + rowData.updated);
    
            var user_email = rowData.odin + "@pdx.edu";   
           
            removeEditorWalkDirectoryTree(item, rowData, rowData.files);

           
            if(item != null && CURRENT_USER === item.getOwner().getEmail() && item.getName() === (rowData.last + ", " + rowData.first + " " + "(" + rowData.odin + ")"))
            {
               // Gather information for the email we will send to the user
               // and for the log entry.
               var url = item.getUrl();
               var file_name = item.getName();
               var last_updated = item.getLastUpdated();
               var prev_owner = item.getOwner().getEmail();
               var curr_time = (new Date()).toString();
               var folder_name = item.getName();
               Logger.log("File name is: " + file_name);
             
               try 
               {
                   lock.waitLock(1000);
               } 
  
               catch(err) 
               {
                   Logger.log('Could not obtain lock after 1 seconds.');
               }
              
               try
               {
                   // Attempt to transfer ownership.
                   item.setOwner(user_email);
                   item.setName(folder_name + " - " + TERM_NAME)
                   var new_owner = item.getOwner().getEmail();
              
                   if(item.getAccess(GROUP_EMAIL) === DriveApp.Permission.EDIT)
                   {
                      item.removeEditor(GROUP_EMAIL);
                   }
                        
                   item.removeEditor(CURRENT_USER);
             
                   var note = "Successfully transferred ownership of root folder.";
                   Logger.log("Ownership transferred!");
                   updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);

                   currentTime = new Date();
                   DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Archived");
                   DATA_SHEET.getRange(UPDATED_COL + (curr_archive_row + 2)).setValue(currentTime.toString());

                   curr_archive_row = curr_archive_row + 1;
                   Utilities.sleep(2000);
                   scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', curr_archive_row); 
               }
           
               catch(err)
               {
                  // If we are not able to transfer ownership, send an email about it
                  // and record it in the log.
                  Logger.log("There was an error transfering root user folder ownership!");
        
                  var new_owner = item.getOwner().getEmail();
                  var note = "Failure in transferring ownership of root folder.";        
                  updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
                  sendErrorEmail(rowData);
                 
                  var currentTime = new Date();
                  DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Failure");
                  DATA_SHEET.getRange(UPDATED_COL + (curr_archive_row + 2)).setValue(currentTime.toString());

                  curr_archive_row = curr_archive_row + 1;
                  Utilities.sleep(1000);
                  scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', curr_archive_row);
               }
             }  
           }
        
           // The user had a failed state, we should not try to transfer ownership.
           // Send the error email about it, remove from database. Leave folder intact.
           if(rowData.status == "Failed")
           {
              sendErrorEmail(rowData);
              updateLog(user_email, "User root folder", "N/A", "N/A", "N/A", "N/A", "Cannot transfer this user has a failed status.");
             
              var currentTime = new Date();
              DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Failure");
              DATA_SHEET.getRange(UPDATED_COL + (curr_archive_row + 2)).setValue(currentTime.toString());
                                            
              curr_archive_row = curr_archive_row + 1;
              Utilities.sleep(2000);  
              scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', curr_archive_row);
           }
        }
     
        // The users folder was probably not found, move on. 
        if(item === null)
        {
            var currentTime = new Date();
            DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Archived");
            DATA_SHEET.getRange(UPDATED_COL + (curr_archive_row + 2)).setValue(currentTime.toString());
                            
           updateLog(rowData.odin + "@pdx.edu", "N/A", "N/A", "N/A", "N/A", "N/A", "User root folder did not exist.");
           curr_archive_row = curr_archive_row + 1;
           Utilities.sleep(2000);
           scriptProperties.setProperty('CURRENT_ARCHIVE_ROW', curr_archive_row);
        }
     
        if((currentTime.getTime() - startTime.getTime()) > 240000)
        {
           return false; 
        }
         
   }
  
   return true; 
};



// Clears the responses from the Form Responses 1 sheet, used during archiving.
function clearResponses()
{    
   try 
   {
      lock.waitLock(1000);
   } 
  
   catch(err) 
   {
      Logger.log('Could not obtain lock after 1 seconds.');
   } 
  
  
  
   // Row 1 is locked and can't be deleted.
   // If there are only two rows, can't delete any rows, unless we add a blank row first.
   if(FORM_SHEET.getMaxRows() >= 2)
   {
      FORM_SHEET.insertRowBefore(2);
      FORM_SHEET.deleteRow(3);
      FORM_SHEET.insertRowBefore(2);
      FORM_SHEET.deleteRows(3, FORM_SHEET.getMaxRows() - 2);
   }
   
   try 
   {
      lock.waitLock(1000);
   } 
  
   catch(err) 
   {
      Logger.log('Could not obtain lock after 1 seconds.');
   }
    
   // Row 1 is locked and can't be deleted.
   // If there are only two rows, can't delete any rows, unless we add a blank row first.
   if(DATA_SHEET.getMaxRows() >= 2)
   {
      DATA_SHEET.insertRowBefore(2);
      DATA_SHEET.deleteRows(3, DATA_SHEET.getMaxRows() - 2);
    
      // Rebuild our cell linking.
      DATA_SHEET.getRange(ODIN_COL + (2)).setValue("=query('Form Responses 1'!B2:Z, \"Select D\")");
      DATA_SHEET.getRange(FIRSTNAME_COL + (2)).setValue("=query('Form Responses 1'!B2:Z, \"Select B\")");
      DATA_SHEET.getRange(LASTNAME_COL + (2)).setValue("=query('Form Responses 1'!B2:Z, \"Select C\")");     
   }
  
   return true;
};


// Locates a trigger by the name of it's handler function.
function findTriggerByHandler(name)
{
    var allTriggers = ScriptApp.getProjectTriggers();
  
    // Loop over all triggers
    for (var i = 0; i < allTriggers.length; i++) 
    {
             
        if (allTriggers[i].getHandlerFunction() == name) 
        {
            // Found the trigger and now delete it
            return true;
        }
    } 
    return false;
  
};


// Removes a trigger by it's handler function name.
function removeTriggerByHandler(name)
{
    var allTriggers = ScriptApp.getProjectTriggers();
    
    // Loop over all triggers
    for (var i = 0; i < allTriggers.length; i++) 
    {
             
        if (allTriggers[i].getHandlerFunction() == name) 
        {
            // Found the trigger and now delete it
            ScriptApp.deleteTrigger(allTriggers[i]);
            return true;
        }
    }
  
    return false;
  
};


// The main archive function. Transfers ownership of each users root folder,
// archives the log sheet, clears the Form Responses 1 sheet.
function archive()
{
    // Stop any concurrent script running by each user.
    try 
    {
        lock.waitLock(1000);
    } 
  
    catch(err) 
    {
       Logger.log('Could not obtain lock after 1 seconds.');
    }
  
  
    // Check to make sure there isn't a public lock, from say archiving or
    // removing a user.
    if(pubLock.hasLock() == true)
    {    
      var ui = SpreadsheetApp.getUi();
      
      ui.alert("Document is locked for archiving or user removal. Please try again later.");
      return false;
    }
  
    else
    {
        try 
        {
           pubLock.waitLock(1000);
        } 
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        } 
    }
  
    var scriptProperties = PropertiesService.getScriptProperties();
  
    // This function will take several executions to finish,
    // we will need to use triggered events to keep executing it.
    // We will also need to use user properties to keep track of the 
    // current archiving data row.
    if(findTriggerByHandler("archive") == false)
    {
        ScriptApp.newTrigger("archive")
        .timeBased()
        .everyMinutes(5) // Frequency is required if you are using atHour() or nearMinute()
        .create();
        
        scriptProperties.setProperty("CURRENT_ARCHIVE_ROW", 0); 
    }
  
    var startTime = new Date();
    var currentTime = new Date();
    
    // Keep track of the time so we can exit before the script stops running at 6 minutes.
    while((currentTime.getTime() - startTime.getTime()) < 240000)
    {
         currentTime = new Date();
        
         if(archiveFolders() == true && archiveLog() == true && clearResponses() == true)
         {
            removeTriggerByHandler("archive");
            scriptProperties.setProperty("CURRENT_ARCHIVE_ROW", 0); 
            SpreadsheetApp.getUi().alert('Archiving finished.');
            lock.releaseLock();
            pubLock.releaseLock();
            return true;
         } 
      
         else
            return false;
    }
};



// Standalone function to clear a failed state for a user.
function clearFailed()
{
    // Stop any concurrent script running by each user.
    try 
    {
        lock.waitLock(1000);
    } 
  
    catch(err) 
    {
       Logger.log('Could not obtain lock after 1 seconds.');
    }
  
  
    // Check to make sure there isn't a public lock, from say archiving or
    // removing a user.
    if(pubLock.hasLock() == true)
    {    
      var ui = SpreadsheetApp.getUi();
      
      ui.alert("Document is locked for archiving or user removal. Please try again later.");
      return false;
    }
  
    else
    {
        try 
        {
           pubLock.waitLock(1000);
        } 
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        } 
    }
  
   // Setup a UI for a user to place input.
   var ui = SpreadsheetApp.getUi(); 

   var result = ui.prompt(
     'Enter ODIN username:',
      ui.ButtonSet.OK_CANCEL);

   // Process the user's response.
   var button = result.getSelectedButton();
   var username = result.getResponseText();
  
   if (button == ui.Button.OK && username != "") 
   {
      // User clicked "OK".
      ui.alert('Clearing failure for ' + username + '.');
   }
  
   else if (button == ui.Button.CANCEL) 
   {
      // User clicked "Cancel".
      ui.alert('Clear failure cancelled.');
      return false;
   } 
  
   else if (button == ui.Button.OK && username === "") 
   {
      // User didn't enter a username.
      ui.alert('Please enter a valid ODIN username.');
      return false;
   } 
  
   var root = getRoot();
   var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
 
   // Create one JavaScript object per row of data.
   objects = getRowsData(DATA_SHEET, dataRange);
   
   // Search for the user that we want to clear the failed status of.
   for (var i = 0; i < (objects.length - 1); ++i) 
   {
    
        // Get a row object from the spreadsheet
        var rowData = objects[i];
         
        // For testing.
        Logger.log("User #: " + i);    
        Logger.log("First name: " + rowData.first);
        Logger.log("Last name: " + rowData.last);
        Logger.log("Odin: " + rowData.odin);
        Logger.log("Last Update: " + rowData.updated);
     
        // Found the user remove the failure status.
        if(username.equals(rowData.odin))
        {
            DATA_SHEET.getRange(STATUS_COL + (i + 2)).setValue("");
            SpreadsheetApp.flush();
            ui.alert('Successfully cleared failure for ' + username + ".");
            return true;
        }    
  }
  
  ui.alert('ODIN username not found.');
  lock.releaseLock();
  pubLock.releaseLock();
  return false;
};


// Used for form submission, based on last row in Form Responses 1 sheet.
function createNewUser(e)
{ 
  Logger.log("Creating new user!");

  // Get the values that were submitted to the form.
  var itemResponses = e.response.getItemResponses();
  
  // Assign values that were submitted in the form.
  var new_user = {
      first: itemResponses[0].getResponse(),
      last: itemResponses[1].getResponse(),
      odin: itemResponses[2].getResponse()
   }
  
   Logger.log(new_user);
  
   // Create the new user's folder.
   if(setNewUser(null, new_user))
   {
      sendWelcomeEmail(new_user);
      return true;
   }
  
   else
      return false;

};


// Updates the log sheet with the information passed through the argument list.
function updateLog(user_email, file_name, new_owner, prev_owner, last_updated, file_url, note)
{ 
   // var toCopy = [];   
   var currentTime = (new Date()).toString();
   // toCopy.push([currentTime, user_email, file_name, new_owner, prev_owner, last_updated, file_url, note])
   
   try 
   {
       lock.waitLock(1000);
   } 
  
   catch(err) 
   {
        Logger.log('Could not obtain lock after 1 seconds.');
   }
   
   // **************************************************************************************
   // Issue: InsertRows/AppendRows seems to cause a script error every so often.
   // Symptom: When there are few or no content free rows and insert is used, the script 
   // appears to fail, but continues executing. A second copy (or identical thread?)
   // of the scripts appears to begin. All actions appear to be duplicated.
   // The original thread and copy seem to have their own script execution times, and will
   // execute until exceeding the time limit.
   // **************************************************************************************
  
   try
   {
        LOG_SHEET.appendRow([currentTime, user_email, file_name, new_owner, prev_owner, last_updated, file_url, note]);
        SpreadsheetApp.flush();
        return true;
   }
  
   catch(err)
   {
      Logger.log("There was an error writing to the log file.");
      return false;
   }
  
};


// Test function for sending emails.
function testSendEmail()
{
   var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
   objects = getRowsData(DATA_SHEET, dataRange);
  
   for (var i = 0; i < (objects.length - START_ROW) ; ++i) 
   {
     var rowData = objects[i];

     if(rowData.odin === "pdx01149")
     {
        sendWelcomeEmail(rowData);
     }    
   }
};


// Sends an email to any user who had their root user folder created.
function sendWelcomeEmail(rowData)
{
     try 
     {
           lock.waitLock(1000);
     } 
  
     catch(err) 
     {
        Logger.log('Could not obtain lock after 1 seconds.');
     }
  
  
     var user_email = rowData.odin + "@pdx.edu";
  
     // Pull template email from the template sheet.
     var emailTemplate = TEMPLATE_SHEET.getRange("A2").getValue();
  
     // Modify the emailTemplate with the folowing.
     var emailText = fillInTemplateFromObject(emailTemplate, rowData);
     var emailSubject = "Alternative Formats registration";
     var advmail=new Object();
     advmail.name = "DRC EMI";
     advmail.replyTo = GROUP_EMAIL;
     advmail.htmlBody = emailText;          
     MailApp.sendEmail(user_email, emailSubject, "", advmail);     
};


// Sends an email to any user who had files transfered to their ownership.
function sendFilesEmail(rowData)
{
     try 
     {
        lock.waitLock(1000);
     } 
  
     catch(err) 
     {
        Logger.log('Could not obtain lock after 1 seconds.');
     }
  
     var user_email = rowData.odin + "@pdx.edu";
  
     // Pull template email from the template sheet.
     var emailTemplate = TEMPLATE_SHEET.getRange("B2").getValue();
  
     // Modify the emailTemplate with the following.
     var emailText = fillInTemplateFromObject(emailTemplate, rowData);
     var emailSubject = "Alternative Formats file delivery";
     var advmail=new Object();
     advmail.name = "DRC EMI";
     advmail.replyTo = GROUP_EMAIL;
     advmail.htmlBody = emailText;          
     MailApp.sendEmail(user_email, emailSubject, "", advmail);     
  
};



// Sends an error email to the root owner and the user upon a failure of file ownership transfer.
function sendErrorEmail(rowData)
{
     try 
     {
        lock.waitLock(1000);
     } 
  
     catch(err) 
     {
        Logger.log('Could not obtain lock after 1 seconds.');
     }
  
     Logger.log("in senderroremail, rowdata is: " + rowData.odin);
     var user_email = rowData.odin + "@pdx.edu";
  
     // Pull template email from the template sheet.
     var emailTemplate = TEMPLATE_SHEET.getRange("C2").getValue();
  
     // Modify the emailTemplate with the following. 
     var emailText = fillInTemplateFromObject(emailTemplate, rowData);
     var emailSubject = "Error in file transfer";
     var advmail=new Object();
     
     advmail.name = "DRC EMI";
     advmail.replyTo = GROUP_EMAIL;
     advmail.cc = ROOT_OWNER;
     advmail.htmlBody = emailText;          
     MailApp.sendEmail(user_email, emailSubject, "", advmail);     
  
};



// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function fillInTemplateFromObject(template, data) {
  var email = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  for (var i = 0; i < templateVars.length; ++i) {
    // normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[normalizeHeader(templateVars[i])];
    email = email.replace(templateVars[i], variableData || "");
  }

  return email;
};



// Test function for walking a users directory structure. 
function testWalkDirectory() 
{
  // Testing 
  var username = "pdx01150";
  var dataRange = DATA_SHEET.getRange(2, 1, DATA_SHEET.getMaxRows() - 1, 5);  
  objects = getRowsData(DATA_SHEET, dataRange);
  
  for (var i = 0; i < (objects.length - START_ROW) ; ++i) 
  {
     var rowData = objects[i];
     rowData.files = [];
    
     if(rowData.odin === username)
     {
         var odin = findUserFolder(username);
         rowData.root = odin.getUrl();
         walkDirectoryTree(odin, rowData.odin, rowData.files);     
     
         Logger.log("Length: " + rowData.files.length);
       
         for(var y = 0; y < rowData.files.length; ++y)
         {
            Logger.log(y + " element: " + rowData.files[y]); 
         }
       
         var urlString = "";
         
         // Collect all of the file url's into a single
         // set of strings.
         for(var z = 0; z < rowData.files.length; ++z)
         {
            Logger.log(z);
       
            urlString = urlString + rowData.files[z];
         }
     }
     
     Logger.log(urlString);
    
     rowData.files = urlString;

     // If we have sent files out, let's send an email about it.
     if(rowData.files.length != 0)
     {
        sendFilesEmail(rowData);
     }
  }  
};



// Transfers file ownership to a user, and then sends an email with information about
// the transfered files.
function xferFiles(user_folder, rowData) 
{
   
     rowData.files = [];
    
     // Keep track of location of the users root folder.
     rowData.root = user_folder.getUrl();
  
     // Now search for files and folders that the staff member
     // owns, and transfer those to the student.
     walkDirectoryTree(user_folder, rowData, rowData.files);     
     
     // For testing.
     Logger.log("Length: " + rowData.files.length);
       
     for(var y = 0; y < rowData.files.length; ++y)
     {
         Logger.log(y + " element: " + rowData.files[y]); 
     }
       
     var urlString = "";
     
     // Pack all of our file url's into one string.
     for(var z = 0; z < rowData.files.length; ++z)
     {
         Logger.log(z);
         urlString = urlString + rowData.files[z];
     }
     
     // Store our string in the data structure.
     rowData.files = urlString;
  
     // If we have sent files out, let's send an email about it.
     if(rowData.files.length != 0)
     {
        sendFilesEmail(rowData);
     }
};



// Function specifically to walk through a users directories
// and remove the staff group editor permissions.
function removeEditorWalkDirectoryTree(folder, rowData, files)
{
   Logger.log("Inside of removeEditorWalkDirectory");
  
   var scriptProperties = PropertiesService.getScriptProperties();         
   Utilities.sleep(1000);
   var curr_archive_row = parseInt(scriptProperties.getProperty('CURRENT_ARCHIVE_ROW'));
     
   Logger.log("curr_archive_row is: " + curr_archive_row); 
  
   Logger.log("Files in " + folder.getName());
   
   removeEditorFiles(folder.getFiles(), rowData, files);
   
   // We need to go to every directory so we can't use searchFolders here.
   var subfolders = folder.getFolders();
   
   // Now search each folder.
   while(subfolders.hasNext() && rowData.status != "Failed")
   {
      var subfolder = subfolders.next();
      removeEditorWalkDirectoryTree(subfolder, rowData, files);
   } 
  
   removeEditorFolders(folder, rowData, files);
};



// Removes the edit permissions for the staff group from
// a users folders.
function removeEditorFolders(item, rowData, files)
{
   Logger.log("Inside of removeEditorFolders");
   var user_email = rowData.odin + "@pdx.edu";
  
   // If we edit the folder then we probably want to remove that.
   if(item != null && item.getAccess(GROUP_EMAIL) === DriveApp.Permission.EDIT)
   {
     
      // Protect against the root owner transfering ownership of the users root folder.
      if(item.getName() != rowData.last + ", " + rowData.first + " " + "(" + rowData.odin + ")" )
      {  
       
         // Collect information for our email to the student and our log entry.
         var file_name = item.getName();
         var url = item.getUrl();
         var last_updated = item.getLastUpdated();
         var prev_owner = item.getOwner().getEmail();
        
         var scriptProperties = PropertiesService.getScriptProperties();         
         Utilities.sleep(1000);
         var curr_archive_row = parseInt(scriptProperties.getProperty('CURRENT_ARCHIVE_ROW'));
        
         try 
         {
            lock.waitLock(1000);
         } 
  
         catch(err) 
         {
            Logger.log('Could not obtain lock after 1 seconds.');
         }
             
         try
         { 
            // Try to transfer ownership, if it works update the log and our position in the data structure.
            var new_owner = item.getOwner().getEmail();
            item.removeEditor(GROUP_EMAIL);      
            var note = "Successfully removed group edit access.";
            updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
          
            DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Success");
            SpreadsheetApp.flush();
         }
        
         catch(err)
         {
            // If there was an error transferring ownership, then set the users status
            // to failed, record the failure in the log, and send an email to the student
            // and administrator.
            Logger.log("Failure to remove group edit access.");
           
            DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Failed");
            SpreadsheetApp.flush();
            rowData.status = "Failed";
  
            var new_owner = item.getOwner().getEmail();
            var note = "Failure to remove group edit access.";        
            updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
            sendErrorEmail(rowData);
       }       
     }
   }  
};



// Removes the edit permissions for the staff group from
// a users files.
function removeEditorFiles(items, rowData, files)
{
   Logger.log("Inside of removeEditorFiles");
  
   var user_email = rowData.odin + "@pdx.edu";
  
   // Move through the files in the directy and transfer ownership
   // if the user doesn't have a failed status.
   while(items.hasNext() && rowData.status != "Failed")
   {
      var item = items.next();
      Logger.log(item.getName());

      Logger.log("Adding file: " + item.getName() + " to list!");
      
      // Collect data for our email to the user and the log entry.
      var file_name = item.getName();
      // Create the direct download link.
      var url = "https://docs.google.com/uc?export=download&id=" + item.getId();
      var this_url = "Link: " + url + "<br><br>";
      var file_msg = file_name + "<br>" + this_url;
      
      var last_updated = item.getLastUpdated();
      var prev_owner = item.getOwner().getEmail();
      
      var scriptProperties = PropertiesService.getScriptProperties();         
      Utilities.sleep(1000);
      var curr_archive_row = parseInt(scriptProperties.getProperty('CURRENT_ARCHIVE_ROW'));
     
      Logger.log("curr_archive_row is: " + curr_archive_row); 
      
      try 
      {
           lock.waitLock(1000);
      } 
  
      catch(err) 
      {
           Logger.log('Could not obtain lock after 1 seconds.');
      }
     
      try
      {  
         // Try to transfer ownership to the user, if it works update the log
         var new_owner = item.getOwner().getEmail();
         item.removeEditor(GROUP_EMAIL);
         var note = "Successfully removed group edit access.";
         updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
                
         DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Success");
         SpreadsheetApp.flush(); 

      }
      
      catch(err)
      {
         // If ownership transfer failed, then set the user's status as failed,
         // update the log, and send an email to the student and administrator.
         Logger.log("There was an error transfering file ownership!");
        
         DATA_SHEET.getRange(STATUS_COL + (curr_archive_row + 2)).setValue("Failed");
         SpreadsheetApp.flush();
         rowData.status = "Failed";
        
         var new_owner = item.getOwner();
         var note = "Failed to remove group edit access.";        
         updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
         sendErrorEmail(rowData);
      } 
   }
};



// Walks a users folder directory structure.
function walkDirectoryTree(folder, rowData, files)
{
   Logger.log("Files in " + folder.getName());

   // Serch files in a folder first.
   var owner_search = '"' + CURRENT_USER + '" in owners';
   walkDirectoryFiles(folder.searchFiles(owner_search), rowData, files);
   
   // We need to go to every directory so we can't use searchFolders here.
   var subfolders = folder.getFolders();
   
   // Now search each folder.
   while(subfolders.hasNext() && rowData.status != "Failed")
   {
      var subfolder = subfolders.next();
      walkDirectoryTree(subfolder, rowData, files);
   } 
  
   walkDirectoryFolders(folder, rowData, files);
 
};



// Transfers ownership of a folder to the user
// Removes access from the current owner.
// Group edit access is maintained.
function walkDirectoryFolders(item, rowData, files)
{   
  var user_email = rowData.odin + "@pdx.edu";
  
  // If we own the folder then we probably want to transfer it to the student.
  if(item.getOwner().getEmail() === CURRENT_USER)
  {
     
     // Protect against the root owner transfering ownership of the users root folder.
     if(item.getName() != rowData.last + ", " + rowData.first + " " + "(" + rowData.odin + ")" )
     {  
       
        // Collect information for our email to the student and our log entry.
        var file_name = item.getName();
        var url = item.getUrl();
        var last_updated = item.getLastUpdated();
        var prev_owner = item.getOwner().getEmail();
       
        try 
        {
           lock.waitLock(1000);
        } 
  
        catch(err) 
        {
           Logger.log('Could not obtain lock after 1 seconds.');
        }
             
        try
        {
           // Try to transfer ownership, if it works update the log and our position in the data structure.
           item.setOwner(user_email);
           item.removeEditor(CURRENT_USER);      
           var new_owner = item.getOwner().getEmail();
           var note = "Success";
           updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
        
           var userProperties = PropertiesService.getUserProperties();         
           Utilities.sleep(1000);
           var current_row = parseInt(userProperties.getProperty('CURRENT_DATA_ROW'));
          
           DATA_SHEET.getRange(STATUS_COL + (current_row + 2)).setValue("Success");
           SpreadsheetApp.flush();
        }
        
        catch(err)
        {
           // If there was an error transferring ownership, then set the users status
           // to failed, record the failure in the log, and send an email to the student
           // and administrator.
           Logger.log("There was an error transfering file ownership!");

           var userProperties = PropertiesService.getUserProperties();
           Utilities.sleep(1000);
           var current_row = parseInt(userProperties.getProperty('CURRENT_DATA_ROW'));
           
           DATA_SHEET.getRange(STATUS_COL + (current_row + 2)).setValue("Failed");
           SpreadsheetApp.flush();
           rowData.status = "Failed";
  
           var new_owner = item.getOwner();
           var note = "Failure";        
           updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
           sendErrorEmail(rowData);
       }       
     }
   }  
};



// Transfers ownership of a files to the user
// Removes access from the current owner.
// Group edit access is maintained.
function walkDirectoryFiles(items, rowData, files)
{
   
   var user_email = rowData.odin + "@pdx.edu";
  
   // Move through the files in the directy and transfer ownership
   // if the user doesn't have a failed status.
   while(items.hasNext() && rowData.status != "Failed")
   {
      var item = items.next();
      Logger.log(item.getName());

      Logger.log("Adding file: " + item.getName() + " to list!");
      
      // Collect data for our email to the user and the log entry.
      var file_name = item.getName();
      // Create the direct download link.
      var url = "https://docs.google.com/uc?export=download&id=" + item.getId();
      var this_url = "Link: " + url + "<br><br>";
      var file_msg = file_name + "<br>" + this_url;
      
         
      files[files.length] = file_msg;
      
      var last_updated = item.getLastUpdated();
      var prev_owner = item.getOwner().getEmail();
      
      try 
      {
           lock.waitLock(1000);
      } 
  
      catch(err) 
      {
           Logger.log('Could not obtain lock after 1 seconds.');
      }
     
      try
      {  
         // Try to transfer ownership to the user, if it works update the log
         item.setOwner(user_email);
         item.removeEditor(CURRENT_USER);
         var new_owner = item.getOwner().getEmail();
         var note = "Success";
         updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
        
         var userProperties = PropertiesService.getUserProperties();
         Utilities.sleep(1000);
         var current_row = parseInt(userProperties.getProperty('CURRENT_DATA_ROW'));
        
         Logger.log("Updating status on row " + (current_row));
         DATA_SHEET.getRange(STATUS_COL + (current_row + 2)).setValue("Success");
         SpreadsheetApp.flush(); 

      }
      
      catch(err)
      {
         // If ownership transfer failed, then set the user's status as failed,
         // update the log, and send an email to the student and administrator.
         Logger.log("There was an error transfering file ownership!");

         var userProperties = PropertiesService.getUserProperties();
         Utilities.sleep(1000);
         var current_row = parseInt(userProperties.getProperty('CURRENT_DATA_ROW'));
        
         DATA_SHEET.getRange(STATUS_COL + (current_row + 2)).setValue("Failed");
         SpreadsheetApp.flush();
         rowData.status = "Failed";
        
         var new_owner = item.getOwner();
         var note = "Failure";        
         updateLog(user_email, item.getName(), new_owner, prev_owner, last_updated, url, note);
         sendErrorEmail(rowData);
      }
     
   }
  
};



// Finds and returns the user's root foldero bject.
function findUserFolder(rowData)
{
   var root = getRoot();
   var match_folder = root.getFoldersByName(rowData.last + ", " + rowData.first + " " + "(" + rowData.odin + ")");
  
   Logger.log("Searching for folder: " + rowData.odin);    
  
   while(match_folder.hasNext())
   {
      var to_find = match_folder.next();
      var name = to_find.getName();
       
      Logger.log("Found folder!");                          
      return to_find;      
   }
  
   // else - We didn't find a user folder with that odin username!
   Logger.log("User folder not found!");
   return null;
};



// Creates a new user folder and sets the initial permissions.
function setNewUser(user_folder, rowData)
{
   var user_email = rowData.odin + "@pdx.edu";

   // The user folder already exists, make sure the correct permissions are set.
   if(user_folder != null && CURRENT_USER === ROOT_OWNER)
   {
       try 
       {
           lock.waitLock(1000);
       } 
  
       catch(err) 
       {
           Logger.log('Could not obtain lock after 1 seconds.');
       }
     
     
       try
       {
          Logger.log("Setting initial permissions for user folder!");
          user_folder.addEditor(GROUP_EMAIL);
          user_folder.addViewer(user_email);
          return true;
       }
  
       catch(err)
       {
          Logger.log("Unable to set initial permission for user folder!");
          return false;
       }
   }
  
   // The folder doesn't exist, we are the root owner, let's create the users root folder
   // and set the permissions. Then update the log.
   else if(user_folder == null && CURRENT_USER === ROOT_OWNER)
   {
      try 
      {
         lock.waitLock(1000);
      } 
  
      catch(err) 
      {
         Logger.log('Could not obtain lock after 1 seconds.');
      }
     
      try
      {
         Logger.log("Creating user folder!");
         var root = getRoot();
         var user_folder = root.createFolder(rowData.last + ", " + rowData.first + " " + "(" + rowData.odin + ")");
         Logger.log("Setting initial permissions for user folder!");
         user_folder.addEditor(GROUP_EMAIL);
         user_folder.addViewer(user_email);
         var new_owner = user_folder.getOwner().getEmail();
         var url = user_folder.getUrl();
         updateLog(user_email, "Root user folder", new_owner, "n/a", "n/a", url, "User folder created");
         return true;
      }
     
      catch(err)
      {
        Logger.log("Unable to create users root folder or share it.");
        return false;
      }
   }
  
   return false;
};



// Finds and returns the root folder object.
function getRoot()
{
   var root_folders = DriveApp.getFoldersByName(ROOT_FOLDER);
   
   // Find the root folder.
   while(root_folders.hasNext())
   {
      var root = root_folders.next();
        
      if(root.getName().equals(ROOT_FOLDER))
      {
          return root;  
      }
       
    }
  
    // If we couldn't find the root folder and we are the administrator, create the root folder.
    if(CURRENT_USER === ROOT_OWNER)
    {
        root = DriveApp.CreateFolder(ROOT_FOLDER);  
        root.addEditor(GROUP_EMAIL);
        return root;
    }
    
    // else - we couldn't find root and we don't have permission to create it, return null!
    return null;
};



/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('DRC EMI menu')
      .addItem('Transfer files to one user', 'syncUserWrapper')
      .addItem('Full file transfer', 'manualSync')
      .addItem('Clear user fialure status', 'clearFailed')
      .addItem('Group member trigger authorization', 'memberWrapper')
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Admin options')
      .addItem('Remove user', 'removeUserWrapper')
      .addItem('Admin trigger authorization', 'adminWrapper')
      .addItem('Archive', 'archiveWrapper'))
      .addToUi();
};



// Wrapper function for sync'ing one users files.
function syncUserWrapper() {

  syncUser();
};



// Wrapper function for removing a user from the database.
function removeUserWrapper() {

  if(CURRENT_USER === ROOT_OWNER)
  {
     removeUser();
  }
  
  else
  {
     SpreadsheetApp.getUi().alert('You do not have permission to do that.');
  }
};



// Wrapper function for archiving the database.
function archiveWrapper() {
  
  if(CURRENT_USER === ROOT_OWNER)
  {
     SpreadsheetApp.getUi().alert('Archiving started.');
     archive()
     return true;
  }
  
  else
  {
     SpreadsheetApp.getUi().alert('You do not have permission to do that.');
     return false;
  }
};



// Allows for manual exection of the main function.
function manualSync() {
  var userProperties = PropertiesService.getUserProperties();
  Utilities.sleep(1000);
  userProperties.setProperty('CURRENT_DATA_ROW', 0);
  
  SpreadsheetApp.getUi().alert('Manual file transfer started.');
  main();
};



// Wrapper function for triggers.
function memberWrapper()
{
   memberTriggers(); 
};



// Group members must authorize the time based trigger.
function memberTriggers() {
 // Runs at 2am in the timezone of the script
 ScriptApp.newTrigger("main")
   .timeBased()
   .atHour(2)
   .everyDays(1) 
   .create();
  
  newGroupMember();
  
  SpreadsheetApp.getUi().alert('Script auto-run trigger authorized');

};



// A wrapper function for later use.
// Currently only executes adminTriggers function.
function adminWrapper()
{
   adminTriggers(); 
};



// Admin must authorize the on form submit trigger.
function adminTriggers() {

  if(ROOT_OWNER === CURRENT_USER)
  {
      // Runs on form submission.
      ScriptApp.newTrigger('createNewUser')
        .forForm(FORM_ID)
        .onFormSubmit()
        .create();
         
      // Runs at 2am in the timezone of the script
      ScriptApp.newTrigger("main")
        .timeBased()
        .atHour(2)
        .everyDays(1) 
        .create();
    
      newGroupMember();   
      SpreadsheetApp.getUi().alert('Form submission trigger authorized');
  }
  
  else
     SpreadsheetApp.getUi().alert('You are not authorized to perform this operation.');  
};



// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
};



// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
};



// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
};



// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
};



// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
};



// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
};



// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
};


/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
};