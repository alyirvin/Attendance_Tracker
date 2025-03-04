/**
 * Member Attendance Tracking System
 * 
 * Description:
 * This system is designed for tracking member attendance and points within a Google Drive folder.
 * The system includes a "Control Panel" Google Sheet where users can:
 * - Generate different types of sign-in forms (Standard, Fundraising, and Photo Evidence Events).
 * - Update the overall point values shown in the "Member Points" Google Sheet.
 * - Modify member names and emails.
 * - Delete members from all events.
 * - Look up individual member attendance records.
 * 
 * The system is organized by time period grouping (semesters, quarters, etc.), and each grouping 
 * is further divided into months for easy tracking.
 * 
 * Author: Alysha Irvin
 * Last Updated: March 1, 2025
 * 
 * Original Prototype Developed By: August Druzgal (github.com/august-druzgal)
 * On: January 7, 2024
 */

// Control Panel Constant
const SHEET = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

// Semester Constant in S3
const semester = SHEET.getRange(3, 19).getCell(1, 1).getValue();

// Folder Constants
const MEMBER_FOLDER = DriveApp.getFoldersByName("Member Points").next();
const MEMBER_FOLDER_ID = MEMBER_FOLDER.getId();

const SEMESTER_FOLDER = MEMBER_FOLDER.getFoldersByName(semester).next();
const SEMESTER_FOLDER_ID = SEMESTER_FOLDER.getId();

const SHEETS_FOLDER = SEMESTER_FOLDER.getFoldersByName('Sheets').next();
const SHEETS_FOLDER_ID = SHEETS_FOLDER.getId();

const FORMS_FOLDER = SEMESTER_FOLDER.getFoldersByName('Forms').next();
const FORMS_FOLDER_ID = FORMS_FOLDER.getId();

const TEMPLATES_FOLDER = MEMBER_FOLDER.getFoldersByName('*Templates').next();
const TEMPLATES_FOLDER_ID = TEMPLATES_FOLDER.getId();

// Template Name Constants
const FORM_TEMPLATE_NAME = "Template Sign-In";
const SHEET_TEMPLATE_NAME = "Template Responses";
const FUND_FORM_TEMPLATE_NAME = "Fund Template Sign-In";
const FUND_SHEET_TEMPLATE_NAME = "Fund Template Responses";
const PHOTO_EV_SHEET_TEMPLATE_NAME = "Photo Ev Template Responses";

const MEMBER_POINT_SHEET_NAME = "Member Point Sheet";


// Styling Constants
const DEFAULT_MEMBER_COLOR = "#cfe2f3";
const ACTIVE_MEMBER_COLOR = "#b6d7a8";
const INVOLVED_MEMBER_COLOR = "#8e7cc3";
const MOST_INVOLVED_MEMBER_COLOR = "#f1c232";

// Offset for the sheet headers
const MEMBER_POINT_SHEET_OFFSET = 3;


// Creates a new sign in form based on the inputs from the Control Panel Form
function createNewForm()
{
    // Initializes the values from the Control Panel's form
    let formInputs = SHEET.getRange(4, 3, 8, 1).getValues();
    let eventName = formInputs[0][0];
    let errorCell = SHEET.getRange(3, 2);
    let eventType = formInputs[2][0];
    let customQuestion = formInputs[4][0];
    let isRequired = formInputs[5][0];
    let eventMonth = formInputs[7][0];

    // Sets if the custom question is required
    if (isRequired == "")
    {
        isRequired = false;
    }

    // Makes sure that there is an event name entered
    if (!eventName)
    {
        errorCell.setValue("Please fill in the event name");
        return;
    }
    
    // Verifies that the event name is a String, and saves it if it is
    if (typeof eventName !== "string")
    {
        errorCell.setValue("Please enter a valid event name");
        return;
    }

    // Checks if a custom question was requested
    if (!customQuestion)
    {
        // Checks if the custom question is a string, and saves it if it is
        if (typeof customQuestion !== "string")
        {
        errorCell.setValue("Please enter a valid question");
        return;
        }
    }

    // Get all of the templates from the Templates folder
    let templateFolder = DriveApp.getFolderById(TEMPLATES_FOLDER_ID);
    let templateSheet, templateForm;

    let templateFiles = templateFolder.getFiles(); 
    let fileMap = {};
    while (templateFiles.hasNext()) {
        let file = templateFiles.next();
        fileMap[file.getName()] = file;
    }

    // Determines which template to use based on the event type
    if (eventType == "Standard")
    {
        templateSheet = fileMap[SHEET_TEMPLATE_NAME];
        templateForm = fileMap[FORM_TEMPLATE_NAME];
    }
    else if (eventType == "Fundraising")
    {
        templateSheet = fileMap[FUND_SHEET_TEMPLATE_NAME];
        templateForm = fileMap[FUND_FORM_TEMPLATE_NAME];
    }
    else if (eventType == "Photo Evidence")
    {
        templateSheet = fileMap[PHOTO_EV_SHEET_TEMPLATE_NAME];
        templateForm = fileMap[FORM_TEMPLATE_NAME];
    }

    // Creates a copy of the Sheet Template and opens it so that sheet-specific information can be set
    let newSheetFile = templateSheet.makeCopy();
    let newSheetContents = SpreadsheetApp.openById(newSheetFile.getId());

    // Gets all of the folders within the Sheets folder
    let sheetsFolder = DriveApp.getFolderById(SHEETS_FOLDER_ID);
    let foldersSheets = sheetsFolder.getFoldersByName(eventMonth);

    // Checks if there is already a subfolder for the month, and creates one if it does not exist
    let sheetsSubfolder;
    if (foldersSheets.hasNext())
    {
        sheetsSubfolder = foldersSheets.next();
    }
    else
    {
        sheetsSubfolder = sheetsFolder.createFolder(eventMonth);
    }

    // Moves the new sheet to the correct month subfolder and names it after the event
    newSheetFile.moveTo(sheetsSubfolder);
    newSheetFile.setName(eventName + " Member Sheet");

    // Creates a copy of the Form Template and opens it so that form-specific information can be set
    let newFormFile = templateForm.makeCopy();
    let newFormContents = FormApp.openById(newFormFile.getId());

    // Gets all of the folders within the Forms folder
    let formsFolder = DriveApp.getFolderById(FORMS_FOLDER_ID);
    let foldersForms = formsFolder.getFoldersByName(eventMonth);

    // Checks if there is already a subfolder for the month, and creates one if it does not exist
    let formsSubfolder;
    if (foldersForms.hasNext()) {
        formsSubfolder = foldersForms.next();
    } 
    else 
    {
        formsSubfolder = formsFolder.createFolder(eventMonth);
    }

    // Moves the new form to the correct month subfolder and names it after the event
    newFormFile.moveTo(formsSubfolder);
    newFormFile.setName(eventName + " Sign-In Form");
    newFormContents.setTitle(eventName + " Sign-In Form");

    // Links the responses to its newly created Sheet
    newFormContents.setDestination(FormApp.DestinationType.SPREADSHEET, newSheetFile.getId());

    // If a custom question needs to be added
    if (customQuestion)
    {
        // Creates a new question and indicates if it needs to be required
        let addQuestion = newFormContents.addTextItem();
        addQuestion.setTitle(customQuestion);
        addQuestion.setRequired(isRequired);
    }

    // Connect the response sheet to the form responses sheet
    let responsesSheet = newSheetContents.getSheetByName("Responses");
    newSheetContents.setActiveSheet(responsesSheet);

    // Resets all of the values in the Control Panel Form
    SHEET.getRange(4, 3, 6, 1).setValues([[""], [""], ["Standard"], [""], [""], [""]]);
    SHEET.getRange(3, 2).setValue("");

    // Logs the creation of the new form and sheet to the console for tracking purposes
    Logger.log("Created " + eventName + " Form and Sheet of Event Type " + eventType);
}

// Updates the Member Points Sheet with the most recent attendance Data
function updateMemberPoints()
{
    // Gets the Member Points Sheet from the Semester's Folder
    let semesterFolder = DriveApp.getFolderById(SEMESTER_FOLDER_ID);
    let memberFile = semesterFolder.getFilesByName(MEMBER_POINT_SHEET_NAME).next();
    let memberSheet = SpreadsheetApp.open(memberFile).getSheetByName("Member Points");

    // Resets the information currently in the sheet
    memberSheet.getRange("A" + MEMBER_POINT_SHEET_OFFSET + ":B1000").setBackgroundRGB(255, 255, 255).clearContent();

    // Gets all of the files in the Sheets folder
    let sheets = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFilesByType("application/vnd.google-apps.spreadsheet");
    let responseSet = null;
    let ids = [];

    // Get all of the IDs of the files in the Sheets folder and its subfolders
    while (sheets.hasNext())
        ids.push(sheets.next().getId());

    let subfolders = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFolders();
    while (subfolders.hasNext())
    {
        sheets = subfolders.next().getFilesByType("application/vnd.google-apps.spreadsheet");
        while(sheets.hasNext())
        ids.push(sheets.next().getId());
    }

    // Iterate through the list of IDs and get the responses from each event
    while (ids.length > 0)
    {
        let sheet = SpreadsheetApp.openById(ids.pop());
        info = getResponsesFromEvent(sheet.getSheetByName("Responses"));
        
        // If responseSet has not been set, set it to info
        if (responseSet == null)
        {
        responseSet = info;
        continue;
        }

        // Iterate through the responses and add the name and points to the responseSet
        for (const property in info.names)
        {
        if (responseSet.names[property] == undefined)
        {
            responseSet.names[property] = info.names[property];
            responseSet.points[property] = info.points[property];
        }
        else
        {
            responseSet.points[property] = responseSet.points[property] + info.points[property];
        }
        }
    }

    // Iterate through all of the responses and get a list of the emails
    let emails = [];
    for (const property in responseSet.names)
    {
        emails.push(property);
    }

    let max = 0;
    let indexes = [];

    // Iterate through the emails and set the name, points, and email in the Member Points Sheet
    for (let i = 0; i < emails.length; i++)
    {
        let total = responseSet.points[emails[i].toString()];
        let color;
        
        memberSheet.getRange("A" + (i + MEMBER_POINT_SHEET_OFFSET)).getCell(1, 1).setValue(responseSet.names[emails[i].toString()]);
        memberSheet.getRange("B" + (i + MEMBER_POINT_SHEET_OFFSET)).getCell(1, 1).setValue(total);
        memberSheet.getRange("C" + (i + MEMBER_POINT_SHEET_OFFSET)).getCell(1, 1).setValue(emails[i].toString());

        // If a user has less than 3 points, set them to the default styling
        if (total < 3)
        color = DEFAULT_MEMBER_COLOR;
        // Otherwise, if they have less than 15 points, set them to active member styling
        else if (total < 15)
        color = ACTIVE_MEMBER_COLOR;
        // Otherwise, set them to involved member styling
        else
        color = INVOLVED_MEMBER_COLOR;
        
        // Apply styling to the row
        memberSheet.getRange("A" + (i + MEMBER_POINT_SHEET_OFFSET) + ":B" + (i + MEMBER_POINT_SHEET_OFFSET)).setBackground(color);

        // If the user has more points than the current max
        if (total > max)
        {
        // Refresh the array of top point earners and set the new person as a top point earner and their total to the max
        indexes = [];
        indexes.push(i);
        max = total;
        }
        // If the user has the same number of points as the max, add them to the list of top point earners
        else if (total == max)
        {
        indexes.push(i);
        }
    }

    // If the max is greater than or equal to 15
    if (max >= 15)
    {
        // Style all of the top point earners with the most involved member styling
        for (let i = 0; i < indexes.length; i++)
        {
        memberSheet.getRange("A" + (indexes[i] + MEMBER_POINT_SHEET_OFFSET) + ":B" + (indexes[i] + MEMBER_POINT_SHEET_OFFSET)).setBackground(MOST_INVOLVED_MEMBER_COLOR);
        }
    }

    // Sort the Member Points Sheet by name
    memberSheet.sort(1);

    // Hide the column containing member emails
    memberSheet.hideColumn(memberSheet.getRange("C1"));

    // Indicate on the sheet when the Member Points Sheet was updated last
    let currentDate = new Date();
    let lastUpdated = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MM/dd 'at' hh:mm a");
    memberSheet.getRange("A1:A1").setValue("Last Updated: " + lastUpdated);
}

// Gets all of the responses from an event sheet
function getResponsesFromEvent(sheet)
{
    let names = {};
    let points = {};
    let cont = true;
    let index = 2;
    let name;
    let email;

    // Gets the default points someone gets for attending the event from the sheet
    let defaultPoints = sheet.getRange("F2").getValue();
    let additionalPoints;

    // Iterate through all of the responses
    while (cont)
    {
        // Get the name, email, and any additional points they were awarded
        cont = false;
        name = sheet.getRange("A" + index).getValue();
        email = sheet.getRange("B" + index).getValue();
        email = email.toLowerCase();
        additionalPoints = sheet.getRange("D" + index).getValue();

        // If the email is not empty, save the name and total points
        if (email != "")
        {
        names[email] = name;
        points[email] = defaultPoints + (additionalPoints != "" ? additionalPoints : 0);
        cont = true;
        index++;
        }
    }

    return {names, points};
}

// Constant locations for the fixEmail form in the Control Panel
const OLD_EMAIL_CELL = "K4";
const NEW_EMAIL_CELL = "K6";
const EMAIL_ERROR_CELL = "J3";

// Corrects an incorrect email to a new one
function fixEmail()
{
    // Get the old email and the new one from the form
    let oldEmail = SHEET.getRange(OLD_EMAIL_CELL).getCell(1, 1).getValue().toLowerCase();
    let newEmail = SHEET.getRange(NEW_EMAIL_CELL).getCell(1, 1).getValue().toLowerCase();

    // If either of the email fields are empty, display an error message and stop
    if (oldEmail == "" || newEmail == "")
    {
        SHEET.getRange(EMAIL_ERROR_CELL).getCell(1, 1).setValue("One of the fields is empty").setFontColor("red");
        return;
    }
    
    // Get all of the files in the Sheets folder
    let sheets = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFilesByType("application/vnd.google-apps.spreadsheet");

    // Iterate through all of the sheets and replace the old email with the new one
    while (sheets.hasNext())
    {
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        replaceEmailInSheet(newEmail, oldEmail, sheet);
    }

    // Get all of the subfolders in the Sheets folder
    let subfolders = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFolders();
    
    // Get all of the sheets in the subfolders
    while (subfolders.hasNext())
    {
        sheets = subfolders.next().getFilesByType("application/vnd.google-apps.spreadsheet");

        // Rpleace the old email with the new one in each sheet
        while (sheets.hasNext())
        {
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        replaceEmailInSheet(newEmail, oldEmail, sheet);
        }
    }

    // Update the Member Points Sheet and reset the email fields in the form
    updateMemberPoints();
    SHEET.getRange(OLD_EMAIL_CELL).setValue("");
    SHEET.getRange(NEW_EMAIL_CELL).setValue("");

    // Log the change in the console for tracking purposes
    Logger.log("Fixed " + oldEmail + " to " + newEmail);
}

// Replaces all instances of an email in a sheet with another
function replaceEmailInSheet(newEmail, oldEmail, sheet)
{
    let index = 2;
    let cont = true;
    let email;

    // While there are still responses
    while (cont)
    {
        // Get the email in the row
        cont = false;
        email = sheet.getSheetByName("Form Responses 1").getRange("C" + index).getCell(1, 1).getValue().toLowerCase();

        // If the email is not empty, keep looking, otherwise stop
        if (email != "")
        cont = true;
        else
        break;
        
        // If the email matches the old email, replace it with the new one
        if (email == oldEmail)
        {
        sheet.getSheetByName("Form Responses 1").getRange("C" + index).getCell(1, 1).setValue(newEmail);
        }

        index++;
    }

    index = 2;
    cont = true;
    
    // Repeat the same process for the Responses sheet
    while (cont)
    {
        cont = false;
        email = sheet.getSheetByName("Responses").getRange("B" + index).getCell(1, 1).getValue().toLowerCase();
        if (email != "")
        cont = true;
        else
        break;
        
        if (email == oldEmail)
        {
        sheet.getSheetByName("Responses").getRange("B" + index).getCell(1, 1).setValue(newEmail);
        }

        index++;
    }
}

// Constant locations for the fixName form in the Control Panel
const OLD_NAME_CELL = "K13";
const NEW_NAME_CELL = "K15";
const NAME_ERROR_CELL = "J12";

// Corrects an incorrect name to a new one
function fixName()
{
    // Get the old name and the new one from the form
    let oldName = SHEET.getRange(OLD_NAME_CELL).getCell(1, 1).getValue();
    let newName = SHEET.getRange(NEW_NAME_CELL).getCell(1, 1).getValue();

    // If either of the name fields are empty, display an error message and stop
    if (oldName == "" || newName == "")
    {
        SHEET.getRange(NAME_ERROR_CELL).getCell(1, 1).setValue("One of the fields is empty").setFontColor("red");
        return;
    }
    
    // Get all of the files in the Sheets folder
    let sheets = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFilesByType("application/vnd.google-apps.spreadsheet");

    // Iterate through all of the sheets and replace the old name with the new one
    while (sheets.hasNext())
    {
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        replaceNameInSheet(newName, oldName, sheet);
    }

    // Get all of the subfolders in the Sheets folder
    let subfolders = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFolders();
    

    // Get all of the sheets in the subfolders
    while (subfolders.hasNext())
    {
        sheets = subfolders.next().getFilesByType("application/vnd.google-apps.spreadsheet");

        // Replace the old name with the new one in each sheet
        while (sheets.hasNext())
        {
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        replaceNameInSheet(newName, oldName, sheet);
        }
    }

    // Update the Member Points Sheet and reset the name fields in the form
    updateMemberPoints();
    SHEET.getRange(OLD_NAME_CELL).getCell(1, 1).setValue("");
    SHEET.getRange(NEW_NAME_CELL).getCell(1, 1).setValue("");

    // Log the change in the console for tracking purposes
    Logger.log("Fixed " + oldName + " to " + newName);
}

// Replaces all instances of a name in a sheet with another
function replaceNameInSheet(newName, oldName, sheet)
{
    let index = 2;
    let cont = true;
    let name;
    let oldNameLower = oldName.toLowerCase();

    // While there are still responses
    while (cont)
    {
        // Get the name in the row
        cont = false;
        name = sheet.getSheetByName("Form Responses 1").getRange("B" + index).getCell(1, 1).getValue().toLowerCase();

        // If the name is not empty, keep looking, otherwise stop
        if (name != "")
        cont = true;
        else
        break;
        
        // If the name matches the old name, replace it with the new one
        if (name == oldNameLower)
        {
        sheet.getSheetByName("Form Responses 1").getRange("B" + index).getCell(1, 1).setValue(newName);
        }

        index++;
    }

    index = 2;
    cont = true;
    
    // Repeat the same process for the Responses sheet
    while (cont)
    {
        cont = false;
        name = sheet.getSheetByName("Responses").getRange("A" + index).getCell(1, 1).getValue().toLowerCase();
        if (name != "")
        cont = true;
        else
        break;
        
        if (name == oldNameLower)
        {
        sheet.getSheetByName("Responses").getRange("A" + index).getCell(1, 1).setValue(newName);
        }

        index++;
    }
}

// Constant locations for the deleteName form in the Control Panel
const DEL_NAME_CELL = "K20";
const DEL_EMAIL_CELL = "K22";
const DEL_ERROR_CELL = "J19";

// Deletes a member from all event sheets
function deleteName()
{
    // Get the name and email from the form
    let delName = SHEET.getRange(DEL_NAME_CELL).getCell(1, 1).getValue().toLowerCase();
    let delEmail = SHEET.getRange(DEL_EMAIL_CELL).getCell(1, 1).getValue().toLowerCase();

    // If either of the name fields are empty, display an error message and stop
    if (delName == "" || delEmail == "")
    {
        SHEET.getRange(DEL_ERROR_CELL).getCell(1, 1).setValue("One of the fields is empty").setFontColor("red");
        return;
    }
    
    // Get all of the files in the Sheets folder
    let sheets = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFilesByType("application/vnd.google-apps.spreadsheet");

    // Iterate through all of the sheets and delete the member
    while (sheets.hasNext())
    {
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        deleteNameInSheet(delEmail, delName, sheet);
    }

    // Get all of the subfolders in the Sheets folder
    let subfolders = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFolders();
    
    // Get all of the sheets in the subfolders
    while (subfolders.hasNext())
    {
        sheets = subfolders.next().getFilesByType("application/vnd.google-apps.spreadsheet");

        // Delete the member from each sheet
        while (sheets.hasNext())
        {
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        deleteNameInSheet(delEmail, delName, sheet);
        }
    }

    // Update the Member Points Sheet and reset the name fields in the form
    updateMemberPoints();
    SHEET.getRange(DEL_NAME_CELL).getCell(1, 1).setValue("");
    SHEET.getRange(DEL_EMAIL_CELL).getCell(1, 1).setValue("");

    // Log the change in the console for tracking purposes
    Logger.log("Deleted " + delName + " and " + delEmail);
}


// Deletes a member from all event sheets
function deleteNameInSheet(delEmail, delName, sheet)
{
    let index = 2;
    let cont = true;
    let name;
    let email;

    // While there are still responses
    while (cont)
    {
        // Get the name and email in the row
        cont = false;
        name = sheet.getSheetByName("Form Responses 1").getRange("B" + index).getCell(1, 1).getValue().toLowerCase();
        email = sheet.getSheetByName("Form Responses 1").getRange("C" + index).getCell(1, 1).getValue().toLowerCase();

        // If the name is not empty, keep looking, otherwise stop
        if (name != "")
        cont = true;
        else
        break;
        
        // If the name and email match the ones to delete, delete the row
        if (name == delName && email == delEmail)
        {
        sheet.getSheetByName("Form Responses 1").deleteRow(index);
        }

        index++;
    }

    index = 2;
    cont = true;
    
    // Repeat the same process for the Responses sheet
    while (cont)
    {
        cont = false;
        name = sheet.getSheetByName("Responses").getRange("A" + index).getCell(1, 1).getValue().toLowerCase();
        email = sheet.getSheetByName("Responses").getRange("B" + index).getCell(1, 1).getValue().toLowerCase();
        if (name != "")
        cont = true;
        else
        break;
        
        if (name == delName && email == delEmail)
        {
        sheet.getSheetByName("Responses").deleteRow(index);
        }

        index++;
    }
}

// Constant locations for the findUser form in the Control Panel
const FIND_USER_CELL = "Q4";
const USER_EVENTS_CELL = "M3";
const FIND_USER_ERROR_CELL = "P3";

// Finds all of the events a user has attended and the points received
function findUser()
{
    // Get the user from the form
    let user = SHEET.getRange(FIND_USER_CELL).getCell(1,1).getValue().toLowerCase();
    let events = "";
    let nextEvent = "";
    let memberPoints = 0;
    let numEvents = 0

    // If the user field is empty, display an error message and stop
    if (!user)
    {
        SHEET.getRange(FIND_USER_CELL).getCell(1,1).setValue("Please enter a member's full name");
        return;
    }

    // Clear out any attendance records currently in the form
    SHEET.getRange(USER_EVENTS_CELL).setValue("");
    
    // Get all of the files in the Sheets folder
    let sheets = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFilesByType("application/vnd.google-apps.spreadsheet");
    
    // Iterate through all of the sheets and get the events the user attended
    while (sheets.hasNext())
    {
        // If the user attended the event, get the information
        let sheet = SpreadsheetApp.openById(sheets.next().getId());
        nextEvent = getNameFromSheet(user, sheet);

        // If the user attended the event, add the event to the list, increase the number of total events, and increase total number of points
        if (nextEvent != undefined)
        {
            numEvents++;
            events = events + nextEvent + "\n";
            memberPoints += nextEvent[1];
        }
    }
            
    // Get all of the subfolders in the Sheets folder
    let subfolders = DriveApp.getFolderById(SHEETS_FOLDER_ID).getFolders();
    
    // Get all of the sheets in the subfolders
    while (subfolders.hasNext())
    {
        sheets = subfolders.next().getFilesByType("application/vnd.google-apps.spreadsheet");
            
        // If the user attended the event, get the information
        while (sheets.hasNext())
        {
            let sheet = SpreadsheetApp.openById(sheets.next().getId());
            nextEvent = getNameFromSheet(user, sheet);

            // If the user attended the event, add the event to the list, increase the number of total events, and increase total number of points
            if (nextEvent != undefined)
            {
                events = events + nextEvent[0] + "\n";
                memberPoints += nextEvent[1];
                numEvents++;
            }
        }
    }

    // If the user did not attend any events, display a message saying so
    if (events == "")
    {
        events = "No Attendance Record Found For " + user.toUpperCase();
    }
    // Otherwise, display the the name of the user, total number of points, and the attendance breakdown
    else
    {
        events = "Attendance Records For " + user.toUpperCase() + " (" + memberPoints + " Member Points)" + "\n\n" + events;
    }

    // If the user attended less than 20 events, set the font size to 12
    if (numEvents <= 20)
    {
        SHEET.getRange(USER_EVENTS_CELL).getCell(1,1).setValue(events).setFontSize(12);
    }
    // If the user attended less than 30 events, set the font size to 10
    else if (numEvents <= 30 && numEvents > 20)
    {
        SHEET.getRange(USER_EVENTS_CELL).getCell(1,1).setValue(events).setFontSize(10);
    }
    // If the user attended less than 40 events, set the font size to 8
    else if (numEvents <= 40 && numEvents > 30)
    {
        SHEET.getRange(USER_EVENTS_CELL).getCell(1,1).setValue(events).setFontSize(8);
    }
    // If the user attended more than 40 events, set the font size to 7
    else
    {
        SHEET.getRange(USER_EVENTS_CELL).getCell(1,1).setValue(events).setFontSize(7);
    }

    // Reset the user field in the form
    SHEET.getRange(FIND_USER_CELL).getCell(1,1).setValue("");

    // Log the search in the console for tracking purposes
    Logger.log("Searching for " + user + "'s Attendance Records");
}

// Check the event sheet for the user
function getNameFromSheet(user, sheet)
{
    let cont = true;
    let name = "";
    let index = 2;
    let point = 0;
    let event = "";

    // Iterate through all of the responses
    while (cont)
    {
        // Get the name from the row
        cont = false;
        name = sheet.getSheetByName("Responses").getRange("A" + index).getCell(1, 1).getValue().toLowerCase();

        // If the name is not empty, keep looking, otherwise stop
        if (name != "")
        {
            cont = true;
        }
        else
        {
            break;
        }

        // If the user attended the event, get the total points the user received
        if (user == name)
        {
            point = Number(sheet.getSheetByName("Responses").getRange("D" + index).getCell(1, 1).getValue()) + Number(sheet.getSheetByName("Responses").getRange("F2").getCell(1, 1).getValue());

            // If the user only received 1 point, set the description to "Member Point"
            if (Number(point) == 1)
            {
                event = DriveApp.getFileById(sheet.getId()).getName().split(" Member Sheet")[0] + " - " + point + " Member Point";
            }
            // Otherwise set it to "Member Points"
            else
            {
                event = DriveApp.getFileById(sheet.getId()).getName().split(" Member Sheet")[0] + " - " + point + " Member Points";
            }
            return ([event, Number(point)]);
        }

        index++;
    }
}
