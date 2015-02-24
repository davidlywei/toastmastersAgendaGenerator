/****************************************************************************************************
 Description:     This program takes inputs from the companion Google Form and 
                  converts those answers into a Flying Toaster's agenda, and 
                  then emails it to the email address provided to the form.
                  
                  Whenever the submit button is pressed on the companion Google 
                  Form, generateAgenda() is triggered. Based off of the responses
                  given in the form, the generateAgenda() picks an appropriate 
                  template file, and proceeds to replace all tags listed in the 
                  template file with appropriate responses gathered from the user.
                  Note, the text format of the tag (bold, italics, underline) will
                  also be applied to the replacement text.
                  
                  The tags used, and their respective replacements are listed below:
                  {DATE}  - Date of the next meeting
                  {THEME} - The theme of the next meeting
                  {SAA}   - Sergeant at Arms
                  {TMOD}  - Toastmaster of the Day
                  {JKM}   - Jokemaster
                  {GE}    - General Evaluator
                  {TMR}   - Timer
                  {AC}    - Ah Counter
                  {GRA}   - Grammarian
                  {WOD}   - The Word of the Day
                  {TTM}   - Table Topics Master
                  {TTMH}  - Table Topics Starting Time (hours)
                  {TTMM}  - Table Topics Starting Time (minutes)
                  {SP#}   - Name of Speaker where # is replaced by the speaker number
                  {PN#}   - Project Numbers ""                         ""
                  {ST#}   - Speech Titles   ""                         "" 
                  {DR#}   - Speech Duration ""                         ""
                  {EVL#}  - Evaluators      ""                         ""

                  Not all tags need to be in the template for the program to work.

                  Features:
                    - The duration of the speeches are generated programmatically
                    - A feature of Google Forms is that whenever a form is submitted
                      the responses to the form are stored into a Google Spreadsheet.
                      This can be used to help the VP of Education track who has filled
                      what roles.
                    - The output Agenda is named "(#) Flying Toaster's Agenda"
                      where the # corresponds to the row number in the Google Spreadsheet
                      that produced that agenda.
                    - The form checks the entered email address ensuring that it fits
                      the REGEX ".*@.*"
                    - Automatically calculates Table Topics start time based off of
                      number of speeches, and Speech start time.
                    - Added ability to include a tailing greeting letter to be printed
                      on the backside of the agenda.
                  
                  Known Issues:
                    - The duration calculator only uses the times from the CC
                      manual. More advanced manuals will require manual edits.
                      Can be implemented programmatically if need arises.
                    - The program only allows meetings with a maximum of 2 
                      speeches. More speeches can be added if need arises.
                    - The name of the Sergeant at Arms is hard coded into the 
                      the script so whenever officer elections happen someone
                      will need to change that value. Not sure if it is
                      worth creating a file to store this one name.
                    - To anyone hosting their own version of this script, note
                      there is a bug with Gmail sending mail to yourself. The
                      bug is noted here:
                      https://code.google.com/p/google-apps-script-issues/issues/detail?id=2966
                    - For some reason when you look at the generated agenda from 
                      the folder link, it shows the agenda as taking up two pages 
                      for the two speakers version. However if you use the direct
                      link to the document, it shows it as one page. I have no idea
                      why. (Possibly conversion from Google Docs to Microsoft Doc?)

  Author:         David Wei
  Date:           Jan  8, 2015
  Last Modified:  Jan 14, 2015
  
  Changelog:      - Added lines to change the generated document's permission to 
                    "Anyone with a link can Edit" (1/9/15)
                  - Changed mailing app from GmailApp to MailApp because GmailApp
                    is overkill in its features (1/9/15)
                  - Changed file naming format from "... (#)" to "(00#) ..." (1/9/15)
                  - Changed common variables to global variables. (1/10/15)
                  - Separated out Email Code into it's own function (1/10/15)
                  - Added feature to automatically adjust Table Topics start time
                    based off of number of speeches and speech start time. (1/10/15)
                  - Removed questionable code, and replaced it with less
                    questionable code (1/14/15)
                  - Made the Sergeant at Arms's name determined by a spreadsheet
                    cell instead of hard coded. (1/19/15)

****************************************************************************************************/

//----------------------------------------------------------------------------------
// Global Constants
//----------------------------------------------------------------------------------  

//-------------------------------------------------------------------------
//                                             Sergeant at Arms Declaration
//-------------------------------------------------------------------------  

// Open Flying Toaster's Meeting Roles Spreadsheet
var sheet = SpreadsheetApp.openById(
                  '<INSERT_DOC_ID_HERE>'
                                    );

// Get the values into the data Array.
var data = sheet.getDataRange().getValues();

// Set the name of current Sergeant at Arms
var SERGEANT_AT_ARMS = data[3][0];

//-------------------------------------------------------------------------
//                                                        Set Agenda Folder
//-------------------------------------------------------------------------

// These variables define the folder where the generated 
// agendas will be sent.

// The ID for the Folder where all the generated agendas will go
var FOLDER_ID = '<INSERT_FOLDER_ID_HERE>';
// set agenda folder for later use
var AGENDA_FOLDER = DriveApp.getFolderById(FOLDER_ID);

//-------------------------------------------------------------------------
//                                                     Set Agenda Templates
//-------------------------------------------------------------------------

// Agenda Templates are the templates used to generate Agendas.
// They are separated based off of how many speakers each meeting
// has.
// To change the number of speakers, you need to change
//   - The number of options in the Google Form
//   - The number of template options below
//   - The number of speaker indicator variables
//   - The options in setTemplateID()
//   - The value of MAX_NUM_SPEECHES if necessary.

// The ID for the file that will be used as a template
var NO_SPEAKERS  = '<INSERT_DOC_ID_HERE>';
var ONE_SPEAKER  = '<INSERT_DOC_ID_HERE>';
var TWO_SPEAKERS = '<INSERT_DOC_ID_HERE>';

// Speaker Indicator variables
// These variables code the indicator used as a response
// for the question "How many speakers are there"
var NO_SPEAKER_INDICATOR  = 'None';
var ONE_SPEAKER_INDICATOR = 'One Speaker';
var TWO_SPEAKER_INDICATOR = 'Two Speakers';
  
// Set max number of speeches that is ever possible
var MAX_NUM_SPEECHES      = 2;

//-------------------------------------------------------------------------
//                                             Set Greeting Letter Template
//-------------------------------------------------------------------------

// Set Greeting Letter Template
var GREETING_LETTER_ID = '<INSERT_DOC_ID_HERE>';

//-------------------------------------------------------------------------
//                                               Set Response Array Indices
//-------------------------------------------------------------------------

// Constants for response array indices
// --- The numbers for each index is determined by the question
// --- number in the survey. There is a 1 to 1 correlation between
// --- the question number, and the index number. 
// --- (0 based counting of course)
var EMAIL_INDEX           =  0;
var DATE_INDEX            =  1;
var THEME_INDEX           =  2;
var TMOD_INDEX            =  3;
var JOKEMASTER_INDEX      =  4;
var GE_INDEX              =  5;
var TIMER_INDEX           =  6;
var AH_COUNTER_INDEX      =  7;
var GRAMMARIAN_INDEX      =  8;
var WORD_OF_THE_DAY_INDEX =  9;
var TTM_INDEX             = 10;
var NUM_SPEAKERS_INDEX    = 11;

//-------------------------------------------------------------------------
//                                                       Set Speaker Offset
//-------------------------------------------------------------------------

// The speaker offset is the number of items required to define a speech
// In this case, this script counts:
//     - Speaker Name
//     - Speech Title
//     - Project Number
//     - Evaluator
// This is used to calculate the speaker's Response Array index.
var SPEAKER_OFFSET        = 4;  

//-------------------------------------------------------------------------
//                                                  Table Topics Start Time
//-------------------------------------------------------------------------

// To calculate the table topics start time, we simply take the maximum
// time allotted for each speech, add the appropriate number of 
// minutes for evaluation, then add the resultant minutes to the starting
// time of speeches.

// Starting time of Speeches, rounded to nearest minute
var SPEECH_START_HOUR     =  6;
var SPEECH_START_MIN      = 57;

// Time given to do evaluation, rounded up to nearest minute.
// --- These are applied ONCE PER SPEECH
var EVALUATION_TIME       = 1;
var INTRODUCTION_TIME     = 1;
// --- Buffer time is applied ONCE
var BUFFER_TIME           = 0;

//----------------------------------------------------------------------------------
// Script
//----------------------------------------------------------------------------------  

/*************************************************************************

    Function Name: generateAgenda

    Purpose:       The launching off point for the agenda generator 
                   script. This function grabs responses from the forms,
                   copies the appropriate template, replaces the tags in
                   that template, and emails the completed template to 
                   email provided from the form
                   
    Parameters:	   N/A

*************************************************************************/
function generateAgenda() 
{
  // set form to current current generate agenda form
  var form = FormApp.getActiveForm();
  
  // get array of All Responses
  var formResponses = form.getResponses();
  // --- get latest response
  var latestItemResponses = formResponses[formResponses.length - 1];
  // --- get individual responses array from latest responses
  var itemResponse = latestItemResponses.getItemResponses();

  // Set the templateID based off of how many speakers
  // there are.
  var templateID = setTemplateID(itemResponse);
  
  // Replace text in a generated agenda, and return the 
  // copy's ID
  var copyID = replaceText(itemResponse, formResponses, templateID)

  // Append the Greeting Letter to the end of the document
  appendGreeting(copyID);
  
  // Email the agenda
  emailAgenda(itemResponse, copyID);
}

/*************************************************************************

    Function Name: setTemplateID

    Purpose:       This function chooses which template to use based off
                   of how many speakers will be present at the meeting.
                   Returns the ID of the template chosen 

    Parameters:
      itemResponse : itemResponse is the array of responses gathered 
                     from the form

*************************************************************************/
function setTemplateID (itemResponse)
{
  // Set templateID based off of how many speakers there are.
  var templateID;
  // --- Check how many speakers, and choose appropriate template
  if(itemResponse[ NUM_SPEAKERS_INDEX ].getResponse() == TWO_SPEAKER_INDICATOR)
  {
    templateID = TWO_SPEAKERS;
  }
  if(itemResponse[ NUM_SPEAKERS_INDEX ].getResponse() == ONE_SPEAKER_INDICATOR)
  {
    templateID = ONE_SPEAKER;
  }
  if(itemResponse[ NUM_SPEAKERS_INDEX ].getResponse() == NO_SPEAKER_INDICATOR)
  {
    templateID = NO_SPEAKERS;
  }
  
  return templateID;
}

/*************************************************************************

    Function Name: replaceText

    Purpose:       This function handles copying the template, and 
                   replacing the predefined tags with form responses.
                   Returns the ID of the copied document

    Parameters:
      itemResponse : itemResponse is the array of responses gathered 
                     from the form
                     
      formResponse : formResponse is the array of all responses. Used
                     to generate the row number.
                     
      templateID   : The id value of the document that will be sent out

*************************************************************************/
function replaceText(itemResponse, formResponses, templateID)
{
  // make a copy of the Toastmaster's Template
  // --- formResponses.length + 1  is added to the name
  // --- so each agenda can be paired to a row number
  // --- in the role sheet.
  var copyID = DriveApp.getFileById(templateID)
       .makeCopy(
                    // Add the row number in with up to 3 leading 0s
                    '(' + ('000' + (formResponses.length + 1).toString()).slice(-4) + ') ' 
                    + 'Flying Toaster\'s Agenda', AGENDA_FOLDER 
                )
       .getId();
  
  // Open copied document
  var docBody = DocumentApp.openById(copyID).getBody();

  // Replace text for roles
  replaceRoles(docBody, itemResponse);
  
  // Time necessary for speeches. Determines when to start Table Topics.
  var tableTopicsTimeMin = SPEECH_START_MIN;

  // Set number of speeches
  var numSpeeches = ((itemResponse.length) - (NUM_SPEAKERS_INDEX + 1)) / SPEAKER_OFFSET;
  
  // Text replacement for speaker roles
  for(i = 1; i <= numSpeeches; i++)
  {
      // replace text for speaker roles
      tableTopicsTimeMin = replaceSpeakerRoles(docBody, itemResponse, i, tableTopicsTimeMin);
  
      // add evluation time.
      tableTopicsTimeMin += EVALUATION_TIME + INTRODUCTION_TIME;  
  }
  
  // add buffer time
  tableTopicsTimeMin += BUFFER_TIME;
  
  // change tableTopicsTimeMin to hour minute format
  var tableTopicsTimeHour = SPEECH_START_HOUR + Math.floor(tableTopicsTimeMin / 60);
  tableTopicsTimeMin %= 60;
  
  // set table topics start time.
  docBody.replaceText('{TTMH}'   , tableTopicsTimeHour.toString());
  docBody.replaceText('{TTMM}'   , ('00' + tableTopicsTimeMin.toString()).slice(-2));
  
  // return the copyID
  return copyID;
}

/*************************************************************************

  Function Name: replaceRoles

  Purpose:       This function replaces all the tags listed in the 
                 template that are NOT associated with the speaker.
                 
                 ie. This includes all roles EXCEPT the speaker(s), and 
                     their evaluator(s)

  Parameters:
    docBody      : docBody is the body of the document of the agenda
    
    itemResponse : itemResponse is the array of responses gathered 
                   from the form
                   
*************************************************************************/
function replaceRoles(docBody, itemResponse)
{
  // SERGEANT at Arms Text Replacement
  docBody.replaceText('{SAA}'   , SERGEANT_AT_ARMS);
  
  // Text replacement in the Template Document
  docBody.replaceText('{DATE}'  , itemResponse[ DATE_INDEX              ].getResponse());
  docBody.replaceText('{THEME}' , itemResponse[ THEME_INDEX             ].getResponse());
  docBody.replaceText('{TMOD}'  , itemResponse[ TMOD_INDEX              ].getResponse());
  docBody.replaceText('{JKM}'   , itemResponse[ JOKEMASTER_INDEX        ].getResponse());
  docBody.replaceText('{GE}'    , itemResponse[ GE_INDEX                ].getResponse());
  docBody.replaceText('{TMR}'   , itemResponse[ TIMER_INDEX             ].getResponse());
  docBody.replaceText('{AC}'    , itemResponse[ AH_COUNTER_INDEX        ].getResponse());
  docBody.replaceText('{GRA}'   , itemResponse[ GRAMMARIAN_INDEX        ].getResponse());
  docBody.replaceText('{WOD}'   , itemResponse[ WORD_OF_THE_DAY_INDEX   ].getResponse());
  docBody.replaceText('{TTM}'   , itemResponse[ TTM_INDEX               ].getResponse());
}

/*************************************************************************

    Function Name: replaceSpeakerRoles

    Purpose:       This function handles the text replacement for the
                   the speaker, and accumulates time for start of 
                   Table Topics
                   returns earliest calculated Table Topics start time

    Parameters:
      docBody           : docBody is the body of the document of the agenda
    
      itemResponse      : itemResponse is the array of responses gathered 
                          from the form
               
      speakerNumber     : speakerNumber is an int that tells this function 
                          which speaker's tags it should replace.

      tableTopicsTimeMin: The current time needed for speeches.

*************************************************************************/
function replaceSpeakerRoles(docBody, itemResponse, speakerNumber, tableTopicsTimeMin)
{
  // Declaring speaker tag
  var speakerTag = '{SP'  + (speakerNumber).toString() + '}';
  var titleTag   = '{ST'  + (speakerNumber).toString() + '}';
  var projNumTag = '{PN'  + (speakerNumber).toString() + '}';
  var evalTag    = '{EVL' + (speakerNumber).toString() + '}';
  var durTag     = '{DR'  + (speakerNumber).toString() + '}';
   
  // Declaring speaker index variables
  var SPEAKER_NAME_INDEX     = 12 + ((speakerNumber - 1) * SPEAKER_OFFSET);
  var SPEAKER_TITLE_INDEX    = 13 + ((speakerNumber - 1) * SPEAKER_OFFSET);
  var SPEAKER_PROJ_NUM_INDEX = 14 + ((speakerNumber - 1) * SPEAKER_OFFSET);
  var SPEAKER_EVAL_INDEX     = 15 + ((speakerNumber - 1) * SPEAKER_OFFSET);
  
  // Replace text of speaker
  docBody.replaceText( speakerTag, itemResponse[ SPEAKER_NAME_INDEX     ].getResponse());
  docBody.replaceText( titleTag  , itemResponse[ SPEAKER_TITLE_INDEX    ].getResponse());
  docBody.replaceText( projNumTag, itemResponse[ SPEAKER_PROJ_NUM_INDEX ].getResponse());
  docBody.replaceText( evalTag   , itemResponse[ SPEAKER_EVAL_INDEX     ].getResponse());
  
  // --- Auto fill in the speech duration
  var sp1Duration;
  
  switch(parseInt(itemResponse[ SPEAKER_PROJ_NUM_INDEX ].getResponse())) 
  {
    case 1:  // speech 1
      sp1Duration = "4 - 6";
      tableTopicsTimeMin += 6;
      break;
      
    case 10: // speech 10
      sp1Duration = "8 - 10";
      tableTopicsTimeMin += 10;
      break;
      
    default: // all other speeches
      sp1Duration = "5 - 7";
      tableTopicsTimeMin += 7;
  }
  
  // fill in speech duration
  docBody.replaceText(durTag   , sp1Duration + ' mins');
  
  return tableTopicsTimeMin;
}

/*************************************************************************

    Function Name: appendGreeting

    Purpose:       This function appends the greeting letter to the end
                   of the meeting agenda.

    Parameters:
      copyID       : The id value of the document where we will append
                     the greeting.

*************************************************************************/
function appendGreeting(copyID)
{
  // Open copied document
  var docBody = DocumentApp.openById(copyID).getBody();

  // Merge body of document with greeting letter document
  // --- Get the text of the Greeting letter
  var greetingText = DocumentApp.openById(GREETING_LETTER_ID).getBody().getParagraphs();
  // --- Add a pagebreak
  docBody.appendPageBreak();
  // --- Merge it with the current document.
  for(i = 0; i < greetingText.length; i++)
  {
    docBody.appendParagraph(greetingText[i].copy());
  }
}

/*************************************************************************

    Function Name: emailAgenda

    Purpose:       This function handles emailing the agenda out after
                   it has been generated.

    Parameters:
      itemResponse : itemResponse is the array of responses gathered 
                     from the form
               
      copyID       : The id value of the document that will be sent out

*************************************************************************/
function emailAgenda(itemResponse, copyID)
{
  // Set sharing permissions to (anyone with a link can edit)
  var editableDoc = DriveApp.getFileById(copyID);
  editableDoc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  
  // Get the URL of the document.
  var url = editableDoc.getUrl();
  
  // Get the name of the document to use as an email subject line.
  var subject = "Flying Toaster's Meeting Agenda";
  
  // Append a new string to the "url" variable to use as an email docBody.
  var emailBody = 'Thank You ' + itemResponse[ TMOD_INDEX ].getResponse() + ' \
  for volunteering to be the Toastmaster of the Day. \nHere is the link to the meeting agenda: ' + url;
  
  // Send yourself an email with a link to the document.
  MailApp.sendEmail( itemResponse[ EMAIL_INDEX ].getResponse(), subject, emailBody);
}
