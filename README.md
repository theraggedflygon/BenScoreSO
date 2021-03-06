# BenScoreSO
Scoring system being debuted at the 2022 Gopher Invitational tournament.

## Setup
A template spreadsheet has been attached in Excel format. Download this spreadsheet through a method of your choice and upload it to Google Sheets. Make sure to convert the file to a Google Sheets file.

In the Google Sheets editor, under the "Extensions" menu in the top ribbon click "Apps Script". This will open up the Google App Scripts editor where there will be a blank .gs file, which is a modified JavaScript format that runs most Vanilla JS code. Copy and past the contents of `benscore.gs` into that file and save the changes. Return to the Google Sheet and reload it. You should now see a "Macros" menu in the top ribbon. The scoring system is now ready for use.

## Entering Teams and Events
Navigate to the "Data" tab of the scoresheet and enter the team names and event names in the desired order of how they should appear on the final scoresheets. Overwrite the sample "Team 1", "Team 2", "Event 1", etc. placeholders currently there. The system supports up to 500 teams and 50 events.

Once teams and events have been entered, go under the Macros menu and run "Setup BenScore". You will need to log in with your Google account to authorize the Macros. After you have authorized the script, if the script does not automatically run, please click it again and it will generate the scoring sheet as well as the indvidual event sheets.

## Setting Up Event Sheets
Each event will now have a mostly blank event sheet that needs to be setup. For each question in the event's test, please enter a point value by overwriting the "0" in the points column. Begin with the question numbered "1" and DO NOT skip any rows. All questions must have a non-zero point value. If you so desire you can enter a "nickname" for the question if "Question #" is not sufficient in identifying what you are talking about (this is particularly useful in the case of build events or an event like Write It, Do It). Lastly, assign tiebreaker priority by overwriting the "None" in the tiebreaker column. The first tiebreaker should have a value of "1", the second tiebreaker should have a value of "2" and so on. Only use each tiebreaker priority value once and ever question does not need to be assigned a tiebreaker priority.

Each event sheet supports up to 500 questions.

After entering the point values, tiebreaker priority, and question nicknames (if desired) navigate to the Macros menu once again and run the "Setup Event Scoresheet" macro. You will now have a proper scoresheet that can be used by event proctors on the day of the tournament.

## Proctoring Events
Each event scoresheet consists of a matrix with each team constituting the rows and each question constituting a column. Going through each question assign each team the points they are awarded on that question by overwriting the value of "0" that is serving as a placeholder. Do this for every question for every team taking the test.

After all teams have had their tests scored, run the Macro "Rank Event Teams" which will rank teams based on their exam score first and then in the order of the tiebreaks. If two teams happen to still be tied, the tie will be broke in favor of the team with the lower team number. This can be manually overwritten by altering the rankings generated by this on both the event score sheet and the main "Scoring" sheet of the spreadsheet.

## Getting Final Rankings
After all events have been scored and the "Scoring" sheet is fully populated with event results, run the "Sort Tournament Results" Macro to get the final rankings on the "Scoring" sheet.
