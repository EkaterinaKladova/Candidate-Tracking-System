As my sorority grew, so did the number of candidates.  Eventually, it was taking the person responsible for them multiple hours each day to track their progress through candidacy.  \
For context, the candidates were required to collect points by going on a certain number of sister dates per week.  Each sister would fill out a form at the end of the date, and enter the "sister password."  Then, it was up to the "membership educator" to make sure the points/dates were displayed on the candidate's spreadsheet.

This project was created to:
- Automate the intital creation of the individual spreadsheets for all candidates
- Automate the collection and dissemination of the points/dates that each candidate collects

The MED Spreadsheet has the overall information on every candidate, which it collects via a connected form.  It is also where the individual candidate spreadsheets pull the data for their candidate from.  
The instructions for using the automation can be found in the Instructions doc, which was left in the organization as guidance for future members.

Problem-specific challenges:
- Creating a customizable program that can be reused by a new person each year
- Minimizing the amount of code that needs to be edited every year, assuming the person using the automation has no experience with code
- Ensuring no one except the people responsible for the candidates have a way to access the spreadsheets

The first two challenges were solved by limiting the number of variables that were dependant on outside factors, which left us with 3 that would need to be changed annually.  They were placed at the top of the code, with commented instructions for how to update them.  The last one was solved by emphasizing in the instructions that the script needs to be copied into a private google drive, and ensuring that the created spreadsheets are only shared with the people they are supposed to be seen by.
