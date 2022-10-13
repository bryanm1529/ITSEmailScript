# ITSEmailScript
Email Script to parse emails for new ellucian banner releases and then import them into excel spreadsheet

# How it works
- The script will parse emails all emails that return after the search 'Ellucian'. Then it will check if the banner releases in those emails are in the spreadsheet exactly at which point it will just update the repost date. If the email contains certain keywords like Austrailia it will automaically import them into the excel sheet as not rcnj related with proper coloring and email date. If the spreadsheet is unsure of a new release it will ask the user if the release is rcnj-related so it can import it into the sheet accordingnly. 
<img width="751" alt="image" src="https://user-images.githubusercontent.com/68625150/195638932-5f02eae5-abfe-47d3-b6a5-988cd978203a.png">

Once it is done running the spreadsheet will be updated with all of the releases it found! 


# Installation
- Have python installed 
- Download zip file of code by pressing on top right green code button and then download as zip


# How to Run
- Once downloaded Open a terminal window and cd to the directory where 'updatespreadsheet.py' is located. It should be in the the ITSEmailScript-main folder zip once extracted. 
`cd ITSEmailSCript-main` in the terminal to get there
- Once you are in the right directory you can run the script with `python3 updatespreadsheet.py`in the terminal

