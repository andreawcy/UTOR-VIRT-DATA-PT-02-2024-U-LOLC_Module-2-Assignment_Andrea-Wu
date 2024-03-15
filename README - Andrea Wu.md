Name: Andrea Wu
Module 2 Assignment

Workflow:
1. loop through all worksheets
# Reference: https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

Within each worksheet
2. Print the header for for summary table first, so users don't need to type it in each sheet.
3. Set variables' names
4. Identify the numbers of row the loop need to run through the sheet.
5. Assign a variable to keep track on the row of summary table.
6. Use a For loop to check through every row. For each row, the script will check if...
    a. it is the first row of each ticker
        i. Assign open price
        ii. Add volume to total volume
    b. it is last row of each ticker
        i. Assign close price, ticker name
        ii. Add volume to total volume
        iiI. Do the calculation and update format for price change, % change, total volume
        iv. Print variables to summary tables
        v. Store variable to greatest increase/decrease/volume if it is larger than the previous ticker.
        vi. Reset the variables used to store summary table
    c. it is the row in between the first and last row of ticker
        i. Add volume to total volume
# Reference for step 6: 
# Loop and summary table: Class activity - week 2 class 2 activity 6
# Format cell: https://www.wallstreetmojo.com/vba-format/
 7. Print the greatest increase/decrease/total volume and reset the variables for next worksheet.
 8. Adjust column width to fit the long numbers of total volume.
 # Reference: https://www.automateexcel.com/vba/row-height-column-width/#autofit-column-width   
 9. Show message to inform user the sheet has been updated and go to the next sheet.