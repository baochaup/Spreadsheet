Author: Bao Chau Pham
Date: 10/04/2019

To me, PS6 is difficult. I couldn't figure out how to implement selecting a different cell using arrow keys.
I tried adding KeyDown event handler for arrow keys, but when running, it just moves the scrollbars of the spreadsheet panel.
So, I thought that I would modify SpreadsheetPanel class, but I didn't have enough time.
The additional feature I added is generating a graph for the data in the spreadsheet.

Guide:
To update the value of a cell, select a cell, enter the contents, and press Enter or click Change.
To create a new spreadsheet, go to File/New
To save the current spreadsheet, go to File/Save
To save the current spreadsheet with a different name, go to File/Save as...
Additional feature:
Generating a data graph for the current spreadsheet.
To generate, go to Tools/Generate data graph
Constraints: this feature can generate only 3 series based on the columns B, C, and D.
The series' names are the values of B1, C1, and D1.
The values of the cells below B1, C1, D1 must be numbers.