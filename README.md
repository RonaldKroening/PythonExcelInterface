# Python Excel Interface

The objective of this project is to allow a user to search through an Excel file and gather information based on a specific column. A sample call would use the following arguements:

1. A character that represents the column in the Excel file that contains the references to the values you wish to view.
2. The row BEFORE the values are listed. Most Excel files have one row with titles of each columns preceding the rows of information. For instance, if rows 3-10 are filled with information and the row with titles is row 2, then you would use 2 as your arguement.
3. A character that represents the column in the Excel file that contains the data you wish to view.

If the information is found at the value, the reference to the value and the value itself are printed. If not, it searches the internet for it. This only requires openpyxl and webbrowser.
