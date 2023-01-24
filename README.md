# Excel Search Functionality with VBA

This repository is about a search functionality within Microsoft Excel using VBA with the guidance of the AI model ChatGPT.

## What does it exactly do?

This code searches the search criteria in the specified Excel sheet and if found shows a message box about how many results are found. If nothing is found, it shows a message box that nothing has been found. It creates a new sheet called "Results" and lists the rows that include the search criteria there. Also, the title of the search sheet is included to make sense of which data/row is coming from which field/column. If a new search is executed, then "Results" sheet is cleaned and filled with the new results only when there are new results. If nothing is found with the new search, then "Results" sheet remains intact unless a search result ends up bringing something.

## Advantages over classical Excel's built-in search

It is always possible to search for what you are looking for in Microsoft Excel by using the built-in "Find and Replace" function. However, this VBA code allows you to search within a sheet and bring all the results to another sheet in a tabular format using the original titles. This is important because when using the built-in "Find and Replace" function, you have to manually transfer the results to another sheet instead of doing it with one click.

## How to use this VBA

1) Create two sheets in a Macro-enabled Microsoft Excel file. One sheet is for inserting search buttons, and the other sheet is for storing the data to search. Eventually, a third sheet will be created by the VBA called "Results" to show the search results. You can always change the name of this Results sheet.
2) In the search functionality sheet, create two Active X controls: a TextBox and a Command Button.

## What should be changed in the code before using it

As this code is created for a specific project, sheet names and range in the code must be replaced to prevent receiving errors. Thus, make the following changes in the code:

1) Replace "Search_Data_Catalog" sheet name. This is the sheet that you should insert the TextBox and Command Button from Active X control group. So replace this sheet name with the sheet name where you use these Active X control utilities.
2) Replace "Data_Catalog" sheet name. This is the sheet where your data resides; in other words, this is the sheet that our search functionality will search for the word(s) you are looking for. Again, replace this with the sheet name you have where the data resides. Note that this name is in two different places, so replace both.
3) Replace the search range of A:AD with your sheet's range. Note that these are in two different places, so replace both.

## How is this VBA created?

This VBA is created by an iterative approach using the ChatGPT AI tool. By giving commands to ChatGPT and getting help from it when errors are encountered, after around 4-6 hours of work which includes around 25-35 iterations, this VBA code is created.
