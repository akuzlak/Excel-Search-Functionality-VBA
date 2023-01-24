# Excel Search Functionality with VBA

This repository is about a search functionality within Microsoft Excel using the VBA with the leading of renowned AI ChatGPT.

## What are the advantages over classical Excel's built-in search?

It is always possible to search for what you are looking for in Microsoft Excel by using the built-in ''Find and Replace'' function. However, this VBA code allows you to search within a sheet and bring all the results in another sheet in a tabular format by using the original titles. This is important because when using the built-in ''Find and Replace'' function, you can take the results manually to another sheet instread of in one-click.

## How to use this VBA?

1) You need to crate two sheets in an Macro-enabled Microsoft Excel file. One sheet is to insert search buttons and the other sheet is for storing the data to search. Eventually a third sheet will be created by the VBA called Results to show the search results. You can always change the name of this Results sheet.
2) In the search fucntionaility sheet, create two Active X controls; TextBox and Command Button

## What should be changed in the code before using it?

As this code is created for a specific project, sheet names and range in the code must be replaced to prevent receiving errors. Thus, make the following changes in the code;
1) Replace Search_Data_Catalog sheet name. This is the sheet that you should insert the TextBox and Command Button from Active X control group. So replace this sheet name to your sheet name where you use these Active X control utilities.
2) Replace Data_Catalog sheet name. This is the sheet that your data resides, in other words, this is the sheet that our search functionality will search for the word(s) you are looking for. Again, replace this with the sheet name you have where the data resides. Note that this name is in two different places, replace both.
3) Replace the search range of A:AD to your sheet's range. Note that these are in two different places, so replace both.

## What does it exactly do?

This code searches the search criteria in specified Excel sheet and if found shows a message box about how many results are found. If nothing found, again shows a message box that nothing has beedn found. The It creates a new sheet called Results and list the rows that includes the search criteria somewhere. Also the title of the search sheet is included to make sense of which data/row is coming from which field/column. If a new search is executed, then Results sheet is cleaned and filled with the new results only when there are new results. If nothing has been found with the new search, then Results sheet remain intact unless a search result ends up bringing something.

## How is this VBA created?

This VBA is created by an iterative approach using ChatGPT AI tool made public on November 2022 and as of 24 January 2023, still free to use. By giving commends to ChatGPT and getting help from it when errors encountered, after around 4-6 hours of work which includes around 25-35 iterations, this VBA code is created.
