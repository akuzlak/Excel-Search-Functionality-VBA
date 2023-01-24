# Excel Search Functionality with VBA

This repository is about a search functionality within Microsoft Excel using the VBA with the leading of renowned AI ChatGPT.

## What are the advantages over classical Excel's built-in search?

It is always possible to search for what you are looking for in Microsoft Excel by using the built-in ''Find and Replace'' function. However, this VBA code allows you to search within a sheet and bring all the results in another sheet in a tabular format by using the original titles. This is important because when using the built-in ''Find and Replace'' function, you can take the results manually to another sheet instread of in one-click.

## How to use this VBA?

You need to crate two sheets in an Macro-enabled Microsoft Excel file. One sheet is to insert search buttons and the other sheet is for storing the data to search. Eventually a third sheet will be created by the VBA called **Results** to show the search results. You can always change the name of this Results sheet.

## What should be changed in the code before using it?

As this code is created for a specific project, sheet names in the code must be replaced by sheet names that your Excel has so that they match with your sheet names to prevent receiving errors. Thus change sheet names that need to change; 
1) This is the sheet that you should insert the TextBox and Command Button from Active X control group. So change this sheet name to your sheet name that you use these Active X control utilities.
2) This is the sheet that your data resides, in other words, this is the sheet that our search functionality will search for the word(s) you are looking for. Again, change this with the sheet name you have where the data resides

## How is this VBA created?

This VBA is created by an iterative approach using ChatGPT AI tool made public on November 2022 and as of 24 January 2023, still free to use. By giving commends to ChatGPT and getting help from it when errors encountered, after around 4-6 hours of work which includes around 25-35 iterations, this VBA code is created.
