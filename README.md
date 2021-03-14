# Excel4DnA
# Welcome! You found here!
You are a data engineer or a data analyst and probably a fan of AutoHotkey(or equivalent) as well.
Or you just want to find how to turn a table to json.

OK,
but you are not allowed to use/install any other favourite tool than... to use Excel.

These hacky scripts are for you.

# Why these VBA scripts?
- To save time and stay cool

As a data **engineer**/analyst, here is a series of VBA scripts I used to speed up my development, analysis on daily basis. e.g: 
- a script to turn a excel table(Range) into a json format;
- a script to turn cells to IN sql statement
- a script to turn table to filter statement
- ...



# How?
## Express
1. Copy file from .\dist\PERSONAL.XLSB to C:\Users\\_user name_\AppData\Local\Microsoft\Excel\XLStart 
2. and start to use it.

## Express + Excel Toolbar
1. Follow **_Express_** and mark down the XLStart path.
2. Setup your Quick Access Toolbar.
    1. Make sure you have the right path to replace _PATH_TO_EXCEL_XLSTART_ **Customizations.exportedUI** in the .\dist folder
    2. Open Excel > File > Options > Quick Access ToolBar > Import 
    3. In this dialogue box, above the OK button, click Import/Export > select **Customizations.exportedUI** in the .\dist folder
    4. Enjoy the shortcut buttons in the Quick Access Toolbar.
