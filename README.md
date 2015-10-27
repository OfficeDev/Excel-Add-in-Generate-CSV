# CSV generator Task Pane Add-in Sample for Excel 2016

_Applies to: Excel 2016_

This task pane add-in shows how to create table from list of column names using the JavaScript APIs in Excel 2016. It comes in two flavors: code editor and Visual Studio.

![CSV Generator Sample](Images/ScreenCap1.PNG)

## Try it out
### Code editor version

The simplest way to deploy and test your add-in is to copy the files to a network share.

1.  Create a folder on a network share (for example, \\\MyShare\CollegeBudgetTracker) and copy all the files in the Code Editor folder. 
2.  Edit the <SourceLocation> element of the manifest file so that it points to the share location from step 1. 
3.  Copy the manifest (CollegeBudgetTrackerManifest.xml) to a network share (for example, \\\MyShare\MyManifests).
4.  Add the share location that contains the manifest as a trusted app catalog in Excel.

    a.  Launch Excel and open a blank spreadsheet.  
    
    b.  Choose the **File** tab, and then choose **Options**.
    
    c.  Choose **Trust Center**, and then choose the **Trust Center Settings** button.
    
    d.  Choose **Trusted Add-in Catalogs**.
    
    e.  In the **Catalog Url** box, enter the path to the network share you created in step 3, and then choose **Add Catalog**.
    
   f.  Select the **Show in Menu** check box, and then choose **OK**. A message appears to inform you that your settings will be applied the next time you start Office. 
        
5.  Test and run the add-in. 

    a.  In the **Insert tab** in Excel 2016, choose **My Add-ins**. 
    
    b.  In the **Office Add-ins** dialog box, choose **Shared Folder**.
    
    c.  Choose **College Budget Tracker Sample**>**Insert**. The add-in opens in a task pane and creates the college budget tracker in the active sheet as shown in this diagram. 
      
   ![College Budget Tracker Sample](images/CollegeBudgetTracker_tracker.PNG) 

    d.  Add some expenses and income using the **Add expenses** and the **Add income** tabs and see how the data and the charts changes dynamically.
    
      ![College Budget Tracker Sample](images/CollegeBudgetTracker_taskpane.PNG) 
    
### Visual Studio version
1.  Copy the project to a local folder and open the Excel-Add-in-JS-CollegeBudgetTracker.sln in Visual Studio.
2.  Press F5 to build and deploy the sample add-in. Excel launches and the add-in opens in a task pane to the right of a blank worksheet, as shown in the following figure. 
        
  ![College Budget Tracker Sample](images/CollegeBudgetTracker_tracker.PNG) 

3.  Add some expenses and income using the **Add expenses** and the **Add income** tabs and see how the data and the charts changes dynamically.

  ![College Budget Tracker Sample](images/CollegeBudgetTracker_taskpane.PNG) 


### Learn more

The Excel JavaScript APIs have much more to offer you as you develop add-ins. The following are just a few of the available resources. 

1.  [Excel Add-ins programming overview](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-programming-overview.md)
2.  [Snippet Explorer for Excel](http://officesnippetexplorer.azurewebsites.net/#/snippets/excel)
3.  [Excel Add-ins code samples](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-code-samples.md) 
4.  [Excel Add-ins JavaScript API Reference](https://github.com/OfficeDev/office-js-docs/blob/master/excel/excel-add-ins-javascript-reference.md)
5.  [Build your first Excel Add-in](https://github.com/OfficeDev/office-js-docs/blob/master/excel/build-your-first-excel-add-in.md)
