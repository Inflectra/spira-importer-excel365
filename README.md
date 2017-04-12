# SpiraTeam Office365

#### SpiraExcel is an add-in for Microsoft Excel that allows you to create requirements for a SpiraTeam project directly from Excel.

## Developers:
If you would like to set up a development environment and add to this project, be sure to run **npm install** and **npm install -g browser-sync**
after you have cloned the repository. This installs the necessary node modules and [browser sync](https://www.browsersync.io/) for easy live updates
in the browser as you save changes to your code. You can remove this dependancy if you want to, as it is not essential to any add-in functionality.

To host your code and test it first run **npm start** on your console. Once it is running copy the local or external URL and replace
the URLs in the [**manifest file**](https://github.com/Inflectra/spira_office365-excel/blob/master/spira-excel-exporter-manifest.xml). They are currently set to **ht<span></span>tps://localhost:3000**. You must replace **all** instances in **all** URLs (e.g. "http<span></span>s://localhost:3000/assets/icon-32.png" must become "https://{NEW-URL}/assets/icon-32.png").

The **manifest file** is what Excel uses to know what your add-in is called, how it should appear, and where the code is hosted. To open your version of the add-in, open up Excel and go to the **INSERT** tab. Then, click on **Office Add-Ins**. At the top right corner of the window that opens, you should see a drop down labeled **Manage My Add-Ins**. Open it and select **Upload My Add-In**. Find your manifest file, upload it, and you should see the Add-In in Excel.

### Important things to remember:
* **When interacting with Excel, your interaction in the code should look like this:**
	```javascript
    return Excel.run(function (context) {
        	//Do stuff here
        return context.sync()
        .then(function(){
        	//Handle returned data here
        })
        .catch(function(error){
        	//Handle errors here
        });
    ```
    If you only need to change cells in Excel without using the current value, this is enough. However, because this is a promise you will have to handle any data you pull from Excel in the **.then()** block.
 * **To access data pulled from cells in Excel, you must first load the data into your code with .load()**
 	
    For example, if you wanted to print the name of your active worksheet in the console you would load it in like this:
    ```javascript
    return Excel.run(function (context) {
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load();
        return context.sync()
        .then(function(){
        	console.log(sheet.name);
        });
    ```
	If you omit **sheet.load()** you will be told you don't have access in the console. If you try to console.log the sheet object's name before the **.then** block, you will get the same error.
* **If you were to load a new HTML page, you would need to reinitialize the Office code**
	
    The call for this is
    ```javascript
     Office.initialize = function (reason) {
     	$(document).ready(//main function);
        });
    ```

### OfficeJS Documentation

For Microsoft's **documentation** on officejs (for excel add-ins specifically), [click here](https://github.com/OfficeDev/office-js-docs/tree/master/reference/excel).

