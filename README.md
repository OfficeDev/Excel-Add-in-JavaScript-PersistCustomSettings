# Excel-Add-in-JavaScript-PersistCustomSettings

**Table of contents**

* [Summary](#summary)
* [Prerequisites](#prerequisites)
* [Key components of the sample](#components)
* [Build and debug](#build)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Additional resources](#additional-resources)

<a name="summary"></a>
##Summary
This sample plug-in for Office demonstrates how to save custom settings in a plug-in in Word 2013, Excel 2013, PowerPoint 2013, or Project Professional 2013. The plug-in stores data as key/value pairs, using the JavaScript API for Office property bag, browser cookies, web storage (**localStorage** and **sessionStorage**), or by storing the data in a hidden div in the document. The plug-in also demonstrates best practices for implementing multiple-page navigation in an plug-in for Office.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:

- Word 2013, Excel 2013, PowerPoint 2013, or Project Professional 2013. 
- Visual Studio 2012 or higher.
- Microsoft Edge, Internet Explorer 9 or higher.
- Basic familiarity with JavaScript, CSS, jQuery, and HTML5. 

<a name="components"></a>
## Key components of the sample
The Persist custom settings sample plug-in contains the following notable files:

The CodeSample_PersistCustomSettings project, including: 

- CodeSample_PersistCustomSettings.xml manifest
- CodeSample_PersistCustomSettings.js file 
- CodeSample_PersistCustomSettings.html file 
- StorageLibrary.js file 
- toast.js file 
- App.css file 

<a name="build"></a>
## Build and debug ##

1. Choose the F5 key in Visual Studio to build and deploy the plug-in.
2. Use the plug-in’s interface to save data as key/value pairs and to retrieve a stored value using its key. 

<a name="troubleshooting"></a>
##Troubleshooting
If the plug-in fails to install, ensure that the  **SourceLocation** element in the CodeSample_PersistCustomSettings.xml has the correct URL value for the **DefaultValue** attribute.

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings/issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Introduction to Web Storage ](http://msdn.microsoft.com/library/cc197062(VS.85).aspx)
- [Settings object ](http://msdn.microsoft.com/library/fp142179(v=office.15))

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
