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
This sample add-in for Office demonstrates how to save custom settings in a add-in in Excel 2013. The add-in stores data as key/value pairs, using the JavaScript API for Office property bag, browser cookies, web storage (**localStorage** and **sessionStorage**), or by storing the data in a hidden div in the document. The add-in also demonstrates best practices for implementing multiple-page navigation in an add-in for Office.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:

- Visual Studio 2013 with Update 5 or Visual Studio 2015.
- Excel 2013
- Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
- One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
 Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components of the sample
The Persist custom settings sample add-in contains the following notable files:

The CodeSample_PersistCustomSettings project, including: 

- CodeSample_PersistCustomSettings.xml manifest
- CodeSample_PersistCustomSettings.js file 
- CodeSample_PersistCustomSettings.html file 
- StorageLibrary.js file 
- toast.js file 
- App.css file 

<a name="build"></a>
## Build and debug ##

1. Choose the F5 key in Visual Studio to build and deploy the add-in.
2. Use the add-in’s interface to save data as key/value pairs and to retrieve a stored value using its key. 

<a name="troubleshooting"></a>
##Troubleshooting
If the add-in fails to install, ensure that the  **SourceLocation** element in the CodeSample_PersistCustomSettings.xml has the correct URL value for the **DefaultValue** attribute.

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
