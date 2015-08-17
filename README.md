# Word-Add-in-JavaScript-ChangeContentWithXML

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
The sample shows how to use JavaScript to extract Open XML from a potentially complex document. The sample also shows how to insert a fragment of Open XML into a document. The sample comes with a test document ComplexDoc.docx, which is set as the StartAction property of the task pane plug-in. The document contains a mixture of images with various layout options and text. Make sure the document loads when you start debugging, and if not, check the StartAction property of the project.

**Note** The sample uses Open XML instead of HTML or plain text, because only Open XML is capable of handling the Base64 data that represents the images in the document. Also, Open XML is potentially more powerful at describing text flows with images.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires:

- Visual Studio 2012 or later
- Office 2013 tools for Visual Studio 2012 or later.
- Word 2013.

<a name="components"></a>
## Key components of the sample
The sample solution contains the following files:

The sample plug-in contains the ChangeContentWithXML project, which contains:

- The ChangeContentWithXML.xml manifest file.
- The ComplexDoc.docx document, which is prepopulated with various images, tables, and formatted textual content.

It also contains the ChangeContentWithXMLWeb project, which includes:

- Home.html. This contains the HTML user interface that is displayed in the task pane. It consists of two HTML buttons that extract and insert Open XML, a DIV where status messages will be written, and a **textarea** HTML control that is used to show you Open XML fragments.
- Home.js (in the Scripts folder). This script file contains code that runs when the plug-in is loaded. This startup wires up the Click event handlers for the two buttons in ChangeContentWithXML.html. One of these buttons retrieves the selected area of the document as Open XML, and the other button inserts Open XML into the document.

<a name="build"></a>
## Build and debug ##

1. Open the ChangeContentWithXML.sln file with Visual Studio. No other configuration is necessary.
2. To build the sample, choose the Ctrl+Shift+B keys.
3. To run the plug-in, choose the F5 key.

**Note** It is recommended that you select all the content between the two instructions as shown the first time you run the sample, so that you can see the full power of Open XML. You can experiment with selecting smaller sections after that.

<a name="troubleshooting"></a>
##Troubleshooting
If the plug-in starts with a blank document instead of the one shown in Figure 1, ensure the **StartAction** property of the project is set to ComplexDoc.docx and not just to Word.

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Word-Add-in-JavaScript-ChangeContentWithXML//issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [Build plug-ins for Office](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Open XML SDK 2.5 CTP for Office](http://msdn.microsoft.com/library/office/bb448854.aspx)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
