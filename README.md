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
The sample Word Add-in shows how to use JavaScript to extract Open XML from a potentially complex document. The sample also shows how to insert a selected fragment of Open XML into a document.

The sample comes with a test document ComplexDoc.docx, which is set as the StartAction property of the task pane add-in. The document contains a mixture of images with various layout options and text. Make sure the document loads when you start debugging, and if not, check the StartAction property of the project.

**Note** The sample uses Open XML instead of HTML or plain text, because only Open XML is capable of handling the Base64 data that represents the images in the document. Also, Open XML is potentially more powerful at describing text flows with images.

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires:

- Visual Studio 2013 with Update 5 or Visual Studio 2015.
- Word 2013
- Internet Explorer 9 or later, which must be installed but doesn't have to be the default browser. To support Office Add-ins, the Office client that acts as host uses browser components that are part of Internet Explorer 9 or later.
- One of the following as the default browser: Internet Explorer 9, Safari 5.0.6, Firefox 5, Chrome 13, or a later version of one of these browsers.
- Familiarity with JavaScript programming and web services.

<a name="components"></a>
## Key components of the sample
The sample solution contains the following files:

The sample add-in contains the ChangeContentWithXML project, which contains:

- The ChangeContentWithXML.xml manifest file.
- The ComplexDoc.docx document, which is prepopulated with various images, tables, and formatted textual content.

It also contains the ChangeContentWithXMLWeb project, which includes:

- Home.html. This contains the HTML user interface that is displayed in the task pane. It consists of two HTML buttons that extract and insert Open XML, a DIV where status messages will be written, and a **textarea** HTML control that is used to show you Open XML fragments.
- Home.js (in the Scripts folder). This script file contains code that runs when the add-in is loaded. This startup wires up the Click event handlers for the two buttons in ChangeContentWithXML.html. One of these buttons retrieves the selected area of the document as Open XML, and the other button inserts Open XML into the document.

<a name="build"></a>
## Build and debug ##

1. Open the ChangeContentWithXML.sln file with Visual Studio. No other configuration is necessary.
2. To build the sample, choose the Ctrl+Shift+B keys.
3. To run the add-in, choose the F5 key.
4. On the **Home** tab, click the **Open** button in the **XML Content** group.

**Note** It is recommended that you select all the content between the two instructions as shown the first time you run the sample, so that you can see the full power of Open XML. You can experiment with selecting smaller sections after that.

<a name="troubleshooting"></a>
##Troubleshooting
If the add-in starts with a blank document instead of the test document, ensure the **Start Document** property of the project is set to ComplexDoc.docx and not just to [New Word document].

To do this, select the **ChangeContentWithXML** project in the Solution Explorer and view the properties in the Properties window. Under App you will see Start Action and Start Document listed.  The values for these should be:

- Start Action: Office Desktop Client
- Start Document: ComplexDoc.docx

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/Word-Add-in-JavaScript-ChangeContentWithXML//issues).
- Questions about Office Add-in development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Make sure that your questions or comments are tagged with [office-addins].


<a name="additional-resources"></a>
## Additional resources ##

- [More Add-in samples](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)
- [Build add-ins for Office](http://msdn.microsoft.com/library/office/jj220060.aspx)
- [Open XML SDK 2.5 CTP for Office](http://msdn.microsoft.com/library/office/bb448854.aspx)

## Copyright
Copyright (c) 2015 Microsoft. All rights reserved.
