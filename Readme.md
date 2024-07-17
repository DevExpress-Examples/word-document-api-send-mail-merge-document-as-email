<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/142556795/18.1.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830554)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
[![](https://img.shields.io/badge/💬_Leave_Feedback-feecdd?style=flat-square)](#does-this-example-address-your-development-requirementsobjectives)
<!-- default badges end -->
# Word Processing Document API - Send a Mail-Merge Document as an E-mail

This code example shows how to transfer the mail-merge document into Outlook using [Outlook Interop API](https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/welcome-to-the-outlook-primary-interop-assembly-reference) and [Word Processing File API](https://docs.devexpress.com/OfficeFileAPI/17488/word-processing-document-api).

# Implementation Details

The Outlook Interpop API prepares a mail item based on the [RichEditDocumentServer]() content. Images are processed using a custom [IUriProvider Interface](https://docs.devexpress.com/OfficeFileAPI/DevExpress.Office.Services.IUriProvider) implementor. Convert native images into Outlook mail item attachments. Refer to the following web articles to learn how to deal with the Outlook-related part of this solution:

* [How to embed image in HTML body in C# into Outlook mai](http://social.msdn.microsoft.com/Forums/en-US/vsto/thread/6c063b27-7e8a-4963-ad5f-ce7e5ffb2c64/)
* [Attach stream data with Outlook mail client](http://social.msdn.microsoft.com/Forums/pl/outlookdev/thread/17efe46b-18fe-450f-9f6e-d8bb116161d8)
* [How to embed images in email](http://stackoverflow.com/questions/4312687/how-to-embed-images-in-email)

# Files to Review

| C# | Visual Basic |
|---|---|
| [Form1.cs](./CS/MailMergeEmail/Form1.cs) | [Form1.vb](./VB/MailMergeEmail/Form1.vb) |
| [RichEditMailMessageExporter.cs](./CS/MailMergeEmail/RichEditMailMessageExporter.cs) | [RichEditMailMessageExporter.vb](./VB/MailMergeEmail/RichEditMailMessageExporter.vb) |

# Documentaton

* [Mail Merge in Word Processing Document API](https://docs.devexpress.com/OfficeFileAPI/15277/word-processing-document-api/mail-merge)
* [How to: Send the Document as an E-mail](https://docs.devexpress.com/OfficeFileAPI/120519/word-processing-document-api/examples/export/how-to-send-the-mail-merge-document-as-an-e-mail)
<!-- feedback -->
## Does this example address your development requirements/objectives?

[<img src="https://www.devexpress.com/support/examples/i/yes-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=word-document-api-send-mail-merge-document-as-email&~~~was_helpful=yes) [<img src="https://www.devexpress.com/support/examples/i/no-button.svg"/>](https://www.devexpress.com/support/examples/survey.xml?utm_source=github&utm_campaign=word-document-api-send-mail-merge-document-as-email&~~~was_helpful=no)

(you will be redirected to DevExpress.com to submit your response)
<!-- feedback end -->
