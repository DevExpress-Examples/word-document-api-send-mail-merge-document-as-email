<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/142556795/18.1.3%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T830554)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# How to Export the Mail-Merge Document into an Outlook Mail Item

This code example shows how to transfer the mail-merge document into Outlook using <a href="https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/welcome-to-the-outlook-primary-interop-assembly-reference">Outlook Interop API</a> and Word Processing File API.

The Outlook Interpop API prepares a mail item based on the RichEditDocumentServer content. Images are processed using a custom <a href="https://docs.devexpress.com/OfficeFileAPI/DevExpress.Office.Services.IUriProvider">IUriProvider Interface</a> implementor. Convert native images into Outlook mail item attachments. Refer to the following web articles to learn how to deal with the Outlook-related part of this solution:

<a href="http://social.msdn.microsoft.com/Forums/en-US/vsto/thread/6c063b27-7e8a-4963-ad5f-ce7e5ffb2c64/">How to embed image in html body in c# into outlook mail</a>

<a href="http://social.msdn.microsoft.com/Forums/pl/outlookdev/thread/17efe46b-18fe-450f-9f6e-d8bb116161d8">Attach stream data with Outlook mail client</a>

<a href="http://stackoverflow.com/questions/4312687/how-to-embed-images-in-email">How to embed images in email</a>
