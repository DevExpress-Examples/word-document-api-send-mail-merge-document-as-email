Imports System
Imports System.Text
Imports DevExpress.XtraRichEdit
Imports DevExpress.Office.Services
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System.IO
Imports DevExpress.Utils
Imports DevExpress.Office.Utils
Imports System.Drawing.Imaging
Imports DevExpress.XtraRichEdit.Export

Namespace MailMergeEmail

    Public Class RichEditMailMessageExporter
        Implements IUriProvider

        Private ReadOnly server As RichEditDocumentServer

        Private ReadOnly mailItem As Outlook.MailItem

        Public Sub New(ByVal richServer As RichEditDocumentServer, ByVal mailItem As Outlook.MailItem)
            Guard.ArgumentNotNull(richServer, "control")
            Guard.ArgumentNotNull(mailItem, "mailItem")
            server = richServer
            Me.mailItem = mailItem
        End Sub

        Private tempFiles As String = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles")

        Public Overridable Sub Export()
            If Not Directory.Exists(tempFiles) Then Directory.CreateDirectory(tempFiles)
            AddHandler server.BeforeExport, AddressOf OnBeforeExport
            Dim htmlBody As String = server.Document.GetHtmlText(server.Document.Range, Me)
            RemoveHandler server.BeforeExport, AddressOf OnBeforeExport
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
            mailItem.HTMLBody = htmlBody
        End Sub

        Private Sub OnBeforeExport(ByVal sender As Object, ByVal e As BeforeExportEventArgs)
            Dim options As HtmlDocumentExporterOptions = TryCast(e.Options, HtmlDocumentExporterOptions)
            If options IsNot Nothing Then
                options.Encoding = Encoding.UTF8
            End If
        End Sub

        Private imageId As Integer

        Public Function CreateCssUri(ByVal rootUri As String, ByVal styleText As String, ByVal relativeUri As String) As String Implements IUriProvider.CreateCssUri
            Return String.Empty
        End Function

        Public Function CreateImageUri(ByVal rootUri As String, ByVal image As OfficeImage, ByVal relativeUri As String) As String Implements IUriProvider.CreateImageUri
            Dim imageName As String = String.Format("image{0}.png", imageId)
            imageId += 1
            Dim imagePath As String = Path.Combine(tempFiles, imageName)
            image.NativeImage.Save(imagePath, ImageFormat.Png)
            mailItem.Attachments.Add(imagePath, Outlook.OlAttachmentType.olByValue, 0, Type.Missing)
            Return "cid:" & imageName
        End Function
    End Class
End Namespace
