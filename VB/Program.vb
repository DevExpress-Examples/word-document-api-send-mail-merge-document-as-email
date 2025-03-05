Imports System
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit
Imports Microsoft.Office.Interop
Imports Microsoft.SqlServer
Imports Outlook = Microsoft.Office.Interop.Outlook

Namespace MailMergeEmail

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        Private server As RichEditDocumentServer
        Sub Main()
            server = New RichEditDocumentServer()
            server.LoadDocument("MailMergeSimple.docx")
            SendAnEmail(server)
        End Sub
        Private Sub SendAnEmail(server As RichEditDocumentServer)
            Dim application As Outlook.Application = New Outlook.Application()
            Dim mailItem As Outlook.MailItem = CType(application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
            Dim exporter As RichEditMailMessageExporter = New RichEditMailMessageExporter(server, mailItem)
            exporter.Export()
            mailItem.Display(False)
        End Sub
    End Module
End Namespace
