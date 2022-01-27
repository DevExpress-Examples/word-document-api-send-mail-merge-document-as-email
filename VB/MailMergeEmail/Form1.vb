Imports DevExpress.XtraEditors
Imports DevExpress.XtraRichEdit
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System

Namespace MailMergeEmail

    Public Partial Class Form1
        Inherits XtraForm

        Private server As RichEditDocumentServer

        Public Sub New()
            InitializeComponent()
            server = New RichEditDocumentServer()
            server.LoadDocument("MailMergeSimple.docx")
        End Sub

        Private Sub SendAnEmail()
            Dim application As Outlook.Application = New Outlook.Application()
            Dim mailItem As Outlook.MailItem = CType(application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
            Dim exporter As RichEditMailMessageExporter = New RichEditMailMessageExporter(server, mailItem)
            exporter.Export()
            mailItem.Display(False)
        End Sub

        Private Sub simpleButton1_Click(ByVal sender As Object, ByVal e As EventArgs)
            SendAnEmail()
        End Sub
    End Class
End Namespace
