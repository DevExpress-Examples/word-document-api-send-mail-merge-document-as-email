Imports DevExpress.XtraEditors
Imports DevExpress.XtraRichEdit
Imports Outlook = Microsoft.Office.Interop.Outlook
Imports System

Namespace MailMergeEmail
    Partial Public Class Form1
        Inherits XtraForm

        Private server As RichEditDocumentServer
        Public Sub New()
            InitializeComponent()
            server = New RichEditDocumentServer()
            server.LoadDocument("MailMergeSimple.docx")
        End Sub
        Private Sub SendAnEmail()
            Dim application As New Outlook.Application()
            Dim mailItem As Outlook.MailItem = CType(application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)

            Dim exporter As New RichEditMailMessageExporter(server, mailItem)
            exporter.Export()

            mailItem.Display(False)
        End Sub

        Private Sub simpleButton1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles simpleButton1.Click
            SendAnEmail()
        End Sub
    End Class
End Namespace
