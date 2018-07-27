using DevExpress.XtraEditors;
using DevExpress.XtraRichEdit;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;

namespace MailMergeEmail
{
    public partial class Form1 : XtraForm
    {
        RichEditDocumentServer server;
        public Form1()
        {
            InitializeComponent();
            server = new RichEditDocumentServer();
            server.LoadDocument("MailMergeSimple.docx");
        }
        private void SendAnEmail()
        {
            Outlook.Application application = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);

            RichEditMailMessageExporter exporter = new RichEditMailMessageExporter(server, mailItem);
            exporter.Export();

            mailItem.Display(false);
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            SendAnEmail();
        }
    }
}
