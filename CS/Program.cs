using DevExpress.XtraRichEdit;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailMergeEmail
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(string[] args)
        {
            RichEditDocumentServer server;
            server = new RichEditDocumentServer();
            server.LoadDocument("MailMergeSimple.docx");
            SendAnEmail(server);
        }
        private static void SendAnEmail(RichEditDocumentServer server)
        {
            Outlook.Application application = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);

            RichEditMailMessageExporter exporter = new RichEditMailMessageExporter(server, mailItem);
            exporter.Export();

            mailItem.Display(false);
        }

    }
}
