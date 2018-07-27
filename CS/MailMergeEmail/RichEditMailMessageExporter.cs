using System;
using System.Text;
using DevExpress.XtraRichEdit;
using DevExpress.Office.Services;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using DevExpress.Utils;
using DevExpress.Office.Utils;
using System.Drawing.Imaging;
using DevExpress.XtraRichEdit.Export;

namespace MailMergeEmail
{
    public class RichEditMailMessageExporter : IUriProvider
    {
        readonly RichEditDocumentServer server;
        readonly Outlook.MailItem mailItem;

        public RichEditMailMessageExporter(RichEditDocumentServer richServer, Outlook.MailItem mailItem)
        {
            Guard.ArgumentNotNull(richServer, "control");
            Guard.ArgumentNotNull(mailItem, "mailItem");

            this.server = richServer;
            this.mailItem = mailItem;
        }
        string tempFiles = Path.Combine(Directory.GetCurrentDirectory(), "TempFiles");

        public virtual void Export()
        {
            if (!Directory.Exists(tempFiles))
                Directory.CreateDirectory(tempFiles);

            server.BeforeExport += OnBeforeExport;
            string htmlBody = server.Document.GetHtmlText(server.Document.Range, this);
            server.BeforeExport -= OnBeforeExport;

            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mailItem.HTMLBody = htmlBody;
        }
        private void OnBeforeExport(object sender, BeforeExportEventArgs e)
        {
            HtmlDocumentExporterOptions options = e.Options as HtmlDocumentExporterOptions;
            if (options != null)
            {
                options.Encoding = Encoding.UTF8;
            }
        }
        int imageId;

        public string CreateCssUri(string rootUri, string styleText, string relativeUri)
        {
            return String.Empty;
        }

        public string CreateImageUri(string rootUri, OfficeImage image, string relativeUri)
        {
            string imageName = String.Format("image{0}.png", imageId);
            imageId++;

            string imagePath = Path.Combine(tempFiles, imageName);

            image.NativeImage.Save(imagePath, ImageFormat.Png);

            mailItem.Attachments.Add(imagePath, Outlook.OlAttachmentType.olByValue, 0, Type.Missing);

            return "cid:" + imageName;
        }
    }
}
