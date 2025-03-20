using DevExpress.Spreadsheet;
using DevExpress.XtraPrinting;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Data;
using RichEditDocumentFormat = DevExpress.XtraRichEdit.DocumentFormat;
using RichEditEncryptionSettings = DevExpress.XtraRichEdit.API.Native.EncryptionSettings;
using SpreadsheetDocumentFormat = DevExpress.Spreadsheet.DocumentFormat;
using SpreadsheetEncryptionSettings = DevExpress.Spreadsheet.EncryptionSettings;

namespace DocumentProcessingWebAPI.BusinessObjects
{
    public static class RichEditHelper
    {
        public static Stream ConvertDocument(Stream inputStream, RichEditFormat format)
        {
            using (RichEditDocumentServer server = new RichEditDocumentServer())
            {
                server.LoadDocument(inputStream);
                return SaveDocument(server, format);
            }
        }

        public static Stream SaveDocument(RichEditDocumentServer server, RichEditFormat format, RichEditEncryptionSettings encryptionSettings = null)
        {
            MemoryStream resultStream = new MemoryStream();
            if (format == RichEditFormat.Pdf)
                server.ExportToPdf(resultStream);
            else
            {
                RichEditDocumentFormat documentFormat = new RichEditDocumentFormat((int)format);
                if (documentFormat == RichEditDocumentFormat.Html)
                    server.Options.Export.Html.EmbedImages = true;
                if (encryptionSettings == null)
                    server.SaveDocument(resultStream, documentFormat);
                else
                    server.SaveDocument(resultStream, documentFormat, encryptionSettings);
            }
            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }

        public static Stream MailMerge(RichEditDocumentServer server, RichEditFormat format, MailMergeOptions options)
        {
            MemoryStream resultStream = new MemoryStream();

            RichEditDocumentFormat documentFormat = new RichEditDocumentFormat((int)format);
            if (documentFormat == RichEditDocumentFormat.Html)
                server.Options.Export.Html.EmbedImages = true;
            server.MailMerge(options, resultStream, documentFormat);

            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }

        public static DataTable FillDataTable()
        {
            DataTable mailMergeDataTable = new DataTable();

            DataColumnCollection columns = mailMergeDataTable.Columns;
            columns.Add("EmployeeID", typeof(int));
            columns.Add("FirstName", typeof(string));
            columns.Add("LastName", typeof(string));
            columns.Add("Employees.Address", typeof(string));
            columns.Add("Employees.City", typeof(string));
            columns.Add("Employees.PostalCode", typeof(string));
            columns.Add("Title", typeof(string));
            columns.Add("Photo", typeof(object));
            columns.Add("ContactName", typeof(string));
            columns.Add("ContactTitle", typeof(string));
            columns.Add("CompanyName", typeof(string));
            columns.Add("Customers.Address", typeof(string));
            columns.Add("Customers.City", typeof(string));
            columns.Add("Customers.PostalCode", typeof(string));
            mailMergeDataTable.Rows.Clear();
            foreach (Employee employee in MailMergeData.Employees)
            {
                foreach (Customer customer in MailMergeData.Customers)
                {
                    mailMergeDataTable.Rows.Add(employee.EmployeeID,
                employee.FirstName,
                employee.LastName,
                employee.Address,
                employee.City,
                employee.PostalCode,
                employee.Title,
                null,
                customer.ContactName,
                customer.ContactTitle,
                customer.CompanyName,
                customer.Address,
                customer.City,
                customer.PostalCode
                );
                }
            }

            return mailMergeDataTable;
        }


        public static string GetContentType(RichEditFormat documentFormat)
        {
            switch (documentFormat)
            {
                case RichEditFormat.Doc:
                case RichEditFormat.Dot:
                    return "application/msword";
                case RichEditFormat.Docm:
                case RichEditFormat.Docx:
                case RichEditFormat.Dotm:
                case RichEditFormat.Dotx:
                    return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                case RichEditFormat.ePub:
                    return "application/epub+zip";
                case RichEditFormat.Mht:
                case RichEditFormat.Html:
                    return "text/html";
                case RichEditFormat.Odt:
                    return "application/vnd.oasis.opendocument.text";
                case RichEditFormat.Txt:
                    return "text/plain";
                case RichEditFormat.Rtf:
                    return "application/rtf";
                case RichEditFormat.Xml:
                    return "application/xml";
                case RichEditFormat.Pdf:
                    return "application/pdf";
                default: return string.Empty;
            }
        }

        public static Stream ExportToAccessiblePdf(RichEditDocumentServer server, AccessiblePdfFormat format)
        {
            MemoryStream resultStream = new MemoryStream();
            PdfExportOptions pdfExportOptions = SharedHelper.GetOptions(format);
            server.ExportToPdf(resultStream, pdfExportOptions);

            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }

        public static RichEditFormat GetFormat(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".rtf":
                    return RichEditFormat.Rtf;
                case ".doc":
                    return RichEditFormat.Doc;
                case ".docx":
                    return RichEditFormat.Docx;
                case ".txt":
                    return RichEditFormat.Txt;
                case ".mht":
                    return RichEditFormat.Mht;
                case ".odt":
                    return RichEditFormat.Odt;
                case ".xml":
                    return RichEditFormat.Xml;
                case ".epub":
                    return RichEditFormat.ePub;
                case ".html":
                case ".htm":
                    return RichEditFormat.Html;
                default:
                    throw new ArgumentException("Unsupported file format");
            }
        }


    }
    public static class SpreadsheetHelper
    {
        public static Stream SaveDocument(Workbook workbook, SpreadsheetFormat format, SpreadsheetEncryptionSettings encryptionSettings = null)
        {
            MemoryStream resultStream = new MemoryStream();
            if (format == SpreadsheetFormat.Pdf)
                workbook.ExportToPdf(resultStream);
            else
            {
                if (format == SpreadsheetFormat.Html)
                {
                    DevExpress.XtraSpreadsheet.Export.HtmlDocumentExporterOptions options = new DevExpress.XtraSpreadsheet.Export.HtmlDocumentExporterOptions();
                    options.EmbedImages = true;
                    options.AnchorImagesToPage = true;
                    workbook.ExportToHtml(resultStream, options);
                }
                else
                {
                    SpreadsheetDocumentFormat documentFormat = new SpreadsheetDocumentFormat((int)format);
                    workbook.SaveDocument(resultStream, documentFormat, encryptionSettings);
                }
            }
            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }

        public static Stream ExportToAccessiblePdf(Workbook workbook, AccessiblePdfFormat format)
        {
            MemoryStream resultStream = new MemoryStream();
            PdfExportOptions pdfExportOptions = SharedHelper.GetOptions(format);
            workbook.ExportToPdf(resultStream, pdfExportOptions);

            resultStream.Seek(0, SeekOrigin.Begin);
            return resultStream;
        }


        public static string GetContentType(SpreadsheetFormat documentFormat)
        {
            switch (documentFormat)
            {
                case SpreadsheetFormat.Xlsx:
                    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                case SpreadsheetFormat.Xlsm:
                    return "application/vnd.ms-excel.sheet.macroEnabled.12";
                case SpreadsheetFormat.Xlsb:
                    return "application/vnd.ms-excel.sheet.binary.macroEnabled.12";
                case SpreadsheetFormat.Xls:
                    return "application/vnd.ms-excel";
                case SpreadsheetFormat.Xltx:
                    return "application/vnd.openxmlformats-officedocument.spreadsheetml.template";
                case SpreadsheetFormat.Xltm:
                    return "application/vnd.ms-excel.template.macroEnabled.12";
                case SpreadsheetFormat.Xlt:
                    return "application/vnd.ms-excel";
                case SpreadsheetFormat.Xml:
                    return "text/xml";
                case SpreadsheetFormat.Csv:
                    return "text/comma-separated-values";
                case SpreadsheetFormat.Txt:
                    return "text/plain";
                case SpreadsheetFormat.Pdf:
                    return "application/pdf";
                case SpreadsheetFormat.Html:
                    return "text/html";
            }
            return string.Empty;
        }
        public static SpreadsheetFormat GetFormat(string fileName)
        {
            switch (Path.GetExtension(fileName))
            {
                case ".xlsx":
                    return SpreadsheetFormat.Xlsx;
                case ".xlsm":
                    return SpreadsheetFormat.Xlsm;
                case ".xlsb":
                    return SpreadsheetFormat.Xlsb;
                case ".xls":
                    return SpreadsheetFormat.Xls;
                case ".xltx":
                    return SpreadsheetFormat.Xltx;
                case ".xltm":
                    return SpreadsheetFormat.Xltm;
                case ".xlt":
                    return SpreadsheetFormat.Xlt;
                case ".csv":
                    return SpreadsheetFormat.Csv;
                default:
                    throw new ArgumentException("Unsupported file format");
            }
        }
    }

    public static class SharedHelper
    {
        public static PdfExportOptions GetOptions(AccessiblePdfFormat format)
        {
            PdfExportOptions options = new PdfExportOptions();
            switch (format)
            {
                case AccessiblePdfFormat.PdfUa:
                    options.PdfUACompatibility = PdfUACompatibility.PdfUA1;
                    return options;
                case AccessiblePdfFormat.PdfA1a:
                    options.PdfACompatibility = PdfACompatibility.PdfA1a;
                    return options;
                case AccessiblePdfFormat.PdfA1b:
                    options.PdfACompatibility = PdfACompatibility.PdfA1b;
                    return options;
                case AccessiblePdfFormat.PdfA2a:
                    options.PdfACompatibility = PdfACompatibility.PdfA2a;
                    return options;
                case AccessiblePdfFormat.PdfA2b:
                    options.PdfACompatibility = PdfACompatibility.PdfA2b;
                    return options;
                case AccessiblePdfFormat.PdfA3a:
                    options.PdfACompatibility = PdfACompatibility.PdfA3a;
                    return options;
                case AccessiblePdfFormat.PdfA3b:
                    options.PdfACompatibility = PdfACompatibility.PdfA3b;
                    return options;
            }
            return options;

        }

    }
}
