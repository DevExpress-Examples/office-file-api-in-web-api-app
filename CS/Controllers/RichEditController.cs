using DevExpress.Compression;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DocumentProcessingWebAPI.BusinessObjects;
using Microsoft.AspNetCore.Mvc;
using Swashbuckle.AspNetCore.Annotations;
using System.Net;

namespace DocumentProcessingWebAPI.Controllers {
    [ApiController]
    [Route("[controller]/[action]")]
    public class RichEditController : ControllerBase {
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Convert(IFormFile file, [FromQuery] RichEditFormat format) {
            try {
                using (var stream = new MemoryStream()) {
                    await file.CopyToAsync(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                    using (RichEditDocumentServer server = new RichEditDocumentServer()) {
                        server.LoadDocument(stream);
                        Stream result = RichEditHelper.SaveDocument(server, format);
                        string contentType = RichEditHelper.GetContentType(format);
                        return File(result, contentType, $"result.{format}");
                    }
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Merge(IFormFile[] files, [FromQuery] RichEditFormat format) {
            try {
                using (RichEditDocumentServer server = new RichEditDocumentServer()) {
                    using (var stream = new MemoryStream()) {
                        await files[0].CopyToAsync(stream);
                        stream.Seek(0, SeekOrigin.Begin);
                        server.LoadDocument(stream);
                    }
                    for (int i = 1; i < files.Length; i++) {
                        using (var stream = new MemoryStream()) {
                            await files[i].CopyToAsync(stream);
                            stream.Seek(0, SeekOrigin.Begin);
                            server.Document.AppendDocumentContent(stream);
                        }
                    }
                    Stream result = RichEditHelper.SaveDocument(server, format);
                    string contentType = RichEditHelper.GetContentType(format);
                    return File(result, contentType, $"result.{format}");
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Split(IFormFile file, [FromQuery] RichEditFormat format) {
            try {
                using (RichEditDocumentServer server = new RichEditDocumentServer()) {
                    using (var stream = new MemoryStream()) {
                        await file.CopyToAsync(stream);
                        stream.Seek(0, SeekOrigin.Begin);
                        server.LoadDocument(stream);
                    }
                    int pageCount = server.DocumentLayout.GetPageCount();
                    using (ZipArchive archive = new ZipArchive()) {
                        for (int i = 0; i < pageCount; i++) {
                            DevExpress.XtraRichEdit.API.Layout.LayoutPage layoutPage = server.DocumentLayout.GetPage(i);
                            DocumentRange mainBodyRange = server.Document.CreateRange(layoutPage.MainContentRange.Start, layoutPage.MainContentRange.Length);
                            using (RichEditDocumentServer tempServer = new RichEditDocumentServer()) {
                                tempServer.Document.AppendDocumentContent(mainBodyRange);
                                tempServer.Document.Delete(tempServer.Document.Paragraphs.First().Range); //Delete last empty paragraph
                                Stream result = RichEditHelper.SaveDocument(tempServer, format);
                                archive.AddStream($"page{i + 1}.{format}", result);
                            }
                        }
                        Stream zipStream = new MemoryStream();
                        archive.Save(zipStream);
                        zipStream.Seek(0, SeekOrigin.Begin);
                        return File(zipStream, "application/zip", "result.zip");
                    }
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Encrypt(IFormFile file, [FromQuery] string password) {
            try {
                using (var stream = new MemoryStream()) {
                    await file.CopyToAsync(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                    using (RichEditDocumentServer server = new RichEditDocumentServer()) {
                        server.LoadDocument(stream);
                        EncryptionSettings encryptionSettings = new EncryptionSettings();
                        encryptionSettings.Type = DevExpress.XtraRichEdit.API.Native.EncryptionType.Strong;
                        encryptionSettings.Password = password;
                        RichEditFormat format = RichEditHelper.GetFormat(file.FileName);
                        Stream result = RichEditHelper.SaveDocument(server, format, encryptionSettings);
                        string contentType = RichEditHelper.GetContentType(format);
                        return File(result, contentType, $"result.{format}");
                    }
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
    }
}