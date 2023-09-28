using DevExpress.Compression;
using DevExpress.Pdf;
using Microsoft.AspNetCore.Mvc;
using Swashbuckle.AspNetCore.Annotations;
using System.Net;

namespace DocumentProcessingWebAPI.Controllers {
    [ApiController]
    [Route("[controller]/[action]")]
    public class PdfController : ControllerBase {
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a File", typeof(FileContentResult))]
        public async Task<IActionResult> Merge(IFormFile[] files) {
            try {
                MemoryStream result = new MemoryStream();
                using (PdfDocumentProcessor processor = new PdfDocumentProcessor()) {
                    processor.CreateEmptyDocument(result);
                    foreach (var file in files) {
                        using (var stream = new MemoryStream()) {
                            await file.CopyToAsync(stream);
                            stream.Seek(0, SeekOrigin.Begin);
                            processor.AppendDocument(stream);
                        }
                    }
                }
                result.Seek(0, SeekOrigin.Begin);
                return File(result, "application/pdf", "result.pdf");
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a File", typeof(FileContentResult))]
        public async Task<IActionResult> Split(IFormFile file) {
            try {
                using (PdfDocumentProcessor processor = new PdfDocumentProcessor()) {
                    using (var stream = new MemoryStream()) {
                        await file.CopyToAsync(stream);
                        stream.Seek(0, SeekOrigin.Begin);
                        processor.LoadDocument(stream, true);
                    }

                    Stream zipStream = new MemoryStream();
                    using (ZipArchive archive = new ZipArchive()) {
                        for (int i = 0; i < processor.Document.Pages.Count; i++) {
                            using (PdfDocumentProcessor tempProcessor = new PdfDocumentProcessor()) {
                                tempProcessor.CreateEmptyDocument();
                                tempProcessor.Document.Pages.Add(processor.Document.Pages[i]);
                                Stream result = new MemoryStream();
                                tempProcessor.SaveDocument(result);
                                result.Seek(0, SeekOrigin.Begin);
                                archive.AddStream($"page{i + 1}.pdf", result);
                            }
                        }
                        archive.Save(zipStream);
                    }
                    zipStream.Seek(0, SeekOrigin.Begin);
                    return File(zipStream, "application/zip", "result.zip");
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a File", typeof(FileContentResult))]
        public async Task<IActionResult> Encrypt(IFormFile file, string password) {
            try {
                MemoryStream result = new MemoryStream();
                using (PdfDocumentProcessor processor = new PdfDocumentProcessor()) {
                    using (var stream = new MemoryStream()) {
                        await file.CopyToAsync(stream);
                        stream.Seek(0, SeekOrigin.Begin);
                        processor.LoadDocument(stream, true);
                    }
                    PdfEncryptionOptions encryptionOptions = new PdfEncryptionOptions();
                    encryptionOptions.UserPasswordString = password;
                    encryptionOptions.Algorithm = PdfEncryptionAlgorithm.AES256;
                    processor.SaveDocument(result, new PdfSaveOptions() { EncryptionOptions = encryptionOptions });
                    result.Seek(0, SeekOrigin.Begin);
                    return File(result, "application/pdf", "result.pdf");
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
    }
}