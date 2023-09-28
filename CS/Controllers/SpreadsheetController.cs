using DevExpress.Compression;
using DevExpress.Spreadsheet;
using DocumentProcessingWebAPI.BusinessObjects;
using Microsoft.AspNetCore.Mvc;
using Swashbuckle.AspNetCore.Annotations;
using System.Net;

namespace DocumentProcessingWebAPI.Controllers {
    [ApiController]
    [Route("[controller]/[action]")]
    public class SpreadsheetController : ControllerBase {
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Convert(IFormFile file, [FromQuery] SpreadsheetFormat format) {
            try {
                using (var stream = new MemoryStream()) {
                    await file.CopyToAsync(stream);
                    stream.Seek(0, SeekOrigin.Begin);
                    using (Workbook workbook = new Workbook()) {
                        workbook.LoadDocument(stream);
                        Stream result = SpreadsheetHelper.SaveDocument(workbook, format);
                        string contentType = SpreadsheetHelper.GetContentType(format);
                        return File(result, contentType, $"result.{format}");
                    }
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Merge(IFormFile[] files, [FromQuery] SpreadsheetFormat format) {
            try {
                using (Workbook workbook = new Workbook()) {
                    using (var stream = new MemoryStream()) {
                        await files[0].CopyToAsync(stream);
                        stream.Seek(0, SeekOrigin.Begin);
                        workbook.LoadDocument(stream);
                    }
                    for (int i = 1; i < files.Length; i++) {
                        using (var stream = new MemoryStream()) {
                            await files[i].CopyToAsync(stream);
                            stream.Seek(0, SeekOrigin.Begin);
                            using (Workbook tempWorkbook = new Workbook()) {
                                tempWorkbook.LoadDocument(stream);
                                workbook.Append(tempWorkbook);
                            }
                        }
                    }
                    Stream result = SpreadsheetHelper.SaveDocument(workbook, format);
                    string contentType = SpreadsheetHelper.GetContentType(format);
                    return File(result, contentType, $"result.{format}");
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
        [HttpPost]
        [SwaggerResponse((int)HttpStatusCode.OK, "Download a file", typeof(FileContentResult))]
        public async Task<IActionResult> Split(IFormFile file, [FromQuery] SpreadsheetFormat format) {
            try {
                using (Workbook workbook = new Workbook()) {
                    using (var stream = new MemoryStream()) {
                        await file.CopyToAsync(stream);
                        stream.Seek(0, SeekOrigin.Begin);
                        workbook.LoadDocument(stream);
                    }
                    using (ZipArchive archive = new ZipArchive()) {
                        int pageCount = workbook.Worksheets.Count;
                        for (int i = 0; i < pageCount; i++) {
                            using (Workbook tempWorkbook = new Workbook()) {
                                tempWorkbook.Worksheets[0].CopyFrom(workbook.Worksheets[i]);
                                Stream result = SpreadsheetHelper.SaveDocument(tempWorkbook, format);
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
                    using (Workbook workbook = new Workbook()) {
                        workbook.LoadDocument(stream);
                        EncryptionSettings encryptionSettings = new EncryptionSettings();
                        encryptionSettings.Type = DevExpress.Spreadsheet.EncryptionType.Strong;
                        encryptionSettings.Password = password;
                        SpreadsheetFormat format = SpreadsheetHelper.GetFormat(file.FileName);
                        Stream result = SpreadsheetHelper.SaveDocument(workbook, format, encryptionSettings);
                        string contentType = SpreadsheetHelper.GetContentType(format);
                        return File(result, contentType, $"result.{format}");
                    }
                }
            } catch (Exception e) {
                return StatusCode(500, e.Message + Environment.NewLine + e.StackTrace);
            }
        }
    }
}