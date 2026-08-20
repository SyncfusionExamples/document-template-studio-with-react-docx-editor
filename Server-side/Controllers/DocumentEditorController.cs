using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIORenderer;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.SpellChecker;
using Syncfusion.Pdf;
using WDocument = Syncfusion.DocIO.DLS.WordDocument;
using WFormatType = Syncfusion.DocIO.FormatType;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace DocumentTemplateStudioService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumentEditorController : ControllerBase
    {
        private readonly string templatePath;

        // Security Constants
        private static readonly string[] ALLOWED_EXTENSIONS = { ".docx", ".doc", ".rtf", ".txt", ".xml", ".html", ".dotx", ".docm", ".dotm" };

        public DocumentEditorController()
        {
            templatePath = Path.Combine(Program.webRootPath, "Templates");
        }

        /// <summary>
        /// Validates file extension against whitelist
        /// </summary>
        private bool IsAllowedExtension(string fileName)
        {
            string ext = Path.GetExtension(fileName).ToLower();
            return ALLOWED_EXTENSIONS.Contains(ext);
        }

        // Handles importing a document file and converting it to JSON format
        // Handles importing a document file and converting it to JSON format
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Import")]
        public string? Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream1 = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream1);
            stream1.Position = 0;

            WordDocument document = WordDocument.Load(stream1, GetFormatType(type.ToLower()));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }

        // Representing parameters for clipboard operations
        public class CustomRestrictParameter
        {
            public string? passwordBase64 { get; set; }
            public string? saltBase64 { get; set; }
            public int spinCount { get; set; }
        }

        // Handles document editing restrictions
        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("RestrictEditing")]
        public string[]? RestrictEditing([FromBody] CustomRestrictParameter param)
        {
            if (param.passwordBase64 == "" && param.passwordBase64 == null)
                return null;
            return WordDocument.ComputeHash(param.passwordBase64, param.saltBase64, param.spinCount);
        }

        // Determines the document format based on file extension
        internal static FormatType GetFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                case ".docx":
                case ".docm":
                case ".dotm":
                    return FormatType.Docx;
                case ".dot":
                case ".doc":
                    return FormatType.Doc;
                case ".rtf":
                    return FormatType.Rtf;
                case ".txt":
                    return FormatType.Txt;
                case ".xml":
                    return FormatType.WordML;
                case ".html":
                    return FormatType.Html;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        // Determines the document format type specifically for Word formats
        internal static WFormatType GetWFormatType(string format)
        {
            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            switch (format.ToLower())
            {
                case ".dotx":
                    return WFormatType.Dotx;
                case ".docx":
                    return WFormatType.Docx;
                case ".docm":
                    return WFormatType.Docm;
                case ".dotm":
                    return WFormatType.Dotm;
                case ".dot":
                    return WFormatType.Dot;
                case ".doc":
                    return WFormatType.Doc;
                case ".rtf":
                    return WFormatType.Rtf;
                case ".html":
                    return WFormatType.Html;
                case ".txt":
                    return WFormatType.Txt;
                case ".xml":
                    return WFormatType.WordML;
                case ".odt":
                    return WFormatType.Odt;
                default:
                    throw new NotSupportedException("EJ2 DocumentEditor does not support this file format.");
            }
        }

        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
            public string Format { get; set; }
        }

         /// <summary>
        /// Validates file name for path traversal and special characters
        /// </summary>
        private bool IsValidFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return false;

            // Check for path traversal attempts
            if (fileName.Contains("..") || fileName.Contains("/") || fileName.Contains("\\"))
                return false;

            // Check for invalid filename characters
            return !Path.GetInvalidFileNameChars()
                .Any(c => fileName.Contains(c));
        }

        private string RetrieveFileType(string name)
        {
            int index = name.LastIndexOf('.');
            string format = index > -1 && index < name.Length - 1 ?
                name.Substring(index) : ".doc";
            return format;
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Export")]
        public FileStreamResult Export([FromBody] SaveParameter data)
        {
            string fileName = data.FileName;
            string format = RetrieveFileType(string.IsNullOrEmpty(data.Format) ? fileName : data.Format);
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = "Document1.docx";
            }
            // Validate filename
            if (!IsValidFileName(fileName))
            {
                throw new ArgumentException("Invalid filename format.");
            }

            WDocument document;
            if (format.ToLower() == ".pdf")
            {
                Stream stream = WordDocument.Save(data.Content, FormatType.Docx);
                document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
            }
            else
            {
                document = WordDocument.Save(data.Content);
            }

            return SaveDocument(document, format, fileName);
        }

        private FileStreamResult SaveDocument(WDocument document, string format, string fileName)
        {
            // Validate filename before saving
            if (!IsValidFileName(fileName))
            {
                throw new ArgumentException("Invalid filename format.");
            }

            Stream stream = new MemoryStream();
            string contentType = "";
            try
            {
                if (format.ToLower() == ".pdf")
                {
                    contentType = "application/pdf";
                    DocIORenderer render = new DocIORenderer();
                    PdfDocument pdfDocument = render.ConvertToPDF(document);
                    stream = new MemoryStream();
                    pdfDocument.Save(stream);
                    pdfDocument.Close();
                }
                else
                {
                    WFormatType type = GetWFormatType(format);
                    switch (type)
                    {
                        case WFormatType.Rtf:
                            contentType = "application/rtf";
                            break;
                        case WFormatType.WordML:
                            contentType = "application/xml";
                            break;
                        case WFormatType.Html:
                            contentType = "application/html";
                            break;
                        case WFormatType.Dotx:
                            contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.template";
                            break;
                        case WFormatType.Docx:
                            contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                            break;
                        case WFormatType.Doc:
                            contentType = "application/msword";
                            break;
                        case WFormatType.Dot:
                            contentType = "application/msword";
                            break;
                        case WFormatType.Odt:
                            contentType = "application/vnd.oasis.opendocument.text";
                            break;
                        case WFormatType.Markdown:
                            contentType = "text/markdown";
                            break;
                    }
                    document.Save(stream, type);
                }

                document.Close();

                stream.Position = 0;
                return new FileStreamResult(stream, contentType)
                {
                    FileDownloadName = fileName
                };
            }
            catch (Exception ex)
            {
                stream?.Dispose();
                throw new InvalidOperationException("Error saving documents: " + ex.Message, ex);
            }
        }

         [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("Save")]
        public void Save([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            string format = !string.IsNullOrWhiteSpace(data.Format)
                ? (data.Format.StartsWith(".") ? data.Format : "." + data.Format)
                : RetrieveFileType(name);

            if (string.IsNullOrEmpty(name))
            {
                name = "Document1" + format;
            }

            // Validate filename
            if (!IsValidFileName(name))
            {
                throw new ArgumentException("Invalid filename format.");
            }

            string existingExt = Path.GetExtension(name);
            if (string.IsNullOrEmpty(existingExt) ||
                !string.Equals(existingExt, format, StringComparison.OrdinalIgnoreCase))
            {
                name = name + format;
            }

            try
            {
                WDocument document = WordDocument.Save(data.Content);
                string savePath = Path.Combine(templatePath, Path.GetFileName(name));

                // Ensure templatePath is within allowed directory
                string fullPath = Path.GetFullPath(savePath);
                string allowedPath = Path.GetFullPath(templatePath);
                if (!fullPath.StartsWith(allowedPath, StringComparison.OrdinalIgnoreCase))
                {
                    throw new UnauthorizedAccessException("Access denied: Cannot save file outside allowed directory.");
                }

                // FileMode.Create truncates any existing file at savePath so
                // a second Save with the same FileName cleanly OVERWRITES
                // the prior .docx (requirement: "Replace existing file with
                // same name — do not save as new file").
                using (FileStream fileStream = new FileStream(savePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    document.Save(fileStream, GetWFormatType(format));
                }

                document.Close();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error saving document: " + ex.Message, ex);
            }
        }

        [AcceptVerbs("Post")]
        [HttpPost]
        [EnableCors("AllowAllOrigins")]
        [Route("MailMerge")]
        public string MailMerge([FromBody] ExportData exportData)
        {
            byte[] data;
            // Input validation
            if (exportData == null || string.IsNullOrEmpty(exportData.documentData))
            {
                throw new ArgumentException("Document data cannot be null or empty.");
            }

            try
            {
                string cleanBase64 = exportData.documentData.Contains(',') ? exportData.documentData.Split(',')[1] : exportData.documentData;
                data = Convert.FromBase64String(cleanBase64);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(data, 0, data.Length);
                    stream.Position = 0;

                    using (Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx))
                    {
                        document.MailMerge.RemoveEmptyGroup = true;
                        document.MailMerge.RemoveEmptyParagraphs = true;
                        document.MailMerge.ClearFields = true;
                        document.MailMerge.Execute(GetJsonData(exportData.mailMergeData));
                        document.Save(stream, Syncfusion.DocIO.FormatType.Docx);
                    }

                    stream.Position = 0;
                    Syncfusion.EJ2.DocumentEditor.WordDocument wordDocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, Syncfusion.EJ2.DocumentEditor.FormatType.Docx);
                    string sfdtText = Newtonsoft.Json.JsonConvert.SerializeObject(wordDocument);
                    wordDocument?.Dispose();
                    return sfdtText;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Error processing mail merge: " + ex.Message, ex);
            }
        }

        public class ExportData
        {
            public string fileName { get; set; }
            public string documentData { get; set; }
            public string mailMergeData { get; set; }
        }

        #region Helper methods for Mail Merge JSON Data
        /// <summary>
        /// Prepares the data table from JSON data for processing.
        /// </summary>
        private static List<object> GetJsonData(string mailMergeData)
        {
            //Reads the JSON object from JSON file.
            JObject jsonObject = JObject.Parse(mailMergeData);
            //Converts JSON object to Dictionary.
            IDictionary<string, object> data = GetData(jsonObject);
            return data.Values.First() as List<object>;
        }

        /// <summary>
        /// Gets data from JSON object.
        /// </summary>
        /// <param name="jsonObject">JSON object.</param>
        /// <returns>Dictionary of data.</returns>
        private static IDictionary<string, object> GetData(JObject jsonObject)
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();
            foreach (var item in jsonObject)
            {
                object keyValue = null;
                if (item.Value is JArray)
                    keyValue = GetData((JArray)item.Value);
                else if (item.Value is JToken)
                    keyValue = ((JToken)item.Value).ToObject<string>();
                dictionary.Add(item.Key, keyValue);
            }
            return dictionary;
        }
        /// <summary>
        /// Gets array of items from JSON array.
        /// </summary>
        /// <param name="jArray">JSON array.</param>
        /// <returns>List of objects.</returns>
        private static List<object> GetData(JArray jArray)
        {
            List<object> jArrayItems = new List<object>();
            foreach (var item in jArray)
            {
                object keyValue = null;
                if (item is JObject)
                    keyValue = GetData((JObject)item);
                jArrayItems.Add(keyValue);
            }
            return jArrayItems;
        }
        #endregion
    }
}
