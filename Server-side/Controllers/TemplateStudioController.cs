using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Mvc;

namespace DocumentTemplateStudioService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TemplateStudioController : ControllerBase
    {
        private readonly IWebHostEnvironment _environment;

        private readonly JsonSerializerOptions _jsonOptions =
            new()
            {
                PropertyNameCaseInsensitive = true,
                WriteIndented = true,
                DefaultIgnoreCondition =
                    JsonIgnoreCondition.WhenWritingNull
            };

        private readonly SemaphoreSlim _templatesLock = new(1, 1);
        private readonly SemaphoreSlim _commonFieldsLock = new(1, 1);

        public TemplateStudioController(
            IWebHostEnvironment environment)
        {
            _environment = environment;
        }

        #region Paths

        private string TemplatesRootPath =>
            Path.Combine(
                _environment.WebRootPath,
                "Templates");

        private string TemplatesJsonPath =>
            Path.Combine(
                TemplatesRootPath,
                "templates.json");

        private string CommonMergeFieldsJsonPath =>
            Path.Combine(
                TemplatesRootPath,
                "common-merge-fields.json");

        #endregion

        #region GET Templates

        [HttpGet("templates")]
        public async Task<IActionResult> GetTemplates(
            CancellationToken cancellationToken)
        {
            try
            {
                await EnsureStorageAsync(cancellationToken);

                var templates =
                    await ReadTemplatesAsync(
                        cancellationToken);

                return Ok(templates);
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message = "Failed to load templates.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region GET Template

        [HttpGet("templates/{id}")]
        public async Task<IActionResult> GetTemplate(
            string id,
            CancellationToken cancellationToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest(new
                    {
                        message = "Template ID is required."
                    });
                }

                await EnsureStorageAsync(cancellationToken);

                var templates =
                    await ReadTemplatesAsync(
                        cancellationToken);

                var template =
                    templates.FirstOrDefault(
                        x => string.Equals(
                            x.Id,
                            id,
                            StringComparison.OrdinalIgnoreCase));

                if (template == null)
                {
                    return NotFound(new
                    {
                        message = "Template not found."
                    });
                }

                return Ok(template);
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message = "Failed to load the template.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region POST Template

        // Creates a metadata entry for a template whose .docx has
        // already been written to wwwroot/Templates/ by the
        // DocumentEditorController.Save round-trip. The actual .docx
        // upload path is taken by the editor's documented Save call;
        // this endpoint just registers the catalog entry next to the
        // existing file. Accepts a JSON body that reflects the shape
        // the React app sends (`id`, `name`, `type`, `description`,
        // `fileName`, `docxUrl`, `thumbnailUrl`, `fieldKeys`). `id`
        // and `fileName` are optional — when missing, the server
        // generates a fresh GUID-style id and derives fileName from
        // the existing on-disk .docx (matched by docxUrl basename).
        [HttpPost("templates")]
        [RequestSizeLimit(50 * 1024 * 1024)]
        public async Task<IActionResult> CreateTemplate(
            [FromBody] CreateTemplateRequest? request,
            CancellationToken cancellationToken)
        {
            try
            {
                if (request == null ||
                    string.IsNullOrWhiteSpace(request.Name))
                {
                    return BadRequest(new
                    {
                        message = "Template name is required."
                    });
                }

                await EnsureStorageAsync(cancellationToken);

                await _templatesLock.WaitAsync(
                    cancellationToken);

                try
                {
                    var templates =
                        await ReadTemplatesAsync(
                            cancellationToken);

                    // Pull the .docx basename from docxUrl when the
                    // client isn't already sending an explicit fileName
                    // (the React upload flow derives both — we just
                    // double-check here so a missing/misnamed file
                    // doesn't silently register nothing on disk).
                    var resolvedFileName =
                        !string.IsNullOrWhiteSpace(request.FileName)
                            ? request.FileName.Trim()
                            : ExtractFileNameFromUrl(request.DocxUrl);

                    var now = DateTime.UtcNow;

                    // Re-derive id when missing so two uploads with the
                    // same `name` (no id) don't collapse each other.
                    var resolvedId = !string.IsNullOrWhiteSpace(request.Id)
                        ? request.Id.Trim()
                        : GenerateTemplateId();

                    var template =
                        new TemplateMetadata
                        {
                            Id = resolvedId,
                            Name = request.Name.Trim(),
                            Type =
                                string.IsNullOrWhiteSpace(request.Type)
                                    ? "General"
                                    : request.Type.Trim(),
                            Description =
                                string.IsNullOrWhiteSpace(
                                    request.Description)
                                    ? null
                                    : request.Description.Trim(),
                            // FileName drives the on-disk .docx; when
                            // empty we treat the entry as placeholder-
                            // only ("+ New Template" before the user
                            // ever pressed Save & Publish). The Round
                            // Trip — editor.Save -> /api/DocumentEditor/Save
                            // — fills the actual file using the
                            // docxUrl-derived basename.
                            FileName = resolvedFileName ?? string.Empty,
                            ThumbnailUrl = request.ThumbnailUrl,
                            DocxUrl = request.DocxUrl,
                            FieldKeys =
                                request.FieldKeys
                                ?? new List<string>(),
                            CreatedAt =
                                request.CreatedAt == default
                                    ? now
                                    : request.CreatedAt,
                            UpdatedAt =
                                request.UpdatedAt == default
                                    ? now
                                    : request.UpdatedAt
                        };

                    // Idempotency: if a template with the same id
                    // already exists, replace it (same Upload flow
                    // retries on a transient Save failure).
                    var existingIndex =
                        templates.FindIndex(
                            x => string.Equals(
                                x.Id,
                                template.Id,
                                StringComparison.OrdinalIgnoreCase));

                    if (existingIndex >= 0)
                    {
                        templates[existingIndex] = template;
                    }
                    else
                    {
                        templates.Add(template);
                    }

                    await WriteTemplatesAsync(
                        templates,
                        cancellationToken);

                    return CreatedAtAction(
                        nameof(GetTemplate),
                        new { id = template.Id },
                        template);
                }
                finally
                {
                    _templatesLock.Release();
                }
            }
            catch (OperationCanceledException)
            {
                return StatusCode(499);
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to create the template.",
                        detail = ex.Message
                    });
            }
        }

        private static string? ExtractFileNameFromUrl(string? docxUrl)
        {
            if (string.IsNullOrWhiteSpace(docxUrl))
            {
                return null;
            }

            try
            {
                // docxUrl is usually "/Templates/<slug>.docx" or
                // "http(s)://host[:port]/Templates/<slug>.docx".
                var uri = new Uri(
                    docxUrl,
                    UriKind.RelativeOrAbsolute);
                var rawLast = uri.IsAbsoluteUri
                    ? uri.AbsolutePath
                    : uri.OriginalString;
                var slashIndex = rawLast.LastIndexOf('/');
                var last = slashIndex >= 0
                    ? rawLast.Substring(slashIndex + 1)
                    : rawLast;
                return string.IsNullOrWhiteSpace(last) ? null : last;
            }
            catch
            {
                return null;
            }
        }

        #endregion

        #region PUT Template

        [HttpPut("templates/{id}")]
        public async Task<IActionResult> UpdateTemplate(
            string id,
            [FromBody] UpdateTemplateRequest request,
            CancellationToken cancellationToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest(new
                    {
                        message = "Template ID is required."
                    });
                }

                if (request == null)
                {
                    return BadRequest(new
                    {
                        message = "Request body is required."
                    });
                }

                await EnsureStorageAsync(
                    cancellationToken);

                await _templatesLock.WaitAsync(
                    cancellationToken);

                try
                {
                    var templates =
                        await ReadTemplatesAsync(
                            cancellationToken);

                    var template =
                        templates.FirstOrDefault(
                            x => string.Equals(
                                x.Id,
                                id,
                                StringComparison.OrdinalIgnoreCase));

                    if (template == null)
                    {
                        return NotFound(new
                        {
                            message = "Template not found."
                        });
                    }

                    if (request.Name != null)
                    {
                        if (string.IsNullOrWhiteSpace(
                                request.Name))
                        {
                            return BadRequest(new
                            {
                                message =
                                    "Template name cannot be empty."
                            });
                        }

                        template.Name =
                            request.Name.Trim();
                    }

                    if (request.Type != null)
                    {
                        template.Type =
                            request.Type.Trim();
                    }

                    if (request.Description != null)
                    {
                        template.Description =
                            request.Description.Trim();
                    }

                    template.UpdatedAt =
                        DateTime.UtcNow;

                    await WriteTemplatesAsync(
                        templates,
                        cancellationToken);

                    return Ok(template);
                }
                finally
                {
                    _templatesLock.Release();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to update the template.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region DELETE Template

        [HttpDelete("templates/{id}")]
        public async Task<IActionResult> DeleteTemplate(
            string id,
            CancellationToken cancellationToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest(new
                    {
                        message = "Template ID is required."
                    });
                }

                await EnsureStorageAsync(
                    cancellationToken);

                await _templatesLock.WaitAsync(
                    cancellationToken);

                try
                {
                    var templates =
                        await ReadTemplatesAsync(
                            cancellationToken);

                    var template =
                        templates.FirstOrDefault(
                            x => string.Equals(
                                x.Id,
                                id,
                                StringComparison.OrdinalIgnoreCase));

                    if (template == null)
                    {
                        return NotFound(new
                        {
                            message = "Template not found."
                        });
                    }

                    var filePath =
                        GetTemplateFilePath(
                            template.FileName);

                    if (System.IO.File.Exists(filePath))
                    {
                        System.IO.File.Delete(filePath);
                    }
                    else if (!string.IsNullOrWhiteSpace(
                                 template.DocxUrl))
                    {
                        // Catalog entries persisted from older clients
                        // may carry only `docxUrl` instead of an
                        // explicit FileName. Cover that case so the
                        // .docx on disk is cleaned up alongside the
                        // catalog row.
                        var fromUrl =
                            ExtractFileNameFromUrl(template.DocxUrl);

                        if (!string.IsNullOrWhiteSpace(fromUrl))
                        {
                            var pathFromUrl =
                                GetTemplateFilePath(fromUrl);

                            if (System.IO.File.Exists(pathFromUrl))
                            {
                                System.IO.File.Delete(pathFromUrl);
                            }
                        }
                    }

                    templates.Remove(template);

                    await WriteTemplatesAsync(
                        templates,
                        cancellationToken);

                    return NoContent();
                }
                finally
                {
                    _templatesLock.Release();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to delete the template.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region GET Common Merge Fields

        [HttpGet("merge-fields/common")]
        public async Task<IActionResult> GetCommonMergeFields(
            CancellationToken cancellationToken)
        {
            try
            {
                await EnsureStorageAsync(
                    cancellationToken);

                var fields =
                    await ReadCommonMergeFieldsAsync(
                        cancellationToken);

                return Ok(fields);
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to load common merge fields.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region POST Common Merge Field

        [HttpPost("merge-fields/common")]
        public async Task<IActionResult> AddCommonMergeField(
            [FromBody] MergeFieldRequest request,
            CancellationToken cancellationToken)
        {
            try
            {
                if (request == null ||
                    string.IsNullOrWhiteSpace(request.Key))
                {
                    return BadRequest(new
                    {
                        message =
                            "Merge field key is required."
                    });
                }

                var key = request.Key.Trim();

                await EnsureStorageAsync(
                    cancellationToken);

                await _commonFieldsLock.WaitAsync(
                    cancellationToken);

                try
                {
                    var fields =
                        await ReadCommonMergeFieldsAsync(
                            cancellationToken);

                    if (fields.Any(
                            x => string.Equals(
                                x,
                                key,
                                StringComparison.OrdinalIgnoreCase)))
                    {
                        return Conflict(new
                        {
                            message =
                                "The common merge field already exists."
                        });
                    }

                    fields.Add(key);

                    fields =
                        fields
                            .Distinct(
                                StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x)
                            .ToList();

                    await WriteCommonMergeFieldsAsync(
                        fields,
                        cancellationToken);

                    return Ok(new
                    {
                        key
                    });
                }
                finally
                {
                    _commonFieldsLock.Release();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to add common merge field.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region GET Template Merge Fields

        [HttpGet("templates/{id}/merge-fields")]
        public async Task<IActionResult> GetTemplateMergeFields(
            string id,
            CancellationToken cancellationToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest(new
                    {
                        message =
                            "Template ID is required."
                    });
                }

                await EnsureStorageAsync(
                    cancellationToken);

                var templates =
                    await ReadTemplatesAsync(
                        cancellationToken);

                var template =
                    templates.FirstOrDefault(
                        x => string.Equals(
                            x.Id,
                            id,
                            StringComparison.OrdinalIgnoreCase));

                if (template == null)
                {
                    return NotFound(new
                    {
                        message =
                            "Template not found."
                    });
                }

                return Ok(
                    template.FieldKeys ??
                    new List<string>());
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to load template merge fields.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region POST Template Merge Field

        [HttpPost("templates/{id}/merge-fields")]
        public async Task<IActionResult> AddTemplateMergeField(
            string id,
            [FromBody] MergeFieldRequest request,
            CancellationToken cancellationToken)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    return BadRequest(new
                    {
                        message =
                            "Template ID is required."
                    });
                }

                if (request == null ||
                    string.IsNullOrWhiteSpace(request.Key))
                {
                    return BadRequest(new
                    {
                        message =
                            "Merge field key is required."
                    });
                }

                var key = request.Key.Trim();

                await EnsureStorageAsync(
                    cancellationToken);

                await _templatesLock.WaitAsync(
                    cancellationToken);

                try
                {
                    var templates =
                        await ReadTemplatesAsync(
                            cancellationToken);

                    var template =
                        templates.FirstOrDefault(
                            x => string.Equals(
                                x.Id,
                                id,
                                StringComparison.OrdinalIgnoreCase));

                    if (template == null)
                    {
                        return NotFound(new
                        {
                            message =
                                "Template not found."
                        });
                    }

                    template.FieldKeys ??=
                        new List<string>();

                    if (template.FieldKeys.Any(
                            x => string.Equals(
                                x,
                                key,
                                StringComparison.OrdinalIgnoreCase)))
                    {
                        return Conflict(new
                        {
                            message =
                                "The merge field already exists for this template."
                        });
                    }

                    template.FieldKeys.Add(key);

                    template.FieldKeys =
                        template.FieldKeys
                            .Distinct(
                                StringComparer.OrdinalIgnoreCase)
                            .OrderBy(x => x)
                            .ToList();

                    template.UpdatedAt =
                        DateTime.UtcNow;

                    await WriteTemplatesAsync(
                        templates,
                        cancellationToken);

                    return Ok(new
                    {
                        key
                    });
                }
                finally
                {
                    _templatesLock.Release();
                }
            }
            catch (Exception ex)
            {
                return StatusCode(
                    StatusCodes.Status500InternalServerError,
                    new
                    {
                        message =
                            "Failed to add template merge field.",
                        detail = ex.Message
                    });
            }
        }

        #endregion

        #region Storage

        private async Task EnsureStorageAsync(
            CancellationToken cancellationToken)
        {
            Directory.CreateDirectory(
                TemplatesRootPath);

            if (!System.IO.File.Exists(
                    TemplatesJsonPath))
            {
                await System.IO.File.WriteAllTextAsync(
                    TemplatesJsonPath,
                    "[]",
                    cancellationToken);
            }

            if (!System.IO.File.Exists(
                    CommonMergeFieldsJsonPath))
            {
                await System.IO.File.WriteAllTextAsync(
                    CommonMergeFieldsJsonPath,
                    "[]",
                    cancellationToken);
            }
        }

        private async Task<List<TemplateMetadata>>
            ReadTemplatesAsync(
                CancellationToken cancellationToken)
        {
            if (!System.IO.File.Exists(
                    TemplatesJsonPath))
            {
                return new List<TemplateMetadata>();
            }

            await using var stream =
                new FileStream(
                    TemplatesJsonPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    81920,
                    useAsync: true);

            try
            {
                var templates =
                    await JsonSerializer.DeserializeAsync<
                        List<TemplateMetadata>>(
                        stream,
                        _jsonOptions,
                        cancellationToken);

                if (templates == null)
                {
                    return new List<TemplateMetadata>();
                }

                // Default timestamps for entries that pre-date the
                // schema, so the response shape is always complete and
                // JavaScript templates that assume CreatedAt/UpdatedAt
                // exist never blow up.
                var now = DateTime.UtcNow;

                foreach (var t in templates)
                {
                    if (t.CreatedAt == default)
                    {
                        t.CreatedAt = now;
                    }

                    if (t.UpdatedAt == default)
                    {
                        t.UpdatedAt = now;
                    }

                    t.FieldKeys ??= new List<string>();
                }

                return templates;
            }
            catch (JsonException)
            {
                // Catalog got corrupted (e.g. partial previous write).
                // Reset to an empty catalog rather than 500'ing the
                // dashboard; the operator can repopulate afterwards.
                await System.IO.File.WriteAllTextAsync(
                    TemplatesJsonPath,
                    "[]",
                    cancellationToken);

                return new List<TemplateMetadata>();
            }
        }

        private async Task WriteTemplatesAsync(
            List<TemplateMetadata> templates,
            CancellationToken cancellationToken)
        {
            var temporaryPath =
                TemplatesJsonPath + ".tmp";

            await using (var stream =
                         new FileStream(
                             temporaryPath,
                             FileMode.Create,
                             FileAccess.Write,
                             FileShare.None,
                             81920,
                             useAsync: true))
            {
                await JsonSerializer.SerializeAsync(
                    stream,
                    templates,
                    _jsonOptions,
                    cancellationToken);
            }

            System.IO.File.Move(
                temporaryPath,
                TemplatesJsonPath,
                overwrite: true);
        }

        private async Task<List<string>>
            ReadCommonMergeFieldsAsync(
                CancellationToken cancellationToken)
        {
            if (!System.IO.File.Exists(
                    CommonMergeFieldsJsonPath))
            {
                return new List<string>();
            }

            await using var stream =
                new FileStream(
                    CommonMergeFieldsJsonPath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read,
                    81920,
                    useAsync: true);

            // The on-disk file is a flat array of field names. Older
            // versions stored `{ fields: { ... }, updatedAt }` from the
            // client-side catalog. Be forgiving: if the payload is an
            // object with a `fields` property that's a flat array, lift
            // those values out so an old file doesn't 500 the catalog
            // load (which would otherwise wipe the templates list view).
            using var doc = await JsonDocument.ParseAsync(
                stream,
                cancellationToken: cancellationToken);

            if (doc.RootElement.ValueKind
                == JsonValueKind.Array)
            {
                var arr = new List<string>();
                foreach (var el in doc.RootElement.EnumerateArray())
                {
                    if (el.ValueKind == JsonValueKind.String)
                    {
                        arr.Add(el.GetString() ?? string.Empty);
                    }
                }

                return arr
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s)
                    .ToList();
            }

            if (doc.RootElement.ValueKind
                == JsonValueKind.Object &&
                doc.RootElement.TryGetProperty(
                    "fields",
                    out var fieldsElement) &&
                fieldsElement.ValueKind == JsonValueKind.Array)
            {
                var arr = new List<string>();
                foreach (var el in fieldsElement.EnumerateArray())
                {
                    if (el.ValueKind == JsonValueKind.String)
                    {
                        arr.Add(el.GetString() ?? string.Empty);
                    }
                }

                return arr
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s)
                    .ToList();
            }

            // Unknown shape — return an empty list so the catalog still
            // loads. Avoids wiping the templates view if a future
            // migration changes the file format again.
            await System.IO.File.WriteAllTextAsync(
                CommonMergeFieldsJsonPath,
                "[]",
                cancellationToken);

            return new List<string>();
        }

        private async Task WriteCommonMergeFieldsAsync(
            List<string> fields,
            CancellationToken cancellationToken)
        {
            var temporaryPath =
                CommonMergeFieldsJsonPath + ".tmp";

            await using (var stream =
                         new FileStream(
                             temporaryPath,
                             FileMode.Create,
                             FileAccess.Write,
                             FileShare.None,
                             81920,
                             useAsync: true))
            {
                await JsonSerializer.SerializeAsync(
                    stream,
                    fields,
                    _jsonOptions,
                    cancellationToken);
            }

            System.IO.File.Move(
                temporaryPath,
                CommonMergeFieldsJsonPath,
                overwrite: true);
        }

        #endregion

        #region Helpers

        private string GetTemplateFilePath(
            string fileName)
        {
            var safeFileName =
                Path.GetFileName(fileName);

            return Path.Combine(
                TemplatesRootPath,
                safeFileName);
        }

        private async Task<string>
            EnsureUniqueFileNameAsync(
                string fileName,
                CancellationToken cancellationToken)
        {
            var extension =
                Path.GetExtension(fileName);

            var name =
                Path.GetFileNameWithoutExtension(
                    fileName);

            var candidate = fileName;

            var counter = 1;

            while (System.IO.File.Exists(
                       GetTemplateFilePath(
                           candidate)))
            {
                candidate =
                    $"{name}-{counter}{extension}";

                counter++;

                await Task.Yield();

                cancellationToken.ThrowIfCancellationRequested();
            }

            return candidate;
        }

        private static string GenerateTemplateId()
        {
            return $"tpl-{Guid.NewGuid():N}";
        }

        private static string CreateSafeFileName(
            string originalFileName)
        {
            var fileName =
                Path.GetFileName(
                    originalFileName);

            foreach (var invalidCharacter
                     in Path.GetInvalidFileNameChars())
            {
                fileName =
                    fileName.Replace(
                        invalidCharacter,
                        '_');
            }

            return string.IsNullOrWhiteSpace(
                       fileName)
                ? "template.docx"
                : fileName;
        }

        #endregion

        #region Models

        public class UpdateTemplateRequest
        {
            public string? Name { get; set; }

            public string? Type { get; set; }

            public string? Description { get; set; }
        }

        // CreateTemplate: the React app POSTs a JSON body that mirrors
        // the in-memory template shape. FieldNames follow the camelCase
        // sent by the client (`id`, `name`, `type`, `description`,
        // `fileName`, `docxUrl`, `thumbnailUrl`, `fieldKeys`,
        // `createdAt`, `updatedAt`). `thumbnailUrl` is the SINGLE
        // canonical thumbnail field — both new and legacy clients must
        // use it; the older `thumbnail` field has been removed.
        public class CreateTemplateRequest
        {
            public string? Id { get; set; }

            public string Name { get; set; } = string.Empty;

            public string? Type { get; set; }

            public string? Description { get; set; }

            public string? FileName { get; set; }

            public string? DocxUrl { get; set; }

            public string? ThumbnailUrl { get; set; }

            public List<string>? FieldKeys { get; set; }

            public DateTime CreatedAt { get; set; }

            public DateTime UpdatedAt { get; set; }
        }

        public class MergeFieldRequest
        {
            public string Key { get; set; } = string.Empty;
        }

        public class TemplateMetadata
        {
            public string Id { get; set; } = string.Empty;

            public string Name { get; set; } = string.Empty;

            public string Type { get; set; } = "General";

            public string? Description { get; set; }

            // FileName is the on-disk .docx (basename + extension) under
            // wwwroot/Templates/. Required for the file-delete flow.
            public string FileName { get; set; } = string.Empty;

            // Thumbnail for the dashboard card — single canonical field
            // (`thumbnailUrl`) on every catalog entry. Stored as a base64
            // data URL so the dashboard renders without a server
            // round-trip / page reload; the upload + save flows refresh
            // it client-side from the live DocumentEditor.
            [System.Text.Json.Serialization.JsonPropertyName("thumbnailUrl")]
            public string? ThumbnailUrl { get; set; }

            // Absolute URL used by the editor's import path to fetch the
            // .docx on disk. Owned by the server (we always rewrite it
            // from FileName at create/update time so a legacy/missing
            // value never causes a stale URL to be served to the
            // editor).
            [System.Text.Json.Serialization.JsonPropertyName("docxUrl")]
            public string? DocxUrl { get; set; }

            [System.Text.Json.Serialization.JsonPropertyName("fieldKeys")]
            public List<string> FieldKeys { get; set; } = new();

            public DateTime CreatedAt { get; set; }

            public DateTime UpdatedAt { get; set; }
        }

        #endregion
    }
}