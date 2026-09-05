using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Scribble.Security;

namespace Scribble.Outlook
{
    public sealed class EmailAttachmentContent
    {
        public EmailAttachmentContent(
            string fileName,
            string kind,
            string text,
            string imageDataUrl = null)
        {
            FileName = TextBoundary.SingleLine(
                fileName ?? string.Empty,
                180);
            Kind = TextBoundary.SingleLine(
                kind ?? string.Empty,
                32);
            var raw = text ?? string.Empty;
            var characterLimit = EmailAttachmentReader.CurrentCharacterLimit;
            var bounded = TextBoundary.PlainText(
                raw,
                characterLimit);
            Truncated = raw.Length > characterLimit;
            Text = Truncated
                ? bounded +
                  "\n[Truncated: more content follows beyond the " +
                  "first " + bounded.Length.ToString() +
                  " extracted characters.]"
                : bounded;
            // A truncated data URL is corrupt base64, so an oversized
            // image is dropped rather than bounded.
            var boundedDataUrl = (imageDataUrl ?? string.Empty).Trim();
            ImageDataUrl = boundedDataUrl.Length <=
                EmailAttachmentReader.MaxImageDataUrlCharacters
                    ? boundedDataUrl
                    : string.Empty;
        }

        public string FileName { get; }

        public string Kind { get; }

        public string Text { get; }

        public bool Truncated { get; }

        public string ImageDataUrl { get; }
    }

    public sealed class LocalAttachmentLoadResult
    {
        public LocalAttachmentLoadResult(
            string path,
            EmailAttachmentContent content,
            string thumbnail)
        {
            Path = path ?? string.Empty;
            Content = content;
            Thumbnail = thumbnail ?? string.Empty;
        }

        public string Path { get; }

        public EmailAttachmentContent Content { get; }

        public string Thumbnail { get; }
    }

    public static class EmailAttachmentReader
    {
        public const int MaxAttachments = 10;
        // Uniform 100 MiB intake: extraction output stays bounded in
        // characters, so large files cost read time, not context.
        public const int MaxBytesPerAttachment =
            (int)AttachmentIntakePolicy.MaxFileBytes;
        public const int MaxImageBytesPerAttachment =
            (int)AttachmentIntakePolicy.MaxFileBytes;
        public const long MaxOperationBytes =
            AttachmentIntakePolicy.MaxOperationBytes;
        public const int MaxCharactersPerAttachment = 48000;
        public const int MaxTotalCharacters = 120000;
        // The data URL cap must stay comfortably above the encoded
        // image byte cap (800 KB is ~1.1M base64 characters).
        public const int MaxImageDataUrlCharacters = 2200000;

        // Text budgets scale with the provider (see ContextScale);
        // byte and count caps do not.
        [ThreadStatic] private static int? _pageExtractionLimit;
        internal static int CurrentCharacterLimit { get { return _pageExtractionLimit ?? ContextScale.Scaled(MaxCharactersPerAttachment); } }
        internal static int? PageCharacterLimit { get { return _pageExtractionLimit; } }

        public static MailboxAttachmentPage LoadLocalPage(string path, int offset, int count, CancellationToken token)
        {
            if (offset < 0 || count < 1 || count > 12000 || offset > int.MaxValue - count - 1024)
                throw new ArgumentOutOfRangeException();
            var previous = _pageExtractionLimit;
            try
            {
                _pageExtractionLimit = offset + count + 1024;
                var content = LoadLocalFile(path, token);
                if (content == null || content.Kind == "unreadable" || content.Kind == "limit" || content.Kind == "resource-limited" ||
                    content.Text.Contains("No machine-readable text") || content.Text.Contains("resource limit"))
                    throw new InvalidOperationException("The attachment could not be fully extracted. Supply a readable export; it has not been counted as reviewed.");
                var text = content.Text ?? "";
                if (offset > text.Length) throw new InvalidOperationException("Attachment content changed or the offset is invalid.");
                var length = Math.Min(count, text.Length - offset);
                return new MailboxAttachmentPage { FileName = content.FileName, Kind = content.Kind,
                    Text = text.Substring(offset, length), ImageDataUrl = content.ImageDataUrl, Offset = offset,
                    NextOffset = offset + length < text.Length || content.Truncated ? (int?)(offset + length) : null };
            }
            finally { _pageExtractionLimit = previous; }
        }

        private static int ScaledCharactersPerAttachment
        {
            get
            {
                return CurrentCharacterLimit;
            }
        }

        private static int ScaledTotalCharacters
        {
            get
            {
                return ContextScale.Scaled(MaxTotalCharacters);
            }
        }
        // Inline images at or under this size are treated as signature
        // graphics (logos, banners) and skipped; pasted screenshots and
        // photos are far larger and always kept.
        public const int SignatureImageMaxBytes = 64 * 1024;
        // Upload size drives vision latency: models tile images at
        // around 768px internally, so 1280px keeps screenshot text
        // legible while cutting the per-image payload several-fold.
        private const int MaxImageBytesForBase64 = 800 * 1024;
        public const long MaxImagePixels = 64L * 1000 * 1000;
        public const int MaxImageDimension = 32768;
        private const int MaxArchiveEntries = 10000;
        private const long MaxArchivePartBytes = 32L * 1024 * 1024;

        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(
                new[]
                {
                    ".png", ".jpg", ".jpeg", ".gif",
                    ".bmp", ".webp", ".tif", ".tiff"
                },
                StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> ExcelExtensions =
            new HashSet<string>(
                new[]
                {
                    ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm",
                    ".xls", ".csv", ".tsv"
                },
                StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> DocumentExtensions =
            new HashSet<string>(
                new[]
                {
                    ".pdf", ".pptx", ".pptm", ".ppsx", ".ppsm",
                    ".potx", ".docx", ".docm", ".dotx", ".dotm",
                    ".ppt", ".doc", ".rtf", ".odt", ".ods",
                    ".odp", ".msg", ".oft"
                },
                StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> TextExtensions =
            new HashSet<string>(
                new[]
                {
                    ".txt", ".md", ".log", ".json",
                    ".xml", ".html", ".htm", ".eml",
                    ".yaml", ".yml", ".ini"
                },
                StringComparer.OrdinalIgnoreCase);

        public static IReadOnlyList<EmailAttachmentContent> Read(
            object outlookApplication,
            MessageSnapshot message)
        {
            return Read(
                outlookApplication,
                message,
                CancellationToken.None,
                new AttachmentReadBudget());
        }

        public static IReadOnlyList<EmailAttachmentContent> Read(
            object outlookApplication,
            MessageSnapshot message,
            CancellationToken cancellationToken)
        {
            return Read(
                outlookApplication,
                message,
                cancellationToken,
                new AttachmentReadBudget());
        }

        public static IReadOnlyList<EmailAttachmentContent> Read(
            object outlookApplication,
            MessageSnapshot message,
            CancellationToken cancellationToken,
            AttachmentReadBudget sourceBudget)
        {
            if (outlookApplication == null)
            {
                throw new ArgumentNullException(nameof(outlookApplication));
            }

            if (message == null ||
                message.EntryId.Length == 0)
            {
                return new EmailAttachmentContent[0];
            }

            object session = null;
            object item = null;
            object attachments = null;
            try
            {
                dynamic application = outlookApplication;
                session = application.Session;
                dynamic outlookSession = session;
                try
                {
                    item = message.StoreId.Length > 0
                        ? outlookSession.GetItemFromID(
                            message.EntryId,
                            message.StoreId)
                        : outlookSession.GetItemFromID(message.EntryId);
                }
                catch
                {
                    return new EmailAttachmentContent[0];
                }

                dynamic mail = item;
                attachments = mail.Attachments;
                if (attachments == null)
                {
                    return new EmailAttachmentContent[0];
                }

                dynamic outlookAttachments = attachments;
                var count = Math.Min(
                    Convert.ToInt32(outlookAttachments.Count),
                    MaxAttachments);
                var results = new List<EmailAttachmentContent>(count);
                if (sourceBudget == null)
                {
                    sourceBudget = new AttachmentReadBudget();
                }
                var totalCharacters = 0;
                var signatureImagesSkipped = 0;
                for (var index = 1;
                     index <= count &&
                     totalCharacters < ScaledTotalCharacters;
                     index++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    object attachment = null;
                    string tempPath = null;
                    try
                    {
                        attachment = outlookAttachments.Item(index);
                        dynamic outlookAttachment = attachment;
                        var fileName = SafeString(
                            () => outlookAttachment.FileName);
                        if (fileName.Length == 0)
                        {
                            continue;
                        }

                        var extension = Path.GetExtension(fileName);
                        var reportedSize = SafeLong(
                            () => outlookAttachment.Size);
                        if (IsLikelySignatureImage(
                            attachment,
                            extension,
                            reportedSize))
                        {
                            signatureImagesSkipped++;
                            continue;
                        }

                        var reportedWarning =
                            AttachmentIntakePolicy.ValidateFile(
                                reportedSize);
                        if (reportedWarning.Length > 0 &&
                            reportedSize >
                                AttachmentIntakePolicy.MaxFileBytes)
                        {
                            results.Add(LimitContent(
                                fileName,
                                reportedSize,
                                reportedWarning));
                            totalCharacters += 80;
                            continue;
                        }

                        if (reportedSize > 0 &&
                            reportedSize > sourceBudget.RemainingBytes)
                        {
                            results.Add(LimitContent(
                                fileName,
                                reportedSize,
                                "Over the 250 MB attachment budget " +
                                "for this operation - content not read."));
                            totalCharacters += 80;
                            continue;
                        }

                        // Every attachment is saved and attempted;
                        // unknown extensions are identified by content
                        // and unreadable ones produce a visible note.
                        var safeExtension =
                            System.Text.RegularExpressions.Regex
                                .IsMatch(
                                    extension,
                                    "^\\.[A-Za-z0-9]{1,10}$")
                                ? extension
                                : ".bin";
                        tempPath = Path.Combine(
                            Path.GetTempPath(),
                            "Scribble-" +
                            Guid.NewGuid().ToString("N") +
                            safeExtension);
                        outlookAttachment.SaveAsFile(tempPath);
                        cancellationToken.ThrowIfCancellationRequested();

                        var fileInfo = new FileInfo(tempPath);
                        if (!fileInfo.Exists)
                        {
                            continue;
                        }

                        string sourceWarning;
                        if (!sourceBudget.TryReserve(
                                fileInfo.Length,
                                out sourceWarning))
                        {
                            results.Add(LimitContent(
                                fileName,
                                fileInfo.Length,
                                sourceWarning));
                            totalCharacters += 80;
                            continue;
                        }

                        var extracted = ExtractContent(
                            tempPath,
                            fileName,
                            extension,
                            cancellationToken);
                        if (extracted == null ||
                            extracted.Text.Length == 0)
                        {
                            results.Add(new EmailAttachmentContent(
                                fileName,
                                "unreadable",
                                "[Attachment: " + fileName + ", " +
                                fileInfo.Length.ToString() +
                                " bytes. This file type could not " +
                                "be converted to text or image " +
                                "input.]"));
                            totalCharacters += 80;
                            continue;
                        }

                        var remaining = ScaledTotalCharacters -
                            totalCharacters;
                        if (remaining <= 0)
                        {
                            break;
                        }

                        // ExtractContent already bounded the text to
                        // the per-attachment cap and appended a
                        // truncation notice when it overflowed; only
                        // the shared per-message budget is applied
                        // here so that notice survives intact.
                        var entry = extracted;
                        if (entry.Text.Length > remaining)
                        {
                            var clipped = TextBoundary.PlainText(
                                entry.Text,
                                remaining);
                            if (clipped.Length == 0)
                            {
                                break;
                            }

                            entry = new EmailAttachmentContent(
                                fileName,
                                entry.Kind,
                                clipped +
                                "\n[Truncated: the shared " +
                                "attachment budget for this " +
                                "message was reached.]",
                                entry.ImageDataUrl);
                        }

                        results.Add(entry);
                        totalCharacters += entry.Text.Length;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (tempPath != null)
                        {
                            TryDelete(tempPath);
                        }

                        Release(attachment);
                    }
                }

                if (signatureImagesSkipped > 0)
                {
                    results.Add(new EmailAttachmentContent(
                        "signature-images",
                        "note",
                        "[" + signatureImagesSkipped.ToString() +
                        " small inline image" +
                        (signatureImagesSkipped == 1 ? "" : "s") +
                        " ignored as signature graphics.]"));
                }

                return results;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return new EmailAttachmentContent[0];
            }
            finally
            {
                Release(attachments);
                Release(item);
                Release(session);
            }
        }

        internal static bool IsLikelySignatureImage(
            object attachment,
            string extension,
            long sizeBytes)
        {
            if (!ImageExtensions.Contains(extension) ||
                sizeBytes <= 0 ||
                sizeBytes > SignatureImageMaxBytes)
            {
                return false;
            }

            return IsInlineAttachment(attachment);
        }

        private static bool IsInlineAttachment(object attachment)
        {
            object accessor = null;
            try
            {
                dynamic outlookAttachment = attachment;
                accessor = outlookAttachment.PropertyAccessor;
                if (accessor == null)
                {
                    return false;
                }

                dynamic propertyAccessor = accessor;
                try
                {
                    // PR_ATTACHMENT_HIDDEN
                    var hidden = propertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B");
                    if (hidden is bool && (bool)hidden)
                    {
                        return true;
                    }
                }
                catch
                {
                }

                try
                {
                    // PR_ATTACH_CONTENT_ID marks cid-referenced
                    // inline body images.
                    var contentId = Convert.ToString(
                        propertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x3712001F"));
                    return !string.IsNullOrEmpty(contentId);
                }
                catch
                {
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                Release(accessor);
            }
        }

        private static long SafeLong(Func<object> reader)
        {
            try
            {
                return Convert.ToInt64(reader());
            }
            catch
            {
                return 0;
            }
        }

        // Loads a user-chosen local file through the same bounded
        // extraction pipeline as email attachments (documents become
        // text, images become vision input). User-initiated only.
        public static EmailAttachmentContent LoadLocalFile(string path)
        {
            return LoadLocalFile(path, CancellationToken.None);
        }

        public static EmailAttachmentContent LoadLocalFile(
            string path,
            CancellationToken cancellationToken)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                var fileName = Path.GetFileName(path ?? string.Empty);
                if (fileName.Length == 0)
                {
                    return null;
                }

                var info = new FileInfo(path);
                if (!info.Exists)
                {
                    return null;
                }

                var extension = Path.GetExtension(fileName);
                var sizeWarning =
                    AttachmentIntakePolicy.ValidateFile(info.Length);
                if (sizeWarning.Length > 0)
                {
                    return LimitContent(
                        fileName,
                        info.Length,
                        sizeWarning);
                }

                var extracted = ExtractContent(
                    path,
                    fileName,
                    extension,
                    cancellationToken);
                return extracted ?? new EmailAttachmentContent(
                    fileName,
                    "unreadable",
                    "[File: " + fileName +
                    ". This file type could not be converted to " +
                    "text or image input.]");
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
                return null;
            }
        }

        public static IReadOnlyList<LocalAttachmentLoadResult>
            LoadLocalFiles(
                IEnumerable<string> paths,
                CancellationToken cancellationToken,
                Action<int, int, string> progress = null)
        {
            var selected = (paths ?? new string[0])
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .ToArray();
            var results = new List<LocalAttachmentLoadResult>(
                selected.Length);
            var budget = new AttachmentReadBudget();
            for (var index = 0; index < selected.Length; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var path = selected[index];
                var fileName = Path.GetFileName(path);
                if (progress != null)
                {
                    progress(index + 1, selected.Length, fileName);
                }

                EmailAttachmentContent content;
                try
                {
                    var info = new FileInfo(path);
                    if (!info.Exists)
                    {
                        continue;
                    }

                    string warning;
                    if (!budget.TryReserve(info.Length, out warning))
                    {
                        content = LimitContent(
                            fileName,
                            info.Length,
                            warning);
                    }
                    else
                    {
                        content = LoadLocalFile(
                            path,
                            cancellationToken);
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch
                {
                    content = new EmailAttachmentContent(
                        fileName,
                        "unreadable",
                        "[File: " + fileName +
                        ". The file could not be read.]");
                }

                if (content == null)
                {
                    continue;
                }

                var thumbnail = content.ImageDataUrl.Length > 0
                    ? BuildThumbnailDataUrl(path)
                    : string.Empty;
                results.Add(new LocalAttachmentLoadResult(
                    path,
                    content,
                    thumbnail));
            }

            return results;
        }

        public static EmailAttachmentContent LimitContent(
            string fileName,
            long sizeBytes,
            string warning)
        {
            return new EmailAttachmentContent(
                fileName,
                "resource-limited",
                "[Attachment: " + fileName + ", " +
                Math.Max(0L, sizeBytes).ToString() +
                " bytes. " +
                TextBoundary.SingleLine(
                    warning ?? "Attachment resource limit reached.",
                    240) + "]");
        }

        // Small JPEG preview for attachment tray thumbnails.
        public static string BuildThumbnailDataUrl(string path)
        {
            try
            {
                using (var original =
                    System.Drawing.Image.FromFile(path))
                {
                    if (!IsSafeImageDimensions(
                            original.Width,
                            original.Height))
                    {
                        return null;
                    }

                    var longSide = Math.Max(
                        original.Width,
                        original.Height);
                    var scale = longSide > 96
                        ? 96.0 / longSide
                        : 1.0;
                    var width = Math.Max(
                        1,
                        (int)Math.Round(original.Width * scale));
                    var height = Math.Max(
                        1,
                        (int)Math.Round(original.Height * scale));
                    using (var bitmap =
                        new System.Drawing.Bitmap(width, height))
                    {
                        using (var graphics =
                            System.Drawing.Graphics.FromImage(
                                bitmap))
                        {
                            graphics.Clear(
                                System.Drawing.Color.White);
                            graphics.DrawImage(
                                original,
                                0,
                                0,
                                width,
                                height);
                        }

                        var encoded = EncodeJpeg(bitmap, 70);
                        return encoded == null
                            ? null
                            : "data:image/jpeg;base64," +
                              Convert.ToBase64String(encoded);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public static bool IsSupportedExtension(string extension)
        {
            return ImageExtensions.Contains(extension) ||
                   ExcelExtensions.Contains(extension) ||
                   DocumentExtensions.Contains(extension) ||
                   TextExtensions.Contains(extension);
        }

        public static bool IsImageExtension(string extension)
        {
            return ImageExtensions.Contains(extension);
        }

        private static EmailAttachmentContent ExtractContent(
            string path,
            string fileName,
            string extension,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ImageExtensions.Contains(extension))
            {
                return ExtractImage(
                    path,
                    fileName,
                    ImageMimeType(extension),
                    cancellationToken);
            }

            if (ExcelExtensions.Contains(extension))
            {
                return ExtractSpreadsheet(
                    path,
                    fileName,
                    extension,
                    cancellationToken);
            }

            if (DocumentExtensions.Contains(extension))
            {
                return ExtractDocument(
                    path,
                    fileName,
                    extension,
                    cancellationToken);
            }

            if (TextExtensions.Contains(extension))
            {
                return new EmailAttachmentContent(
                    fileName,
                    "text",
                    ReadTextFile(path, cancellationToken));
            }

            return SniffUnknownContent(
                path,
                fileName,
                cancellationToken);
        }

        // Unknown or missing extensions are identified by content so
        // every attachment is at least attempted: image magic bytes,
        // OOXML zip parts, OLE compound streams, then plain text.
        private static EmailAttachmentContent SniffUnknownContent(
            string path,
            string fileName,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sniffedMimeType = SniffImageMimeType(path);
            if (sniffedMimeType != null)
            {
                return ExtractImage(
                    path,
                    fileName,
                    sniffedMimeType,
                    cancellationToken);
            }

            var header = ReadHeader(path, 4096);
            if (header.Length > 3 &&
                header[0] == (byte)'P' &&
                header[1] == (byte)'K')
            {
                if (ZipContainsEntry(path, "word/document.xml"))
                {
                    return ExtractDocument(
                        path, fileName, ".docx", cancellationToken);
                }

                if (ZipContainsEntry(path, "ppt/presentation.xml"))
                {
                    return ExtractDocument(
                        path, fileName, ".pptx", cancellationToken);
                }

                if (ZipContainsEntry(path, "xl/workbook.xml"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xlsx",
                        cancellationToken);
                }

                if (ZipContainsEntry(path, "xl/workbook.bin"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xlsb",
                        cancellationToken);
                }

                if (ZipContainsEntry(path, "content.xml"))
                {
                    return ExtractDocument(
                        path, fileName, ".odt", cancellationToken);
                }

                return null;
            }

            if (LegacyOfficeTextExtractor.LooksLikeCompoundFile(
                header))
            {
                var compoundExtension =
                    LegacyOfficeTextExtractor
                        .IdentifyCompoundExtension(
                            path,
                            cancellationToken);
                return compoundExtension == ".xls"
                    ? ExtractSpreadsheet(
                        path,
                        fileName,
                        compoundExtension,
                        cancellationToken)
                    : compoundExtension.Length > 0
                        ? ExtractDocument(
                            path,
                            fileName,
                            compoundExtension,
                            cancellationToken)
                        : null;
            }

            if (header.Length > 4 &&
                header[0] == (byte)'%' &&
                header[1] == (byte)'P' &&
                header[2] == (byte)'D' &&
                header[3] == (byte)'F')
            {
                return ExtractDocument(
                    path,
                    fileName,
                    ".pdf",
                    cancellationToken);
            }

            if (LooksLikeText(header))
            {
                return new EmailAttachmentContent(
                    fileName,
                    "text",
                    ReadTextFile(path, cancellationToken));
            }

            return null;
        }

        private static byte[] ReadHeader(string path, int count)
        {
            using (var stream = File.OpenRead(path))
            {
                var buffer = new byte[count];
                var read = stream.Read(buffer, 0, buffer.Length);
                if (read == buffer.Length)
                {
                    return buffer;
                }

                var bounded = new byte[Math.Max(0, read)];
                Array.Copy(buffer, bounded, bounded.Length);
                return bounded;
            }
        }

        private static bool ZipContainsEntry(
            string path,
            string entryName)
        {
            try
            {
                using (var zip = ZipFile.OpenRead(path))
                {
                    return zip.GetEntry(entryName) != null;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void ValidateArchive(ZipArchive archive)
        {
            if (archive == null ||
                archive.Entries.Count > MaxArchiveEntries)
            {
                throw new AttachmentResourceLimitException(
                    "The archive exceeded the 10,000-entry cap.");
            }
        }

        private static Stream OpenBoundedPart(
            ZipArchiveEntry entry,
            ArchiveReadBudget budget)
        {
            return new BoundedArchivePartStream(
                entry.Open(),
                budget,
                MaxArchivePartBytes);
        }

        private static bool LooksLikeText(byte[] header)
        {
            if (header.Length == 0)
            {
                return false;
            }

            var printable = 0;
            foreach (var value in header)
            {
                if (value == 0)
                {
                    return false;
                }

                if (value == 9 ||
                    value == 10 ||
                    value == 13 ||
                    (value >= 32 && value < 127) ||
                    value >= 160)
                {
                    printable++;
                }
            }

            return printable * 100 >= header.Length * 90;
        }

        private static string WithLegacyHeader(
            string fileName,
            string format,
            string extracted)
        {
            return extracted.Trim().Length > 0
                ? "[" + format + " attachment: " + fileName +
                  " - legacy format, best-effort text extraction]\n" +
                  extracted
                : string.Empty;
        }

        private static EmailAttachmentContent ExtractDocument(
            string path,
            string fileName,
            string extension,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var kind = "document";
            string text;
            try
            {
                switch (extension.ToLowerInvariant())
                {
                    case ".pdf":
                        kind = "pdf";
                        text = ExtractPdfText(path, cancellationToken);
                        if (CountReadableCharacters(text) < 40)
                        {
                            text =
                                "[PDF attachment: " + fileName +
                                ". No machine-readable text could be " +
                                "extracted. The PDF is likely scanned " +
                                "pages or uses embedded font encodings. " +
                                "Ask the user to export it as text or " +
                                "paste the content into the email.]";
                        }
                        else
                        {
                            text =
                                "[PDF attachment: " + fileName +
                                " - best-effort text extraction; layout " +
                                "and some characters may be lost]\n" +
                                text;
                        }

                        break;
                    case ".pptx":
                    case ".pptm":
                    case ".ppsx":
                    case ".ppsm":
                    case ".potx":
                        kind = "powerpoint";
                        text = ExtractPptxText(path, cancellationToken);
                        break;
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        kind = "word";
                        text = ExtractDocxText(path, cancellationToken);
                        break;
                    case ".odt":
                    case ".ods":
                    case ".odp":
                        kind = "document";
                        text = ExtractOdfText(path, cancellationToken);
                        break;
                    case ".msg":
                    case ".oft":
                        kind = "email";
                        text = LegacyOfficeTextExtractor
                            .ExtractMsgText(path, cancellationToken);
                        if (text.Trim().Length > 0)
                        {
                            text =
                                "[Attached Outlook message: " +
                                fileName + "]\n" + text;
                        }

                        break;
                    case ".ppt":
                        kind = "powerpoint";
                        text = WithLegacyHeader(
                            fileName,
                            "PowerPoint",
                            LegacyOfficeTextExtractor
                                .ExtractPptText(
                                    path,
                                    cancellationToken));
                        break;
                    case ".doc":
                        kind = "word";
                        text = WithLegacyHeader(
                            fileName,
                            "Word",
                            LegacyOfficeTextExtractor
                                .ExtractDocText(
                                    path,
                                    cancellationToken));
                        break;
                    case ".rtf":
                        kind = "word";
                        text = LegacyOfficeTextExtractor
                            .ExtractRtfText(
                                path,
                                cancellationToken);
                        break;
                    default:
                        text = string.Empty;
                        break;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (AttachmentResourceLimitException exception)
            {
                return LimitContent(
                    fileName,
                    new FileInfo(path).Length,
                    exception.Message);
            }
            catch
            {
                text =
                    "[Attachment: " + fileName +
                    ". The file could not be parsed for text.]";
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                text =
                    "[Attachment: " + fileName +
                    ". No readable text was extracted.]";
            }

            return new EmailAttachmentContent(
                fileName,
                kind,
                text);
        }

        private static string ExtractPptxText(
            string path,
            CancellationToken cancellationToken)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                ValidateArchive(zip);
                var archiveBudget = new ArchiveReadBudget();
                var slides = zip.Entries
                    .Where(entry =>
                        entry.FullName.StartsWith(
                            "ppt/slides/slide",
                            StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(
                            ".xml",
                            StringComparison.OrdinalIgnoreCase))
                    .OrderBy(entry => SlideNumber(entry.FullName))
                    .ToList();
                var builder = new StringBuilder();
                foreach (var slide in slides)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var lineHasText = false;
                    using (var stream = OpenBoundedPart(
                        slide,
                        archiveBudget))
                    using (var reader = System.Xml.XmlReader.Create(
                        stream,
                        StreamingXmlSettings()))
                    {
                        builder.AppendLine(
                            "[Slide " +
                            SlideNumber(slide.FullName).ToString() +
                            "]");
                        while (reader.Read())
                        {
                            cancellationToken
                                .ThrowIfCancellationRequested();
                            if (reader.NodeType ==
                                    System.Xml.XmlNodeType.Element &&
                                reader.LocalName == "t")
                            {
                                var content =
                                    reader.ReadElementContentAsString();
                                if (content.Length > 0)
                                {
                                    builder.Append(content);
                                    lineHasText = true;
                                }

                                continue;
                            }

                            if (reader.NodeType ==
                                    System.Xml.XmlNodeType.EndElement &&
                                reader.LocalName == "p" &&
                                lineHasText)
                            {
                                builder.AppendLine();
                                lineHasText = false;
                            }

                            if (builder.Length >=
                                ScaledCharactersPerAttachment)
                            {
                                break;
                            }
                        }
                    }

                    if (builder.Length >= ScaledCharactersPerAttachment)
                    {
                        break;
                    }
                }

                return builder.ToString();
            }
        }

        private static int SlideNumber(string entryName)
        {
            var digits = new StringBuilder();
            foreach (var character in entryName)
            {
                if (char.IsDigit(character))
                {
                    digits.Append(character);
                }
                else if (digits.Length > 0)
                {
                    break;
                }
            }

            int number;
            return int.TryParse(digits.ToString(), out number)
                ? number
                : 0;
        }

        // Streamed with XmlReader so a large document part never
        // loads as a whole DOM.
        private static string ExtractDocxText(
            string path,
            CancellationToken cancellationToken)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                ValidateArchive(zip);
                var archiveBudget = new ArchiveReadBudget();
                var entry = zip.GetEntry("word/document.xml");
                if (entry == null)
                {
                    return string.Empty;
                }

                var builder = new StringBuilder();
                using (var stream = OpenBoundedPart(
                    entry,
                    archiveBudget))
                using (var reader = System.Xml.XmlReader.Create(
                    stream,
                    StreamingXmlSettings()))
                {
                    if (!reader.Read())
                    {
                        return string.Empty;
                    }

                    var lineHasText = false;
                    while (!reader.EOF)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (reader.NodeType ==
                                System.Xml.XmlNodeType.Element &&
                            reader.LocalName == "t")
                        {
                            var content =
                                reader.ReadElementContentAsString();
                            if (content.Length > 0)
                            {
                                builder.Append(content);
                                lineHasText = true;
                            }

                            continue;
                        }

                        if (reader.NodeType ==
                                System.Xml.XmlNodeType.EndElement &&
                            reader.LocalName == "p")
                        {
                            if (lineHasText)
                            {
                                builder.AppendLine();
                                lineHasText = false;
                            }

                            if (builder.Length >
                                ScaledCharactersPerAttachment)
                            {
                                break;
                            }
                        }

                        if (!reader.Read())
                        {
                            break;
                        }
                    }
                }

                return builder.ToString();
            }
        }

        // OpenDocument (.odt, .ods, .odp): all visible text lives in
        // content.xml. Text nodes are streamed and paragraph or
        // heading ends become line breaks, which also yields one line
        // per spreadsheet cell.
        private static string ExtractOdfText(
            string path,
            CancellationToken cancellationToken)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                ValidateArchive(zip);
                var archiveBudget = new ArchiveReadBudget();
                var entry = zip.GetEntry("content.xml");
                if (entry == null)
                {
                    return string.Empty;
                }

                var builder = new StringBuilder();
                using (var stream = OpenBoundedPart(
                    entry,
                    archiveBudget))
                using (var reader = System.Xml.XmlReader.Create(
                    stream,
                    StreamingXmlSettings()))
                {
                    if (!reader.Read())
                    {
                        return string.Empty;
                    }

                    var lineHasText = false;
                    while (!reader.EOF)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (reader.NodeType ==
                                System.Xml.XmlNodeType.Text ||
                            reader.NodeType ==
                                System.Xml.XmlNodeType.CDATA)
                        {
                            if (reader.Value.Length > 0)
                            {
                                builder.Append(reader.Value);
                                lineHasText = true;
                            }
                        }
                        else if (reader.NodeType ==
                                     System.Xml.XmlNodeType
                                         .EndElement &&
                                 (reader.LocalName == "p" ||
                                  reader.LocalName == "h"))
                        {
                            if (lineHasText)
                            {
                                builder.AppendLine();
                                lineHasText = false;
                            }

                            if (builder.Length >
                                ScaledCharactersPerAttachment)
                            {
                                break;
                            }
                        }

                        if (!reader.Read())
                        {
                            break;
                        }
                    }
                }

                return builder.ToString();
            }
        }

        private static System.Xml.XmlReaderSettings
            StreamingXmlSettings()
        {
            return new System.Xml.XmlReaderSettings
            {
                IgnoreWhitespace = false,
                IgnoreComments = true,
                DtdProcessing =
                    System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            };
        }

        private static string ExtractPdfText(
            string path,
            CancellationToken cancellationToken)
        {
            return PdfTextExtractor.Extract(
                path,
                ScaledCharactersPerAttachment,
                cancellationToken);
        }

        private static int CountReadableCharacters(string text)
        {
            var count = 0;
            foreach (var character in text ?? string.Empty)
            {
                if (char.IsLetterOrDigit(character))
                {
                    count++;
                }
            }

            return count;
        }

        private static string SniffImageMimeType(string path)
        {
            byte[] header;
            using (var stream = File.OpenRead(path))
            {
                header = new byte[12];
                var read = stream.Read(header, 0, header.Length);
                if (read < 4)
                {
                    return null;
                }
            }

            if (header[0] == 0x89 &&
                header[1] == 0x50 &&
                header[2] == 0x4E &&
                header[3] == 0x47)
            {
                return "image/png";
            }

            if (header[0] == 0xFF &&
                header[1] == 0xD8 &&
                header[2] == 0xFF)
            {
                return "image/jpeg";
            }

            if (header[0] == 'G' &&
                header[1] == 'I' &&
                header[2] == 'F' &&
                header[3] == '8')
            {
                return "image/gif";
            }

            if (header[0] == 'B' && header[1] == 'M')
            {
                return "image/bmp";
            }

            if (header[0] == 'R' &&
                header[1] == 'I' &&
                header[2] == 'F' &&
                header[3] == 'F' &&
                header[8] == 'W' &&
                header[9] == 'E' &&
                header[10] == 'B' &&
                header[11] == 'P')
            {
                return "image/webp";
            }

            if ((header[0] == 'I' &&
                 header[1] == 'I' &&
                 header[2] == 0x2A &&
                 header[3] == 0x00) ||
                (header[0] == 'M' &&
                 header[1] == 'M' &&
                 header[2] == 0x00 &&
                 header[3] == 0x2A))
            {
                return "image/tiff";
            }

            return null;
        }

        private static EmailAttachmentContent ExtractImage(
            string path,
            string fileName,
            string mimeType,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceBytes = new FileInfo(path).Length;
            var builder = new StringBuilder();
            builder.Append("[Image attachment: ");
            builder.Append(fileName);
            builder.Append(", ");
            builder.Append(sourceBytes.ToString());
            builder.Append(" bytes, type ");
            builder.Append(mimeType);
            builder.Append(']');

            string dataUrl = null;
            if (sourceBytes <= MaxImageBytesForBase64)
            {
                var bytes = ReadSmallFile(
                    path,
                    MaxImageBytesForBase64);
                dataUrl =
                    "data:" +
                    mimeType +
                    ";base64," +
                    Convert.ToBase64String(bytes);
            }
            else
            {
                string resourceWarning;
                var downscaled = TryDownscaleToJpeg(
                    path,
                    cancellationToken,
                    out resourceWarning);
                if (resourceWarning.Length > 0)
                {
                    return LimitContent(
                        fileName,
                        sourceBytes,
                        resourceWarning);
                }
                if (downscaled != null)
                {
                    builder.Append(
                        "\nThe image was downscaled locally to fit " +
                        "the vision size limit.");
                    dataUrl =
                        "data:image/jpeg;base64," +
                        Convert.ToBase64String(downscaled);
                }
            }

            if (dataUrl != null)
            {
                builder.Append(
                    "\nVision-capable models receive this image " +
                    "through multimodal input after tool results.");
            }
            else
            {
                builder.Append(
                    "\nImage exceeds the vision size limit and could " +
                    "not be downscaled. Only metadata is included.");
            }

            return new EmailAttachmentContent(
                fileName,
                "image",
                builder.ToString(),
                dataUrl);
        }

        private static byte[] TryDownscaleToJpeg(
            string path,
            CancellationToken cancellationToken,
            out string resourceWarning)
        {
            resourceWarning = string.Empty;
            try
            {
                using (var original =
                    System.Drawing.Image.FromFile(path))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSafeImageDimensions(
                            original.Width,
                            original.Height))
                    {
                        resourceWarning =
                            "Image dimensions exceed the safe " +
                            "32,768-pixel or 64-megapixel limit. " +
                            "Only metadata is included.";
                        return null;
                    }

                    var longSide = Math.Max(
                        original.Width,
                        original.Height);
                    var targetSide = Math.Min(longSide, 1280);
                    while (targetSide >= 256)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var scale = (double)targetSide / longSide;
                        var width = Math.Max(
                            1,
                            (int)Math.Round(original.Width * scale));
                        var height = Math.Max(
                            1,
                            (int)Math.Round(original.Height * scale));
                        using (var bitmap =
                            new System.Drawing.Bitmap(width, height))
                        {
                            using (var graphics =
                                System.Drawing.Graphics.FromImage(
                                    bitmap))
                            {
                                graphics.Clear(
                                    System.Drawing.Color.White);
                                graphics.InterpolationMode =
                                    System.Drawing.Drawing2D
                                        .InterpolationMode
                                        .HighQualityBicubic;
                                graphics.DrawImage(
                                    original,
                                    0,
                                    0,
                                    width,
                                    height);
                            }

                            foreach (var quality in
                                new long[] { 75, 55 })
                            {
                                var encoded = EncodeJpeg(
                                    bitmap,
                                    quality);
                                if (encoded != null &&
                                    encoded.Length <=
                                    MaxImageBytesForBase64)
                                {
                                    return encoded;
                                }
                            }
                        }

                        targetSide /= 2;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
            }

            return null;
        }

        public static bool IsSafeImageDimensions(
            int width,
            int height)
        {
            return width > 0 &&
                   height > 0 &&
                   width <= MaxImageDimension &&
                   height <= MaxImageDimension &&
                   (long)width * height <= MaxImagePixels;
        }

        private static byte[] ReadSmallFile(
            string path,
            int maximumBytes)
        {
            using (var stream = File.OpenRead(path))
            {
                if (stream.Length > maximumBytes)
                {
                    throw new AttachmentResourceLimitException(
                        "The image exceeded the model-image input cap.");
                }

                var bytes = new byte[(int)stream.Length];
                var offset = 0;
                while (offset < bytes.Length)
                {
                    var read = stream.Read(
                        bytes,
                        offset,
                        bytes.Length - offset);
                    if (read <= 0)
                    {
                        break;
                    }

                    offset += read;
                }

                if (offset != bytes.Length)
                {
                    throw new EndOfStreamException();
                }

                return bytes;
            }
        }

        private static byte[] EncodeJpeg(
            System.Drawing.Bitmap bitmap,
            long quality)
        {
            var encoder = System.Drawing.Imaging.ImageCodecInfo
                .GetImageEncoders()
                .FirstOrDefault(codec =>
                    codec.FormatID ==
                    System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
            if (encoder == null)
            {
                return null;
            }

            using (var parameters =
                new System.Drawing.Imaging.EncoderParameters(1))
            {
                parameters.Param[0] =
                    new System.Drawing.Imaging.EncoderParameter(
                        System.Drawing.Imaging.Encoder.Quality,
                        quality);
                using (var stream = new MemoryStream())
                {
                    bitmap.Save(stream, encoder, parameters);
                    return stream.ToArray();
                }
            }
        }

        private static EmailAttachmentContent ExtractSpreadsheet(
            string path,
            string fileName,
            string extension,
            CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            string text;
            try
            {
                if (extension.Equals(
                        ".csv",
                        StringComparison.OrdinalIgnoreCase) ||
                    extension.Equals(
                        ".tsv",
                        StringComparison.OrdinalIgnoreCase))
                {
                    text = ReadTextFile(path, cancellationToken);
                }
                else if (extension.Equals(
                    ".xlsb",
                    StringComparison.OrdinalIgnoreCase))
                {
                    text = XlsbTextExtractor.Extract(
                        path,
                        ScaledCharactersPerAttachment,
                        cancellationToken);
                }
                else if (extension.Equals(
                    ".xls",
                    StringComparison.OrdinalIgnoreCase))
                {
                    var extracted = LegacyOfficeTextExtractor
                        .ExtractXlsText(path, cancellationToken);
                    text = extracted.Trim().Length > 0
                        ? "[Excel attachment: " + fileName +
                          " - legacy .xls cell text without positions]\n" +
                          extracted
                        : "[Excel attachment: " + fileName +
                          ". No readable cell text was extracted from " +
                          "the legacy workbook. Save as .xlsx or .csv " +
                          "for full extraction.]";
                }
                else
                {
                    text = ExtractXlsxText(
                        path,
                        cancellationToken);
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (AttachmentResourceLimitException exception)
            {
                return LimitContent(
                    fileName,
                    new FileInfo(path).Length,
                    exception.Message);
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                text =
                    "[Excel attachment: " + fileName +
                    ". No readable cell text was extracted.]";
            }

            return new EmailAttachmentContent(
                fileName,
                "excel",
                text);
        }

        // Streamed with XmlReader: shared strings and rows are read
        // sequentially and extraction stops at the character budget, so
        // a 100 MB workbook never materializes a full XML DOM.
        private static string ExtractXlsxText(
            string path,
            CancellationToken cancellationToken)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                ValidateArchive(zip);
                var archiveBudget = new ArchiveReadBudget();
                var sharedStrings = ReadSharedStrings(
                    zip,
                    archiveBudget,
                    cancellationToken);
                var sheetEntries = zip.Entries
                    .Where(entry =>
                        entry.FullName.StartsWith(
                            "xl/worksheets/sheet",
                            StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(
                            ".xml",
                            StringComparison.OrdinalIgnoreCase))
                    .OrderBy(entry => SlideNumber(entry.FullName))
                    .ToList();
                var builder = new StringBuilder();
                foreach (var sheetEntry in sheetEntries)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (sheetEntries.Count > 1)
                    {
                        builder.AppendLine(
                            "[Sheet " +
                            SlideNumber(
                                sheetEntry.FullName).ToString() +
                            "]");
                    }

                    ExtractXlsxSheet(
                        sheetEntry,
                        sharedStrings,
                        builder,
                        archiveBudget,
                        cancellationToken);
                    if (builder.Length > ScaledCharactersPerAttachment)
                    {
                        break;
                    }
                }

                return builder.ToString();
            }
        }

        private static void ExtractXlsxSheet(
            ZipArchiveEntry sheetEntry,
            IList<string> sharedStrings,
            StringBuilder builder,
            ArchiveReadBudget archiveBudget,
            CancellationToken cancellationToken)
        {
            using (var stream = OpenBoundedPart(
                sheetEntry,
                archiveBudget))
            using (var reader = System.Xml.XmlReader.Create(
                stream,
                StreamingXmlSettings()))
            {
                if (!reader.Read())
                {
                    return;
                }

                string cellType = null;
                var rowValues = new List<string>();
                while (!reader.EOF)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType ==
                            System.Xml.XmlNodeType.Element &&
                        (reader.LocalName == "v" ||
                         reader.LocalName == "t"))
                    {
                        var lookupShared =
                            reader.LocalName == "v" &&
                            cellType == "s";
                        var content =
                            reader.ReadElementContentAsString();
                        if (lookupShared)
                        {
                            int sharedIndex;
                            content =
                                int.TryParse(
                                    content,
                                    out sharedIndex) &&
                                sharedIndex >= 0 &&
                                sharedIndex <
                                sharedStrings.Count
                                    ? sharedStrings[sharedIndex]
                                    : string.Empty;
                        }

                        if (content.Length > 0)
                        {
                            rowValues.Add(content);
                        }

                        continue;
                    }

                    if (reader.NodeType ==
                            System.Xml.XmlNodeType.Element &&
                        reader.LocalName == "c")
                    {
                        cellType = reader.GetAttribute("t");
                    }
                    else if (reader.NodeType ==
                                 System.Xml.XmlNodeType
                                     .EndElement &&
                             reader.LocalName == "row")
                    {
                        if (rowValues.Count > 0)
                        {
                            builder.AppendLine(
                                string.Join(
                                    "\t",
                                    rowValues));
                            rowValues.Clear();
                        }

                        if (builder.Length >
                            ScaledCharactersPerAttachment)
                        {
                            break;
                        }
                    }

                    if (!reader.Read())
                    {
                        break;
                    }
                }
            }
        }

        private static IList<string> ReadSharedStrings(
            ZipArchive zip,
            ArchiveReadBudget archiveBudget,
            CancellationToken cancellationToken)
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return new string[0];
            }

            var strings = new List<string>();
            var totalCharacters = 0L;
            using (var stream = OpenBoundedPart(
                entry,
                archiveBudget))
            using (var reader = System.Xml.XmlReader.Create(
                stream,
                StreamingXmlSettings()))
            {
                if (!reader.Read())
                {
                    return strings;
                }

                StringBuilder current = null;
                while (!reader.EOF &&
                       strings.Count < 500000 &&
                       totalCharacters < 16 * 1024 * 1024)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.NodeType ==
                            System.Xml.XmlNodeType.Element &&
                        reader.LocalName == "t")
                    {
                        var content =
                            reader.ReadElementContentAsString();
                        if (current != null)
                        {
                            current.Append(content);
                        }

                        continue;
                    }

                    if (reader.NodeType ==
                            System.Xml.XmlNodeType.Element &&
                        reader.LocalName == "si")
                    {
                        current = new StringBuilder();
                    }
                    else if (reader.NodeType ==
                                 System.Xml.XmlNodeType
                                     .EndElement &&
                             reader.LocalName == "si")
                    {
                        var value = current?.ToString() ??
                            string.Empty;
                        strings.Add(value);
                        totalCharacters += value.Length;
                        current = null;
                    }

                    if (!reader.Read())
                    {
                        break;
                    }
                }
            }

            return strings;
        }

        private static string ReadTextFile(
            string path,
            CancellationToken cancellationToken)
        {
            using (var reader = new StreamReader(
                path,
                Encoding.UTF8,
                true))
            {
                // One character beyond the cap so truncation is
                // detectable and disclosed; the content boundary
                // trims the text back to the cap.
                var buffer = new char[
                    ScaledCharactersPerAttachment + 1];
                var total = 0;
                while (total < buffer.Length)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var read = reader.Read(
                        buffer,
                        total,
                        buffer.Length - total);
                    if (read <= 0)
                    {
                        break;
                    }

                    total += read;
                }

                return new string(buffer, 0, total);
            }
        }

        private static string ImageMimeType(string extension)
        {
            switch (extension.ToLowerInvariant())
            {
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".webp":
                    return "image/webp";
                case ".tif":
                case ".tiff":
                    return "image/tiff";
                default:
                    return "application/octet-stream";
            }
        }

        private static string SafeString(Func<object> reader)
        {
            try
            {
                return Convert.ToString(reader()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.ReleaseComObject(value);
            }
        }
    }
}
