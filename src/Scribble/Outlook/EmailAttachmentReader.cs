using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
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
            var characterLimit = ContextScale.Scaled(
                EmailAttachmentReader.MaxCharactersPerAttachment);
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

    public static class EmailAttachmentReader
    {
        public const int MaxAttachments = 10;
        // Uniform 25 MB intake: extraction output stays bounded in
        // characters, so large files cost read time, not context.
        public const int MaxBytesPerAttachment = 25 * 1024 * 1024;
        public const int MaxImageBytesPerAttachment = 25 * 1024 * 1024;
        public const int MaxCharactersPerAttachment = 48000;
        public const int MaxTotalCharacters = 120000;
        // The data URL cap must stay comfortably above the encoded
        // image byte cap (800 KB is ~1.1M base64 characters).
        public const int MaxImageDataUrlCharacters = 2200000;

        // Text budgets scale with the provider (see ContextScale);
        // byte and count caps do not.
        private static int ScaledCharactersPerAttachment
        {
            get
            {
                return ContextScale.Scaled(
                    MaxCharactersPerAttachment);
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
                var totalCharacters = 0;
                var signatureImagesSkipped = 0;
                for (var index = 1;
                     index <= count &&
                     totalCharacters < ScaledTotalCharacters;
                     index++)
                {
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
                        if (IsLikelySignatureImage(
                            attachment,
                            extension,
                            SafeLong(() => outlookAttachment.Size)))
                        {
                            signatureImagesSkipped++;
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

                        var fileInfo = new FileInfo(tempPath);
                        if (!fileInfo.Exists)
                        {
                            continue;
                        }

                        var sizeLimit =
                            ExcelExtensions.Contains(extension) ||
                            TextExtensions.Contains(extension)
                                ? MaxBytesPerAttachment
                                : MaxImageBytesPerAttachment;
                        if (fileInfo.Length > sizeLimit)
                        {
                            results.Add(new EmailAttachmentContent(
                                fileName,
                                "unreadable",
                                "[Attachment: " + fileName + ", " +
                                fileInfo.Length.ToString() +
                                " bytes. Too large for Scribble to " +
                                "read.]"));
                            totalCharacters += 80;
                            continue;
                        }

                        var extracted = ExtractContent(
                            tempPath,
                            fileName,
                            extension);
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
            try
            {
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
                var sizeLimit =
                    ExcelExtensions.Contains(extension) ||
                    TextExtensions.Contains(extension)
                        ? MaxBytesPerAttachment
                        : MaxImageBytesPerAttachment;
                if (info.Length > sizeLimit)
                {
                    return new EmailAttachmentContent(
                        fileName,
                        "unreadable",
                        "[File: " + fileName + ", " +
                        info.Length.ToString() +
                        " bytes. Too large for Scribble to read.]");
                }

                var extracted = ExtractContent(
                    path,
                    fileName,
                    extension);
                return extracted ?? new EmailAttachmentContent(
                    fileName,
                    "unreadable",
                    "[File: " + fileName +
                    ". This file type could not be converted to " +
                    "text or image input.]");
            }
            catch
            {
                return null;
            }
        }

        // Small JPEG preview for attachment tray thumbnails.
        public static string BuildThumbnailDataUrl(string path)
        {
            try
            {
                using (var original =
                    System.Drawing.Image.FromFile(path))
                {
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
            string extension)
        {
            if (ImageExtensions.Contains(extension))
            {
                return ExtractImage(
                    path,
                    fileName,
                    ImageMimeType(extension));
            }

            if (ExcelExtensions.Contains(extension))
            {
                return ExtractSpreadsheet(path, fileName, extension);
            }

            if (DocumentExtensions.Contains(extension))
            {
                return ExtractDocument(path, fileName, extension);
            }

            if (TextExtensions.Contains(extension))
            {
                return new EmailAttachmentContent(
                    fileName,
                    "text",
                    ReadTextFile(path));
            }

            return SniffUnknownContent(path, fileName);
        }

        // Unknown or missing extensions are identified by content so
        // every attachment is at least attempted: image magic bytes,
        // OOXML zip parts, OLE compound streams, then plain text.
        private static EmailAttachmentContent SniffUnknownContent(
            string path,
            string fileName)
        {
            var sniffedMimeType = SniffImageMimeType(path);
            if (sniffedMimeType != null)
            {
                return ExtractImage(path, fileName, sniffedMimeType);
            }

            var header = ReadHeader(path, 4096);
            if (header.Length > 3 &&
                header[0] == (byte)'P' &&
                header[1] == (byte)'K')
            {
                if (ZipContainsEntry(path, "word/document.xml"))
                {
                    return ExtractDocument(path, fileName, ".docx");
                }

                if (ZipContainsEntry(path, "ppt/presentation.xml"))
                {
                    return ExtractDocument(path, fileName, ".pptx");
                }

                if (ZipContainsEntry(path, "xl/workbook.xml"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xlsx");
                }

                if (ZipContainsEntry(path, "xl/workbook.bin"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xlsb");
                }

                if (ZipContainsEntry(path, "content.xml"))
                {
                    return ExtractDocument(path, fileName, ".odt");
                }

                return null;
            }

            if (LegacyOfficeTextExtractor.LooksLikeCompoundFile(
                header))
            {
                var bytes = File.ReadAllBytes(path);
                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                    bytes,
                    "WordDocument"))
                {
                    return ExtractDocument(path, fileName, ".doc");
                }

                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                    bytes,
                    "PowerPoint Document"))
                {
                    return ExtractDocument(path, fileName, ".ppt");
                }

                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                        bytes,
                        "Workbook") ||
                    LegacyOfficeTextExtractor.CompoundStreamExists(
                        bytes,
                        "Book"))
                {
                    return ExtractSpreadsheet(
                        path,
                        fileName,
                        ".xls");
                }

                if (LegacyOfficeTextExtractor.CompoundStreamExists(
                        bytes,
                        "__properties_version1.0") ||
                    LegacyOfficeTextExtractor.CompoundStreamExists(
                        bytes,
                        "__substg1.0_0037001F"))
                {
                    return ExtractDocument(path, fileName, ".msg");
                }

                return null;
            }

            if (header.Length > 4 &&
                header[0] == (byte)'%' &&
                header[1] == (byte)'P' &&
                header[2] == (byte)'D' &&
                header[3] == (byte)'F')
            {
                return ExtractDocument(path, fileName, ".pdf");
            }

            if (LooksLikeText(header))
            {
                return new EmailAttachmentContent(
                    fileName,
                    "text",
                    ReadTextFile(path));
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
            string extension)
        {
            var kind = "document";
            string text;
            try
            {
                switch (extension.ToLowerInvariant())
                {
                    case ".pdf":
                        kind = "pdf";
                        text = ExtractPdfText(path);
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
                        text = ExtractPptxText(path);
                        break;
                    case ".docx":
                    case ".docm":
                    case ".dotx":
                    case ".dotm":
                        kind = "word";
                        text = ExtractDocxText(path);
                        break;
                    case ".odt":
                    case ".ods":
                    case ".odp":
                        kind = "document";
                        text = ExtractOdfText(path);
                        break;
                    case ".msg":
                    case ".oft":
                        kind = "email";
                        text = LegacyOfficeTextExtractor
                            .ExtractMsgText(
                                File.ReadAllBytes(path));
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
                                    File.ReadAllBytes(path)));
                        break;
                    case ".doc":
                        kind = "word";
                        text = WithLegacyHeader(
                            fileName,
                            "Word",
                            LegacyOfficeTextExtractor
                                .ExtractDocText(
                                    File.ReadAllBytes(path)));
                        break;
                    case ".rtf":
                        kind = "word";
                        text = LegacyOfficeTextExtractor
                            .ExtractRtfText(
                                File.ReadAllBytes(path));
                        break;
                    default:
                        text = string.Empty;
                        break;
                }
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

        private static string ExtractPptxText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                XNamespace drawingNamespace =
                    "http://schemas.openxmlformats.org/drawingml/2006/main";
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
                    XDocument document;
                    using (var stream = slide.Open())
                    {
                        document = XDocument.Load(stream);
                    }

                    builder.AppendLine(
                        "[Slide " +
                        SlideNumber(slide.FullName).ToString() +
                        "]");
                    foreach (var paragraph in document.Descendants(
                        drawingNamespace + "p"))
                    {
                        var text = string.Concat(
                            paragraph
                                .Descendants(drawingNamespace + "t")
                                .Select(node => node.Value));
                        if (text.Trim().Length > 0)
                        {
                            builder.AppendLine(text);
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
        private static string ExtractDocxText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var entry = zip.GetEntry("word/document.xml");
                if (entry == null)
                {
                    return string.Empty;
                }

                var builder = new StringBuilder();
                using (var stream = entry.Open())
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
        private static string ExtractOdfText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var entry = zip.GetEntry("content.xml");
                if (entry == null)
                {
                    return string.Empty;
                }

                var builder = new StringBuilder();
                using (var stream = entry.Open())
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

        private static string ExtractPdfText(string path)
        {
            return PdfTextExtractor.Extract(
                File.ReadAllBytes(path),
                ScaledCharactersPerAttachment);
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
            string mimeType)
        {
            var bytes = File.ReadAllBytes(path);
            var builder = new StringBuilder();
            builder.Append("[Image attachment: ");
            builder.Append(fileName);
            builder.Append(", ");
            builder.Append(bytes.Length.ToString());
            builder.Append(" bytes, type ");
            builder.Append(mimeType);
            builder.Append(']');

            string dataUrl = null;
            if (bytes.Length <= MaxImageBytesForBase64)
            {
                dataUrl =
                    "data:" +
                    mimeType +
                    ";base64," +
                    Convert.ToBase64String(bytes);
            }
            else
            {
                var downscaled = TryDownscaleToJpeg(path);
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

        private static byte[] TryDownscaleToJpeg(string path)
        {
            try
            {
                using (var original =
                    System.Drawing.Image.FromFile(path))
                {
                    var longSide = Math.Max(
                        original.Width,
                        original.Height);
                    var targetSide = Math.Min(longSide, 1280);
                    while (targetSide >= 256)
                    {
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
            catch
            {
            }

            return null;
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
            string extension)
        {
            string text;
            if (extension.Equals(
                    ".csv",
                    StringComparison.OrdinalIgnoreCase) ||
                extension.Equals(
                    ".tsv",
                    StringComparison.OrdinalIgnoreCase))
            {
                text = ReadTextFile(path);
            }
            else if (extension.Equals(
                ".xlsb",
                StringComparison.OrdinalIgnoreCase))
            {
                text = XlsbTextExtractor.Extract(
                    File.ReadAllBytes(path),
                    ScaledCharactersPerAttachment);
            }
            else if (extension.Equals(
                ".xls",
                StringComparison.OrdinalIgnoreCase))
            {
                var extracted = LegacyOfficeTextExtractor
                    .ExtractXlsText(File.ReadAllBytes(path));
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
                text = ExtractXlsxText(path);
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
        // a 25 MB workbook never materializes a full XML DOM.
        private static string ExtractXlsxText(string path)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var sharedStrings = ReadSharedStrings(zip);
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
                        builder);
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
            StringBuilder builder)
        {
            using (var stream = sheetEntry.Open())
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

        private static IList<string> ReadSharedStrings(ZipArchive zip)
        {
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return new string[0];
            }

            var strings = new List<string>();
            var totalCharacters = 0L;
            using (var stream = entry.Open())
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

        private static string ReadTextFile(string path)
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
