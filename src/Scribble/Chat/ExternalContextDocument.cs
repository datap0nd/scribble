using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class ExternalContextDocument
    {
        public const int MaxDocuments = 3;
        public const int MaxCharactersPerDocument = 48000;
        public const int MaxTotalCharacters = 120000;

        public ExternalContextDocument(
            string name,
            string content)
        {
            Name = TextBoundary.SingleLine(name, 180);
            Content = TextBoundary.PlainText(
                content,
                ContextScale.Scaled(MaxCharactersPerDocument));
        }

        public string Name { get; }

        public string Content { get; }

        public static IReadOnlyList<ExternalContextDocument> Normalize(
            IEnumerable<ExternalContextDocument> documents)
        {
            var result = new List<ExternalContextDocument>();
            var identities = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            var totalCharacters = 0;
            foreach (var document in documents ??
                Enumerable.Empty<ExternalContextDocument>())
            {
                if (document == null ||
                    document.Name.Length == 0 ||
                    document.Content.Length == 0 ||
                    !identities.Add(
                        document.Name + "\n" + document.Content))
                {
                    continue;
                }

                var remaining =
                    ContextScale.Scaled(MaxTotalCharacters) -
                    totalCharacters;
                if (remaining <= 0)
                {
                    break;
                }

                var content = TextBoundary.PlainText(
                    document.Content,
                    Math.Min(
                        ContextScale.Scaled(
                            MaxCharactersPerDocument),
                        remaining));
                result.Add(
                    new ExternalContextDocument(
                        document.Name,
                        content));
                totalCharacters += content.Length;
                if (result.Count == MaxDocuments)
                {
                    break;
                }
            }

            return result;
        }
    }

    public static class ExternalContextLoader
    {
        private static readonly HashSet<string> SupportedExtensions =
            new HashSet<string>(
                new[]
                {
                    ".txt", ".md", ".csv", ".json", ".xml",
                    ".html", ".htm", ".log", ".yaml", ".yml",
                    ".ini"
                },
                StringComparer.OrdinalIgnoreCase);

        public static IReadOnlyList<ExternalContextDocument> LoadFiles(
            IEnumerable<string> paths)
        {
            var selectedPaths = (paths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Take(ExternalContextDocument.MaxDocuments)
                .ToArray();
            if (selectedPaths.Length == 0)
            {
                throw new InvalidOperationException(
                    "Choose at least one supported text file.");
            }

            var documents = new List<ExternalContextDocument>();
            foreach (var path in selectedPaths)
            {
                var extension = Path.GetExtension(path);
                if (!SupportedExtensions.Contains(extension))
                {
                    throw new InvalidOperationException(
                        "Unsupported context file '" +
                        Path.GetFileName(path) +
                        "'. Use TXT, MD, CSV, JSON, XML, HTML, LOG, YAML, or INI.");
                }

                var file = new FileInfo(path);
                if (!file.Exists)
                {
                    throw new FileNotFoundException(
                        "The context file was not found.",
                        path);
                }

                if (file.Length > 2 * 1024 * 1024)
                {
                    throw new InvalidOperationException(
                        "Context files must be 2 MB or smaller before text limits are applied.");
                }

                string content;
                using (var reader = new StreamReader(
                    path,
                    Encoding.UTF8,
                    true))
                {
                    var buffer = new char[
                        ContextScale.Scaled(
                            ExternalContextDocument
                                .MaxCharactersPerDocument)];
                    var read = reader.Read(
                        buffer,
                        0,
                        buffer.Length);
                    content = new string(buffer, 0, read);
                }

                documents.Add(
                    new ExternalContextDocument(
                        Path.GetFileName(path),
                        content));
            }

            return ExternalContextDocument.Normalize(documents);
        }
    }
}
