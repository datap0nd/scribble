namespace Scribble.Outlook
{
    public sealed class DraftReference
    {
        public DraftReference(
            string kind,
            string subject,
            string to,
            string cc,
            string body)
        {
            Kind = kind ?? string.Empty;
            Subject = subject ?? string.Empty;
            To = to ?? string.Empty;
            Cc = cc ?? string.Empty;
            Body = body ?? string.Empty;
        }

        public string Kind { get; }

        public string Subject { get; }

        public string To { get; }

        public string Cc { get; }

        public string Body { get; }
    }
}
