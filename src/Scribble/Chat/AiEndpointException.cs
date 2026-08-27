using System;
using System.Globalization;
using System.Text;
using Scribble.Security;

namespace Scribble.Chat
{
    public sealed class AiEndpointException : Exception
    {
        public AiEndpointException(
            string code,
            string message,
            Exception innerException = null,
            int? httpStatus = null,
            string providerCode = null,
            string requestId = null,
            string responseSnippet = null,
            string transportDetails = null)
            : base(message, innerException)
        {
            Code = string.IsNullOrWhiteSpace(code)
                ? "AI_ENDPOINT_ERROR"
                : code.Trim();
            HttpStatus = httpStatus;
            ProviderCode = Clean(providerCode, 200);
            RequestId = Clean(requestId, 200);
            ResponseSnippet = Clean(responseSnippet, 1600);
            TransportDetails = Clean(transportDetails, 2400);
        }

        public string Code { get; }

        public int? HttpStatus { get; }

        public string ProviderCode { get; }

        public string RequestId { get; }

        public string ResponseSnippet { get; }

        public string TransportDetails { get; }

        public string ToDiagnosticText()
        {
            var builder = new StringBuilder();
            builder.Append("[");
            builder.Append(Code);
            builder.Append("] ");
            builder.Append(Message);

            if (HttpStatus.HasValue)
            {
                builder.Append(" HTTP status: ");
                builder.Append(
                    HttpStatus.Value.ToString(CultureInfo.InvariantCulture));
                builder.Append(".");
            }

            if (ProviderCode.Length > 0)
            {
                builder.Append(" Provider code: ");
                builder.Append(ProviderCode);
                builder.Append(".");
            }

            if (RequestId.Length > 0)
            {
                builder.Append(" Request ID: ");
                builder.Append(RequestId);
                builder.Append(".");
            }

            if (ResponseSnippet.Length > 0)
            {
                builder.AppendLine();
                builder.Append("Endpoint response: ");
                builder.Append(ResponseSnippet);
            }

            if (TransportDetails.Length > 0)
            {
                builder.AppendLine();
                builder.Append("Transport details: ");
                builder.Append(TransportDetails);
            }

            return builder.ToString();
        }

        private static string Clean(string value, int maximumLength)
        {
            return TextBoundary.PlainText(value, maximumLength)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
        }
    }
}
