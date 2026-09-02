using System;
using System.IO;

namespace Scribble.Outlook
{
    public static class AttachmentIntakePolicy
    {
        public const long MaxFileBytes = 100L * 1024 * 1024;
        public const long MaxOperationBytes = 250L * 1024 * 1024;

        public static string ValidateFile(long sizeBytes)
        {
            if (sizeBytes < 0)
            {
                return "The attachment size is unavailable.";
            }

            if (sizeBytes > MaxFileBytes)
            {
                return "Over the 100 MB per-file cap - content not read.";
            }

            return string.Empty;
        }
    }

    public sealed class AttachmentReadBudget
    {
        private long _usedBytes;

        public long UsedBytes
        {
            get { return _usedBytes; }
        }

        public long RemainingBytes
        {
            get
            {
                return Math.Max(
                    0L,
                    AttachmentIntakePolicy.MaxOperationBytes -
                    _usedBytes);
            }
        }

        public bool TryReserve(long sizeBytes, out string warning)
        {
            warning = AttachmentIntakePolicy.ValidateFile(sizeBytes);
            if (warning.Length > 0)
            {
                return false;
            }

            if (sizeBytes > RemainingBytes)
            {
                warning =
                    "Over the 250 MB attachment budget for this " +
                    "operation - content not read.";
                return false;
            }

            _usedBytes += sizeBytes;
            return true;
        }
    }

    internal sealed class AttachmentResourceLimitException : IOException
    {
        public AttachmentResourceLimitException(string message)
            : base(message)
        {
        }
    }

    internal sealed class ArchiveReadBudget
    {
        public const long MaxTotalBytes = 128L * 1024 * 1024;

        private long _readBytes;

        public void Add(long count)
        {
            if (count < 0 || _readBytes > MaxTotalBytes - count)
            {
                throw new AttachmentResourceLimitException(
                    "The archive exceeded the 128 MB decompressed " +
                    "text-input cap.");
            }

            _readBytes += count;
        }
    }

    internal sealed class BoundedArchivePartStream : Stream
    {
        private readonly Stream _inner;
        private readonly ArchiveReadBudget _totalBudget;
        private readonly long _maxBytes;
        private long _readBytes;

        public BoundedArchivePartStream(
            Stream inner,
            ArchiveReadBudget totalBudget,
            long maxBytes)
        {
            _inner = inner ??
                throw new ArgumentNullException(nameof(inner));
            _totalBudget = totalBudget ??
                throw new ArgumentNullException(nameof(totalBudget));
            _maxBytes = maxBytes;
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length
        {
            get { throw new NotSupportedException(); }
        }
        public override long Position
        {
            get { return _readBytes; }
            set { throw new NotSupportedException(); }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_readBytes >= _maxBytes)
            {
                var probe = _inner.ReadByte();
                if (probe < 0)
                {
                    return 0;
                }

                _totalBudget.Add(1);
                throw new AttachmentResourceLimitException(
                    "An archive text part exceeded the 32 MB " +
                    "decompressed cap.");
            }

            var allowed = (int)Math.Min(
                count,
                _maxBytes - _readBytes);
            var read = _inner.Read(buffer, offset, allowed);
            _readBytes += read;
            _totalBudget.Add(read);
            return read;
        }

        public override void Flush()
        {
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
