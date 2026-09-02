using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace Scribble.Outlook
{
    internal sealed class CompoundFileReader : IDisposable
    {
        private const uint EndOfChain = 0xFFFFFFFE;
        private const uint FreeSector = 0xFFFFFFFF;
        private const int MaxChainLength = 300000;
        private const long MaxStreamBytes = 32L * 1024 * 1024;

        private sealed class DirectoryEntry
        {
            public byte Type;
            public uint Start;
            public long Size;
        }

        private readonly FileStream _stream;
        private readonly CancellationToken _cancellationToken;
        private readonly int _sectorSize;
        private readonly int _miniSectorSize;
        private readonly uint _miniCutoff;
        private readonly uint _firstMiniFatSector;
        private readonly List<uint> _fat;
        private readonly Dictionary<string, DirectoryEntry> _entries;
        private readonly DirectoryEntry _root;
        private List<uint> _miniFat;
        private List<uint> _rootChain;

        public CompoundFileReader(string path)
            : this(path, CancellationToken.None)
        {
        }

        public CompoundFileReader(
            string path,
            CancellationToken cancellationToken)
        {
            _cancellationToken = cancellationToken;
            _stream = new FileStream(
                path,
                FileMode.Open,
                FileAccess.Read,
                FileShare.Read);
            try
            {
                var header = ReadAt(0, 512);
                if (header.Length < 512 ||
                    header[0] != 0xD0 ||
                    header[1] != 0xCF ||
                    header[2] != 0x11 ||
                    header[3] != 0xE0)
                {
                    throw new InvalidDataException(
                        "Not an OLE compound file.");
                }

                var sectorShift = ReadUInt16(header, 30);
                if (sectorShift < 7 || sectorShift > 12)
                {
                    throw new InvalidDataException(
                        "Unsupported compound-file sector size.");
                }

                _sectorSize = 1 << sectorShift;
                var miniShift = ReadUInt16(header, 32);
                _miniSectorSize = 1 << Math.Max(
                    4,
                    Math.Min(12, (int)miniShift));
                if (_miniSectorSize > _sectorSize)
                {
                    throw new InvalidDataException(
                        "Invalid compound-file mini-sector size.");
                }
                var firstDirectorySector = ReadUInt32(header, 48);
                _miniCutoff = ReadUInt32(header, 56);
                _firstMiniFatSector = ReadUInt32(header, 60);
                var firstDifatSector = ReadUInt32(header, 68);

                var fatSectors = new List<uint>();
                for (var index = 0; index < 109; index++)
                {
                    AddSector(
                        fatSectors,
                        ReadUInt32(header, 76 + index * 4));
                }

                var difatSector = firstDifatSector;
                var difatGuard = 0;
                while (IsSector(difatSector) &&
                       difatGuard++ < MaxChainLength)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var sector = ReadSector(difatSector);
                    var perSector = _sectorSize / 4 - 1;
                    for (var index = 0; index < perSector; index++)
                    {
                        AddSector(
                            fatSectors,
                            ReadUInt32(sector, index * 4));
                    }

                    difatSector = ReadUInt32(
                        sector,
                        _sectorSize - 4);
                }

                _fat = new List<uint>();
                foreach (var fatSector in fatSectors)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var sector = ReadSector(fatSector);
                    for (var index = 0;
                         index < _sectorSize / 4;
                         index++)
                    {
                        _fat.Add(ReadUInt32(sector, index * 4));
                    }
                }

                var directory = ReadRegularChain(
                    firstDirectorySector,
                    MaxStreamBytes);
                _entries = new Dictionary<string, DirectoryEntry>(
                    StringComparer.OrdinalIgnoreCase);
                for (var offset = 0;
                     offset + 128 <= directory.Length;
                     offset += 128)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var nameLength = ReadUInt16(
                        directory,
                        offset + 64);
                    if (nameLength < 2 || nameLength > 64)
                    {
                        continue;
                    }

                    var name = Encoding.Unicode.GetString(
                        directory,
                        offset,
                        nameLength - 2);
                    var entry = new DirectoryEntry
                    {
                        Type = directory[offset + 66],
                        Start = ReadUInt32(directory, offset + 116),
                        Size = (long)Math.Min(
                            ReadUInt64(directory, offset + 120),
                            (ulong)long.MaxValue)
                    };
                    if (!_entries.ContainsKey(name))
                    {
                        _entries[name] = entry;
                    }

                    if (entry.Type == 5)
                    {
                        _root = entry;
                    }
                }
            }
            catch
            {
                _stream.Dispose();
                throw;
            }
        }

        public bool ContainsStream(string name)
        {
            DirectoryEntry entry;
            return _entries.TryGetValue(name, out entry) &&
                   entry.Type == 2;
        }

        public byte[] ReadStream(string name)
        {
            DirectoryEntry entry;
            if (!_entries.TryGetValue(name, out entry) ||
                entry.Type != 2)
            {
                return null;
            }

            if (entry.Size > MaxStreamBytes)
            {
                throw new AttachmentResourceLimitException(
                    "A legacy Office data stream exceeded the " +
                    "32 MB extraction cap.");
            }

            return entry.Size >= _miniCutoff
                ? ReadRegularChain(entry.Start, entry.Size)
                : ReadMiniChain(entry.Start, entry.Size);
        }

        private byte[] ReadMiniChain(uint start, long size)
        {
            EnsureMiniStructures();
            using (var output = new MemoryStream())
            {
                var sector = start;
                var remaining = size;
                var guard = 0;
                while (IsSector(sector) &&
                       remaining > 0 &&
                       guard++ < MaxChainLength)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    var logicalOffset =
                        (long)sector * _miniSectorSize;
                    var rootIndex =
                        (int)(logicalOffset / _sectorSize);
                    var withinSector =
                        (int)(logicalOffset % _sectorSize);
                    if (rootIndex < 0 ||
                        rootIndex >= _rootChain.Count)
                    {
                        break;
                    }

                    var take = (int)Math.Min(
                        Math.Min(
                            _miniSectorSize,
                            _sectorSize - withinSector),
                        remaining);
                    var bytes = ReadAt(
                        SectorOffset(_rootChain[rootIndex]) +
                        withinSector,
                        take);
                    output.Write(bytes, 0, bytes.Length);
                    remaining -= bytes.Length;
                    sector = sector < _miniFat.Count
                        ? _miniFat[(int)sector]
                        : EndOfChain;
                }

                return output.ToArray();
            }
        }

        private void EnsureMiniStructures()
        {
            if (_miniFat != null)
            {
                return;
            }

            _miniFat = new List<uint>();
            var bytes = ReadRegularChain(
                _firstMiniFatSector,
                MaxStreamBytes);
            for (var offset = 0;
                 offset + 4 <= bytes.Length;
                 offset += 4)
            {
                _miniFat.Add(ReadUInt32(bytes, offset));
            }

            _rootChain = _root == null
                ? new List<uint>()
                : BuildChain(_root.Start);
        }

        private byte[] ReadRegularChain(uint start, long maximumBytes)
        {
            var chain = BuildChain(start);
            if (maximumBytes == MaxStreamBytes &&
                (long)chain.Count * _sectorSize > MaxStreamBytes)
            {
                throw new AttachmentResourceLimitException(
                    "A compound-file metadata stream exceeded the " +
                    "32 MB extraction cap.");
            }

            using (var output = new MemoryStream())
            {
                foreach (var sectorNumber in chain)
                {
                    _cancellationToken.ThrowIfCancellationRequested();
                    if (output.Length >= maximumBytes)
                    {
                        break;
                    }

                    var sector = ReadSector(sectorNumber);
                    var take = (int)Math.Min(
                        sector.Length,
                        maximumBytes - output.Length);
                    output.Write(sector, 0, take);
                }

                return output.ToArray();
            }
        }

        private List<uint> BuildChain(uint start)
        {
            var chain = new List<uint>();
            var sector = start;
            var guard = 0;
            while (IsSector(sector) &&
                   sector < _fat.Count &&
                   guard++ < MaxChainLength)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                chain.Add(sector);
                sector = _fat[(int)sector];
            }

            return chain;
        }

        private byte[] ReadSector(uint sector)
        {
            return ReadAt(SectorOffset(sector), _sectorSize);
        }

        private long SectorOffset(uint sector)
        {
            return ((long)sector + 1L) * _sectorSize;
        }

        private byte[] ReadAt(long offset, int count)
        {
            if (offset < 0 ||
                count < 0 ||
                offset > _stream.Length - count)
            {
                throw new InvalidDataException(
                    "Compound-file sector is outside the source file.");
            }

            var bytes = new byte[count];
            _stream.Position = offset;
            var readTotal = 0;
            while (readTotal < count)
            {
                _cancellationToken.ThrowIfCancellationRequested();
                var read = _stream.Read(
                    bytes,
                    readTotal,
                    count - readTotal);
                if (read <= 0)
                {
                    throw new EndOfStreamException();
                }

                readTotal += read;
            }

            return bytes;
        }

        private static bool IsSector(uint value)
        {
            return value != EndOfChain && value != FreeSector;
        }

        private static void AddSector(List<uint> values, uint value)
        {
            if (IsSector(value))
            {
                values.Add(value);
            }
        }

        private static ushort ReadUInt16(byte[] bytes, int offset)
        {
            return (ushort)(bytes[offset] | bytes[offset + 1] << 8);
        }

        private static uint ReadUInt32(byte[] bytes, int offset)
        {
            return (uint)(bytes[offset] |
                bytes[offset + 1] << 8 |
                bytes[offset + 2] << 16 |
                bytes[offset + 3] << 24);
        }

        private static ulong ReadUInt64(byte[] bytes, int offset)
        {
            return ReadUInt32(bytes, offset) |
                   ((ulong)ReadUInt32(bytes, offset + 4) << 32);
        }

        public void Dispose()
        {
            _stream.Dispose();
        }
    }
}
