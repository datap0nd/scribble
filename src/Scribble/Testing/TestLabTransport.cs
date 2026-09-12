using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;

namespace Scribble.Testing
{
    // Office security products can transparently wrap every file written by an
    // Office process. Test Lab therefore sends live receipts and trace events
    // to the standalone runner, which is the only process that persists them.
    internal sealed class TestLabTransport : IDisposable
    {
        private readonly Thread worker;
        private readonly object gate = new object();
        private NamedPipeServerStream listener;
        private bool stopping;

        internal string PipeName { get; private set; }

        internal TestLabTransport()
        {
            PipeName = "scribble-testlab-" + Guid.NewGuid().ToString("N");
            worker = new Thread(Listen) { IsBackground = true, Name = "Scribble Test Lab transport" };
            worker.Start();
        }

        internal static void Send(string pipeName, string kind, string runId, string itemId, string payload)
        {
            if (string.IsNullOrEmpty(pipeName)) throw new IOException("Test Lab transport is unavailable.");
            using (var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut, PipeOptions.None))
            {
                client.Connect(5000);
                using (var writer = new BinaryWriter(client, new UTF8Encoding(false), true))
                using (var reader = new BinaryReader(client, new UTF8Encoding(false), true))
                {
                    writer.Write(kind ?? "");
                    writer.Write(runId ?? "");
                    writer.Write(itemId ?? "");
                    writer.Write(payload ?? "");
                    writer.Flush();
                    var accepted = reader.ReadBoolean();
                    var error = reader.ReadString();
                    if (!accepted) throw new IOException("Test Lab transport rejected the record: " + error);
                }
            }
        }

        private void Listen()
        {
            while (true)
            {
                NamedPipeServerStream server = null;
                try
                {
                    lock (gate)
                    {
                        if (stopping) return;
                        server = new NamedPipeServerStream(PipeName, PipeDirection.InOut, 1,
                            PipeTransmissionMode.Byte, PipeOptions.None, 65536, 65536);
                        listener = server;
                    }
                    server.WaitForConnection();
                    using (var reader = new BinaryReader(server, new UTF8Encoding(false), true))
                    using (var writer = new BinaryWriter(server, new UTF8Encoding(false), true))
                    {
                        try
                        {
                            var kind = reader.ReadString();
                            var runId = reader.ReadString();
                            var itemId = reader.ReadString();
                            var payload = reader.ReadString();
                            if (kind == "event") TestLab.PersistTransportedEvent(runId, payload);
                            else if (kind == "receipt") TestLabSuitePane.PersistTransportedReceipt(runId, itemId, payload);
                            else throw new InvalidDataException("Unknown transport record.");
                            writer.Write(true); writer.Write(""); writer.Flush();
                        }
                        catch (Exception error)
                        {
                            try { writer.Write(false); writer.Write(error.GetType().Name + ": " + error.Message); writer.Flush(); }
                            catch { }
                        }
                    }
                }
                catch (ObjectDisposedException) { if (stopping) return; }
                catch (IOException) { if (stopping) return; }
                finally
                {
                    lock (gate) { if (ReferenceEquals(listener, server)) listener = null; }
                    if (server != null) server.Dispose();
                }
            }
        }

        public void Dispose()
        {
            lock (gate)
            {
                stopping = true;
                if (listener != null) listener.Dispose();
            }
            if (Thread.CurrentThread != worker) worker.Join(2000);
        }
    }
}
