using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;

namespace LedgerLens.Core
{
    /// <summary>
    /// Distinguishes open workbook objects even when Save Copy preserves their document GUID.
    /// The Excel adapter calls this on its main thread and removes entries on close/unload.
    /// </summary>
    public sealed class WorkbookSessions
    {
        private readonly Dictionary<object, Session> sessions = new Dictionary<object, Session>(new ReferenceComparer());

        public string GetId(object workbook, string documentId)
        {
            if (workbook == null)
                throw new ArgumentNullException(nameof(workbook));
            if (!Guid.TryParseExact(documentId, "N", out _))
                throw new InvalidOperationException("The workbook identity is invalid. Create a new LedgerLens model.");
            if (!sessions.TryGetValue(workbook, out var session))
            {
                session = new Session(documentId);
                sessions.Add(workbook, session);
            }
            if (session.DocumentId != documentId)
                throw new InvalidOperationException("The workbook identity changed. Reopen it before preparing another update.");
            return session.Id;
        }

        public bool Remove(object workbook, out string id)
        {
            if (!sessions.TryGetValue(workbook, out var session))
            {
                id = "";
                return false;
            }
            sessions.Remove(workbook);
            id = session.Id;
            return true;
        }

        public void Clear() => sessions.Clear();

        private sealed class Session
        {
            public string DocumentId
            {
                get;
            }
            public string Id { get; } = Guid.NewGuid().ToString("N");
            public Session(string documentId)
            {
                DocumentId = documentId;
            }
        }

        private sealed class ReferenceComparer : IEqualityComparer<object>
        {
            public new bool Equals(object? x, object? y) => ReferenceEquals(x, y);
            public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
        }
    }
}
