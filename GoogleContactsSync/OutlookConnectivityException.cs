using System;
using System.Runtime.Serialization;

namespace GoContactSyncMod
{
    [Serializable]
    public class OutlookConnectivityException : Exception
    {
        public OutlookConnectivityException() : base() { }
        public OutlookConnectivityException(string message) : base(message) { }
        public OutlookConnectivityException(string message, Exception inner) : base(message, inner) { }
        protected OutlookConnectivityException(SerializationInfo info, StreamingContext context) : base(info, context) { }
    }
}

