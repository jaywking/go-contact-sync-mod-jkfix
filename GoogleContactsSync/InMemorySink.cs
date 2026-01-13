using Serilog.Core;
using Serilog.Events;
using Serilog.Formatting;
using Serilog.Formatting.Display;
using System;
using System.IO;

namespace GoContactSyncMod
{
    class InMemorySink : ILogEventSink
    {
        readonly ITextFormatter _textFormatter = new MessageTemplateTextFormatter("{Timestamp:yyyy-MM-dd HH:mm:ss} | {Level} | {Message}\r\n");

        public delegate void LogHandler(string str);
        public event LogHandler OnLogReceived;

        public void Emit(LogEvent logEvent)
        {
            if (logEvent == null)
            {
                throw new ArgumentNullException(nameof(logEvent));
            }

            if (OnLogReceived != null)
            {
                var renderSpace = new StringWriter();
                _textFormatter.Format(logEvent, renderSpace);
                OnLogReceived.Invoke(renderSpace.ToString());
            }
        }
    }
}
