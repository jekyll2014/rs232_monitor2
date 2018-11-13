using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RS232_monitor
{
    internal class Logger
    {
        private readonly object threadLock;

        public enum DirectionType
        {
            DataIn,
            DataOut,
            SignalIn,
            SignalOut,
            Error
        }

        public string[] DirectionMark = new string[]
        {
            "<<",
            ">>",
            "<!",
            ">!",
            "!!",
        };

        public struct LogRecord
        {
            public DateTime dateTime;
            public string portName;
            public DirectionType direction;
            public byte[] message;
            public string signalPin;
            public bool mark;
        }

        internal Queue<LogRecord> messageBase = new Queue<LogRecord>();

        //public Queue<logRecord> Logs { get { return messageBase; } }

        public Logger()
        {
            threadLock = new object();
            DirectionMark = new string[]
        {
            "<<",
            ">>",
            "<!",
            ">!",
            "!!",
        };
            lock (threadLock)
            {
                messageBase = new Queue<LogRecord>();
            }
        }

        public void Add(string portName, DirectionType signal, DateTime time, byte[] byteArray,string signalPin, bool mark)
        {
            LogRecord tmp = new LogRecord
            {
                dateTime = time,
                portName = portName,
                direction = signal,
                message = byteArray,
                mark = mark
            };
            lock (threadLock)
            {
                messageBase.Enqueue(tmp);
            }
        }

        public void Clear()
        {
            lock (threadLock)
            {
                messageBase = new Queue<LogRecord>();
            }
        }

        public int QueueSize()
        {
            lock (threadLock)
            {
                return messageBase.Count;
            }
        }

        public LogRecord GetNext()
        {
            lock (threadLock)
            {
                return messageBase.Dequeue();
            }
        }

        public LogRecord GetQueueElement(int n)
        {
            lock (threadLock)
            {
                return messageBase.ElementAt(n);
            }
        }
    }
}
