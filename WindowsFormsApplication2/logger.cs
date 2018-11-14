using System;
using System.Collections.Generic;
using System.Linq;

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

        internal List<LogRecord> messageBase = new List<LogRecord>();

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
                messageBase = new List<LogRecord>();
            }
        }

        public void Add(string _portName, DirectionType _direction, DateTime _time, byte[] _byteArray, string _signalPin, bool _mark)
        {
            if (_byteArray == null) _byteArray = new byte[0];
            LogRecord tmp = new LogRecord
            {
                dateTime = _time,
                portName = _portName,
                direction = _direction,
                message = _byteArray,
                signalPin = _signalPin,
                mark = _mark
            };
            lock (threadLock)
            {
                messageBase.Add(tmp);
            }
        }

        public void Clear()
        {
            lock (threadLock)
            {
                messageBase = new List<LogRecord>();
            }
        }

        public int QueueSize()
        {
            lock (threadLock)
            {
                return messageBase.Count();
            }
        }

        public List<LogRecord> GetLog()
        {
            lock (threadLock)
            {
                List<LogRecord> tmp = messageBase.ToList();
                messageBase.Clear();
                return tmp;
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
