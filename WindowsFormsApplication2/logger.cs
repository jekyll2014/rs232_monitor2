using System;
using System.Collections.Generic;
using System.Linq;

namespace RS232_monitor2
{
    internal class Logger
    {
        private readonly object o_threadLock;

        public enum DirectionType
        {
            DataIn,
            DataOut,
            SignalIn,
            SignalOut,
            Error
        }

        public string[] DirectionMark;

        public struct LogRecord
        {
            public DateTime dateTime;
            public string portName;
            public DirectionType direction;
            public byte[] message;
            public string signalPin;
            public bool mark;
        }

        private List<LogRecord> messageBase;

        public Logger()
        {
            o_threadLock = new object();
            DirectionMark = new[]
            {
                "<<",
                ">>",
                "<!",
                ">!",
                "!!"
            };
            lock (o_threadLock)
            {
                messageBase = new List<LogRecord>();
            }
        }

        public void Add(string _portName, DirectionType _direction, DateTime _time, byte[] _byteArray,
            string _signalPin, bool _mark)
        {
            if (_byteArray == null) _byteArray = new byte[0];
            var tmp = new LogRecord
            {
                dateTime = _time,
                portName = _portName,
                direction = _direction,
                message = _byteArray,
                signalPin = _signalPin,
                mark = _mark
            };
            lock (o_threadLock)
            {
                messageBase.Add(tmp);
            }
        }

        public void Clear()
        {
            lock (o_threadLock)
            {
                messageBase = new List<LogRecord>();
            }
        }

        public int QueueSize()
        {
            lock (o_threadLock)
            {
                return messageBase.Count();
            }
        }

        public List<LogRecord> GetLog()
        {
            lock (o_threadLock)
            {
                var tmp = messageBase.ToList();
                messageBase.Clear();
                return tmp;
            }
        }

        public LogRecord GetQueueElement(int n)
        {
            lock (o_threadLock)
            {
                return messageBase.ElementAt(n);
            }
        }
    }
}
