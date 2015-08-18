using System;
using System.Collections.Concurrent;
using System.Messaging;
using System.Runtime.InteropServices;
using System.Threading;
using NLog;

namespace ModbusExcel
{
    /// <summary>
    /// In-process (.DLL) Excel RTD server.
    /// </summary>
    /// 
    [Guid("86854ACF-9902-4888-8E57-46D66516CA6F"),
     ProgId("ModbusExcel.RTD"),
     ComVisible(true)]
    public partial class RTD : IRtdServer
    {
        /// <summary>
        ///   Number of CPU processsors as queried from environment
        /// </summary>
        private static readonly int NumProcs = Environment.ProcessorCount;

        /// <summary>
        ///   Used by _topics ConcurrentDictionary operations. CPU * 2
        /// </summary>
        private static readonly int ConcurrencyLevel = NumProcs*2;

        /// <summary>
        ///   Instance of NLog manger for logging
        /// </summary>
        private static readonly Logger Log = LogManager.GetLogger("logfile");

        /// <summary>
        ///   Callback object for timer passed to serverstart by Excel.
        /// </summary>
        public static IRTDUpdateEvent MCallback;

        /// <summary>
        ///   Store item disconnect topics so they can be handled outside of Excel DisconnectData() function or ignored if shuttingdown
        ///   Any time a topic is deleted or modified, Excel calls Disconnect(TopicID). For Modified topics, Excel will reinvoke ConnectData() with new
        ///   topic paramaters. If topics are deleted or modified while editing cell contents, we eventualy need to remove the items from our 
        ///   polling schedule.
        /// </summary>
        private readonly ConcurrentDictionary<int, bool> _disconnectList = new ConcurrentDictionary<int, bool>();

        private static MessageQueue _msmq;

        /// <summary>
        ///   ConcurrentDictionary containing list of read requests made on sockets. Each Async request is added to this dictionary and entries
        ///   are removed from it in the ReadAsync callback event. This is how we match async requests with async responses
        /// </summary>
        private readonly ConcurrentDictionary<ushort, SocketTopic> _requests =
            new ConcurrentDictionary<ushort, SocketTopic>(ConcurrencyLevel, InitialCapacity);

        /// <summary>
        ///   ConcurrentDictionary containing list connections. Key can be either ip for single connection per device, or ip+|+register for single connection per register
        /// </summary>
        private readonly ConcurrentDictionary<string, SocketInfo> _sockets =
            new ConcurrentDictionary<string, SocketInfo>(ConcurrencyLevel, InitialCapacity);

        /// <summary>
        ///   ConcurrentDictionary containing list of poll requests (topics). Key=ip+|+register
        /// </summary>
        private readonly ConcurrentDictionary<string, TopicInfo> _topics =
            new ConcurrentDictionary<string, TopicInfo>(ConcurrencyLevel, InitialCapacity);

        /// <summary>
        ///   Stores time of last update sent to Excel.  We will not ask to update Excel any faster than the PollRefreshRate interval even though
        ///   we may have new results available quicket than that.  Also used to allow initial update to Excel even though server has not been started
        /// </summary>
        private double _lastupdate;

        /// <summary>
        ///   One-Shot Timer to poll devices.  On-shot means it never fires on a preset interval.  It fires one-time only when explicitly changed to invoke.
        ///   The timer callback always schedules the next callback before exiting.
        /// </summary>
        private Timer _modbusscantimer;

        /// <summary>
        ///   Sequence number for each ModbusScan. Used only for debugging info to track progress and events in timer callback
        /// </summary>
        private int _scannumber;

        /// <summary>
        ///   Sequence number used to uniquely identify each Modbus read request. Rolls over back to 1 at 64,000. (not zero because receiver ignores 0 transaction ids as keepalives)
        /// </summary>
        private ushort _sequence = 0;

        /// <summary>
        ///   Set when Excel terminates service. Allows us to skip deleting topics when shutting down
        /// </summary>
        private bool _serverterminating;

        /// <summary>
        ///   Gets set to true if data has changed. Allows us to notify Excel only when new results are available.  
        ///   More efficient than repeatedly updateing Excel with unchanged values
        /// </summary>
        private bool _updateavailable;

        private const String rtdName = "ModbusExcel.RTD";

        /// <summary>
        ///   MSMQ Queue name to send raw results to if LogLEvel && ResultsToMSMQ
        /// </summary>
        private const string QueueName = "lnd1042kdz\\private$\\kastest";

        /// <summary>
        ///   Size of initial _topics ConcurrentDictionary, and number of Socket Connections
        /// </summary>
        private const int InitialCapacity = 40000;

        /// <summary>
        ///   Default Modbus port.
        /// </summary>
        private const ushort Defaultport = 1502;

        /// <summary>
        ///   Stores information about each socket, along with a Master Modbus polling object
        /// </summary>
        private class SocketInfo
        {
            private readonly ModbusExcel.SocketInfo _socketInfo = new ModbusExcel.SocketInfo();

            public ModbusExcel.SocketInfo SocketInfo1
            {
                get { return _socketInfo; }
            }
        }

        /// <summary>
        ///   Stores a Socket Topic Pair used to match read reqeusts with async results.
        /// </summary>
        private class SocketTopic
        {
            public string SocketKey;
            public string TopicKey;
        }

        /// <summary>
        ///   Stores elements specific to a poll topic request (ip+register+unit+modicon)
        /// </summary>
        private class TopicInfo
        {
            private readonly ModbusExcel.TopicInfo _topicInfo = new ModbusExcel.TopicInfo();

            public ModbusExcel.TopicInfo TopicInfo1
            {
                get { return _topicInfo; }
            }
        }
    }
}