using System;
using System.Collections.Concurrent;
using System.Globalization;
using System.Linq;
using System.Messaging;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using NLog;

//using PerformanceCounterHelper;

//using PerformanceCounterHelper;

namespace ModbusExcel
{
    /// <summary>
    ///   RTD server.
    /// </summary>
// ReSharper disable InconsistentNaming
    public partial class RTD : IDisposable
// ReSharper restore InconsistentNaming
    {

      //  private static CounterHelper<SingleInstance_PerformanceCounters> counterHelper;

        //public RTD()
        //{
        ////    counterHelper = PerformanceHelper.CreateCounterHelper<SingleInstance_PerformanceCounters>();
        //    //if (counterHelper == null)
        //    //{
        //    //    PerformanceHelper.Install(typeof(SingleInstance_PerformanceCounters));
        //    //}
        //}

        #region Fields

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

        #region Nested type: CfgInfo

        /// <summary>
        ///   Stores user configurable values for poller.  Even though the client (Excel) may initalize these values, they may
        ///   not get assigned before other data collection topics have been registered and polling has potentially started, so we assign 
        ///   default values that will be used prior to receiving new values for them.
        /// </summary>
        private static class CfgInfo
        {
            /// <summary>
            ///   Rate in ms for scheduling timer.
            /// </summary>
            public static int IdleTimerInterval = 2000;

            /// <summary>
            ///   ms delay before invoking timer callback if we have queued any async Connect or Read requests.
            /// </summary>
            public static int BusyTimerInterval = 100;

            /// <summary>
            ///   Rate in ms for polling modbus devices -1 = disabled, 0 = once, nn=ms  =RTD("modbusexcel.rtd",,"PollRate","2000",,,)
            /// </summary>
            public static int PollRate = -1;

            public static int PollRateTopicID = -1;

            /// <summary>
            ///   Socket Timeout ms  =RTD("modbusexcel.rtd",,"SocketTimeout","10000",,,)
            /// </summary>
            public static int SocketTimeout = 180000;

            public static int SocketTimeoutTopicID = -1;

            /// <summary>
            ///   Rate in ms for Excel to call RDT.Update(). =RTD("modbusexcel.rtd",,"ExcelUpdateRate","2000",,,) 
            ///   Not implemented/managed in this RTD Server. Actual usage is HeartBeat() call from VBA. 
            ///   Caution: Changing this permanently changes for this and all future Excell Sessions
            /// </summary>
            public static int ExcelUpdateRate = 5000;

            public static int ExcelUpdateRateTopicID = -1;

            /// <summary>
            ///  Shuold 
            /// Should Excel delete previously cached values rather than wait for retrieve of next value? =RTD("modbusexcel.rtd",,"DeleteCachedValues","True",,,)
            /// </summary>
            public static bool DeleteCachedValues = true;

            public static int DeleteCachedValuesTopicID = -1;

            /// <summary>
            ///   DebugLevel 0 = None(default) [=RTD("modbusexcel.rtd",,"DebugLevel","-1",,,)]
            ///   DCBA987654321 Bits
            ///   16384 0            = Synchronous Read Requests + Asnc Connect
            ///   8192  0            = Connect Status
            ///   4096   0           = Publish Raw data to MSMQ
            ///   2048    0          = Errors
            ///   1024     0         = RefreshData()
            ///   512       0        = Hex Output    
            ///   256        0       = StateInfo+Stats+Data
            ///   128         0      = StateInfo+Stats
            ///   64          0      = StateInfo+Data
            ///   32           0     = Read Request
            ///   16            0    = Response Received
            ///   8             0    = Connect/Disconnect
            ///   4              0   = Server Start/Stop
            ///   2               0  = Other Topic Register
            ///   1                0 = Config Topic Register
            /// 
            /// </summary>
            public static int DebugLevel = 0; //63;

            public static int DebugLevelTopicID = -1;

            /// <summary>
            ///   TopicCount =RTD("modbusexcel.rtd",,"TopicCount",,,,,)
            /// </summary>
            public static int TopicCount;

            public static int TopicCountTopicID = -1;

            /// <summary>
            ///   ConnectionCount =RTD("modbusexcel.rtd",,"ConnectionCount",,,,,)
            /// </summary>
            public static int ConnectionCount;

            public static int ConnectionCountTopicID = -1;

            /// <summary>
            ///   ConnectionPending =RTD("modbusexcel.rtd",,"ConnectionPending",,,,,)
            /// </summary>
            public static int ConnectionPending;

            public static int ConnectionPendingTopicID = -1;

            /// <summary>
            ///   ReadPending =RTD("modbusexcel.rtd",,"ReadPending",,,,,)
            /// </summary>
            public static int ReadPending;

            public static int ReadPendingTopicID = -1;

            /// <summary>
            ///   ServerRunning =RTD("modbusexcel.rtd",,"ServerRunning",,,,,)
            /// </summary>
            public static bool ServerRunning = false;

            public static int ServerRunningTopicID = -1;

            /// <summary>
            ///   Stores count of modbus read requests =RTD("modbusexcel.rtd",,"SendReceiveCount",,,,,)
            /// </summary>
            public static int SendCount;

            /// <summary>
            ///   Stores count of modbus responses =RTD("modbusexcel.rtd",,"SendReceiveCount",,,,,)
            /// </summary>
            public static int ReceiveCount;

            public static int SendReceiveCountTopicID = -1;

            /// <summary>
            ///   ServerDisable =RTD("modbusexcel.rtd",,"ServerDisable","true/false",,,,)
            /// </summary>
            public static bool ServerDisable = true; // default disabled until client enables us

            public static int ServerDisableTopicID = -1;

            /// <summary>
            ///   limit outstanding async connect requests
            /// </summary>
            public static int Maxqueue = 0;

            public static int MaxqueueTopicID = -1;

            /// <summary>
            ///   Max outsanding async read requests
            /// </summary>
            public static int Maxread = 0;

            public static int MaxreadTopicID = -1;

            /// <summary>
            ///   Minimum connections before read allowed
            /// </summary>
            public static int Minconnections = 1;

            public static int MinconnectionsTopicID = -1;

            /// <summary>
            ///   0 = no, +n = value, -n=(connect)
            /// </summary>
            public static int Simulate = 0;

            public static int SimulateTopicID = -1;

            public static int SimulateValue = 0;


        }

        #endregion

        #region Nested type: EventStates

        /// <summary>
        ///   Socket and Topic States enums
        /// </summary>
        private enum EventStates
        {
            None,
            New,
            Opening,
            Idle,
            Reading,
            Exception
        }

        #endregion

        #region Nested type: LogLevel

        /// <summary>
        ///   Debug Log Level Bitmasks
        /// </summary>
        private static class LogLevel
        {
            // 16385 Set All, 7229 
            public static int SyncRead = 16384;
            public static int NotException = 8192;
            public static int ResultsToMSMQ = 4096;
            public static int Errors = 2048;
            public static int RefreshData = 1024;
            public static int HexOutput = 512;
            public static int StateInfoStatsData = 256;
            public static int StateInfoStats = 128;
            public static int StateInfoData = 64;
            public static int ReadRequest = 32;
            public static int ResponseReceived = 16;
            public static int ConnectDisconnect = 8;
            public static int ServerStartStop = 4;
            public static int OtherTopicRegister = 2;
            public static int ConfigTopicRegister = 1;
            public static int None = 0;
        }

        #endregion

        #region Nested type: SocketInfo

        /// <summary>
        ///   Stores information about each socket, along with a Master Modbus polling object
        /// </summary>
        private class SocketInfo
        {
            /// <summary>
            ///   ms time the event began. used to detect timeouts
            /// </summary>
            public double Eventbegin;

            /// <summary>
            ///   IP of device
            /// </summary>
            public string Ip;

            /// <summary>
            ///   Enumberated list of event states for socket (New, Opening, Reading, Idle)
            /// </summary>
            public EventStates Lastevent;

            /// <summary>
            ///   Modbus Master polling object. Handles all Socket IO for this socket
            /// </summary>
            public Master Mbmaster;

            /// <summary>
            ///   Modicon Flag = Y/N/R (Y=Omni, N=Standard, S=Serial RTU, R=All Raw Results, D=Disabled)
            /// </summary>
            public string Modicon;

            /// <summary>
            ///   Port to connect to
            /// </summary>
            public ushort Port;

            /// <summary>
            ///   Stores number of outstanding/pending read requests for this socket. Socket is idle when this reaches zero.
            /// </summary>
            public int QueueDepth;

            /// <summary>
            ///   Number of topics sharing this connection.  When it reaches zero, the connection can be closed or reused
            /// </summary>
            public int References;

            /// <summary>
            ///   Modbus Unit ID
            /// </summary>
            public byte Unit;
        }

        #endregion

        #region Nested type: SocketTopic

        /// <summary>
        ///   Stores a Socket Topic Pair used to match read reqeusts with async results.
        /// </summary>
        private class SocketTopic
        {
            public string SocketKey;
            public string TopicKey;
        }

        #endregion

        #region Nested type: TopicInfo

        /// <summary>
        ///   Stores elements specific to a poll topic request (ip+register+unit+modicon)
        /// </summary>
        private class TopicInfo
        {
            /// <summary>
            ///   Modbus Register
            /// </summary>
            public ushort Addr;

            /// <summary>
            ///   PerRegister uses seperate connection for each ip+register pair. PerDevice uses shared connection for all registers per device
            /// </summary>
            public string ConnectionType;

            /// <summary>
            ///   String representation of polled register result
            /// </summary>
            public string Currentvalue;

            /// <summary>
            ///   Datatype for selected register. Calculated by Conversion class based on register address
            /// </summary>
            public string Datatype;

            /// <summary>
            ///   ms time at which the current event began. Used to detect timeouts, refresh due,
            /// </summary>
            public double Eventbegin;

            /// <summary>
            ///   IP address of device being polled
            /// </summary>
            public string Ip;

            /// <summary>
            ///   Enumerated states possible for any connection.
            /// </summary>
            public EventStates Lastevent;

            /// <summary>
            ///   Number of registers to read
            /// </summary>
            public byte Length;

            /// <summary>
            ///   Modicon Flag = Y/N. Determines when poller substracts 1 from register addresses, swaps words of floating point results, or increases request length
            /// </summary>
            public string Modicon;

            /// <summary>
            ///   Modbus Port Number. Currently takes on value from DefaultPort. Not user configurable at run-time
            /// </summary>
            public ushort Port;

            /// <summary>
            ///   Raw Modbus response data
            /// </summary>
            public byte[] Rawdata;

            /// <summary>
            ///   Number of poll responses received for this topic.  Used to
            /// </summary>
            public double ReceiveCount;

            /// <summary>
            ///   How frequently data is to be refreshed. Currently takes on value from globel PollRate.  Not user configurable per-device
            /// </summary>
            public double Refreshrate;

            /// <summary>
            ///   Number of read requests made for this topic.
            /// </summary>
            public double SendCount;

            /// <summary>
            ///   Unique number generated and provided by Excel.  Unique number for each RTD() request.
            /// </summary>
            public int Topicid;

            /// <summary>
            ///   Modbus Unit ID
            /// </summary>
            public byte Unit;
        }

        #endregion

        #endregion // Fields

        #region ModbusScan

        /// <summary>
        ///   Callback invoked by one-shot timer.  Every IdleTimerInterval (2000 ms), or quicker (10 ms) if we are expecting async Modbus read or connect responses
        /// </summary>
        private void ModbusScan()
        {
            // Stop timer 
            //_modbusscantimer.Change( Timeout.Infinite, Timeout.Infinite);

            int callbackinterval = CfgInfo.IdleTimerInterval;
            
            if (!_serverterminating && _disconnectList.Count > 0)
                DisconnectTopics();

            if (CfgInfo.ServerDisable)
            {
                // If server is stopped, force UpdateNotify every 10 seconds or thereabouts
                var random = new Random();
                int randomNumber = random.Next(0, 10);
                if (randomNumber >= 9 || _lastupdate == 0)
                    MCallback.UpdateNotify();
                // Reschedule callback 
                _modbusscantimer.Change(callbackinterval, Timeout.Infinite);
                return;
            }

            // Incremented each time ModbusScan is invoked. Used for informational/debug info only.
            _scannumber++;

            // Informational/debug item only. How many sockets are detected idle
            int idlecount = 0;
            // Informational/debug item only. How many Modbus read requests were queued this scan
            int triggerread = 0;

            bool quickcallback = false;
            byte[] syncBuffer = new byte[250];
            // Try reconnecting in the event of timeout.
            // foreach() operates on a snapshot of keys of _sockets concurrent dictionary.  Items deleted or added will not appear
            // in the enumeration of items, but other threads changes to <.,socketinfo> will be allowed 
            // 10-25-13 for debugging enforce orderby

            #region foreachsockets

            double now;
            foreach (var s in _sockets) //.OrderBy(p => p.Key))
            {
                now = Now();

                // If this is a new socket or this socket appears to have timed out and it has non-zero references we must reconnect
                if ((s.Value.Lastevent == EventStates.New ||
                     (now - s.Value.Eventbegin > CfgInfo.SocketTimeout && s.Value.References > 0)) &&
                    (
                    ( CfgInfo.ConnectionPending < CfgInfo.Maxqueue || CfgInfo.Maxqueue == -1)
                    || CfgInfo.Maxqueue == 1))
                {
                    #region connectsocket

                    if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                        Log.Info(
                            "Socket BeginConnect: Key:{0}, ip:{1}, Port:{2}, Unit:{3}, Modicon:{4}, References:{5}, LastEvent:{6}, Elapsed:{7}, Connections:{8}, ConnectionsPending:{9}, ConnectionsPending:{10}, ScanNumber:{11}",
                            s.Key, s.Value.Ip, s.Value.Port, s.Value.Unit, s.Value.Modicon, s.Value.References,
                            s.Value.Lastevent, now - s.Value.Eventbegin, CfgInfo.ConnectionCount,
                            CfgInfo.ConnectionPending, CfgInfo.ReadPending, _scannumber);

                    // If we timed out, decrease connection count
                    if (s.Value.Lastevent != EventStates.New && CfgInfo.ConnectionCount > 0)
                    {
                        CfgInfo.ConnectionCount--;
                        CfgInfo.ConnectionPending--;
                    }

                    // Change socket state to Opening and call ConnectAsync()
                    s.Value.Eventbegin = now;
                    s.Value.Lastevent = EventStates.Opening;
                    // If Connection failed to be QUEUED (not necessarilly connected) set state back to new and we will try again

                    if (CfgInfo.Maxqueue == -1)
                    {
                        if (s.Value.Mbmaster.ConnectSync(s.Value.Ip, s.Value.Port))
                        {
                            CfgInfo.ConnectionCount++;
                            s.Value.Eventbegin = Now();
                            s.Value.Lastevent = EventStates.Idle;
                            // Yield every 1000 sync connections to let Excel have some CPU cycles or it will think we died
                            if ((CfgInfo.ConnectionCount)%1000 == 0)
                            {
                                // Let Excel refresh from our data every 1000 connections
                                if ((CfgInfo.ConnectionCount)%1000 == 0)
                                    MCallback.UpdateNotify();

                                // Give Excel more time when we have more connections
                                _modbusscantimer.Change(CfgInfo.ConnectionCount/10, Timeout.Infinite);
                                return;
                            }
                        }
                    }
                    else
                    {
                        if (!s.Value.Mbmaster.ConnectAsync(s.Value.Ip, s.Value.Port))
                            s.Value.Lastevent = EventStates.New;
                        CfgInfo.ConnectionPending++;
                    }

                    #endregion connectsocket
                }
                // If last socket event was Opening, and Mbmaster indicates socket is now connected set socket state to Idle
                if (s.Value.Lastevent == EventStates.Opening && s.Value.Mbmaster.Connected)
                {
                    s.Value.Eventbegin = now;
                    s.Value.Lastevent = EventStates.Idle;
                }

            }


            #region readfromsocket
          
            now = Now();

            // This returns a list of registers and if possible, a useable/free socket, where prior read request has timed out or where it succeeded
            // and idletime exceeds PollRate, meaning it needs to be retrieved again (or for the first and only time if PollRate = 0)

            // Sync Fix

            if ((CfgInfo.ReadPending < CfgInfo.Maxread || CfgInfo.Maxread == -1) && (CfgInfo.ConnectionCount >= CfgInfo.Minconnections || CfgInfo.Minconnections == -1))
            {

               
                #region foreachRtopicinQsocket

                double now1 = now;

                // TODO: 8-10-15 Union is not efficient. All connection types will be the same so figure out what it is and only run the Linq for the needed one.
                // Some dead-code dated 10-27-13 from earlier attempts deleted. Search for and restore it as a starting point or just start over.
                // perdevice 
                var qd = (from sd in _sockets
                          join td in _topics on new { ip = sd.Value.Ip } equals new { ip = td.Value.Ip }
                          where td.Value.Ip == sd.Value.Ip && sd.Value.Mbmaster.Connected && td.Value.ConnectionType == "perdevice"
                                &&
                                ((td.Value.Lastevent == EventStates.Idle &&
                                  now1 - td.Value.Eventbegin > CfgInfo.PollRate) ||
                                 now1 - td.Value.Eventbegin > CfgInfo.SocketTimeout) &&
                                (sd.Value.Lastevent != EventStates.New && sd.Value.Lastevent != EventStates.Opening) &&
                                (!(now1 - td.Value.Eventbegin < CfgInfo.PollRate) &&
                                 (CfgInfo.PollRate != 0 || !(td.Value.ReceiveCount > 0)) && CfgInfo.PollRate != -1)
                          select new { entry = td, socket = sd.Key })//.OrderBy(x => Guid.NewGuid());
                    .Union(from sd in _sockets
                             join td in _topics on new { ip = sd.Key } equals new { ip = td.Key }
                             where td.Key == sd.Key && sd.Value.Mbmaster.Connected && td.Value.ConnectionType == "perregister"
                                   &&
                                   ((td.Value.Lastevent == EventStates.Idle &&
                                     now1 - td.Value.Eventbegin > CfgInfo.PollRate) ||
                                    now1 - td.Value.Eventbegin > CfgInfo.SocketTimeout) &&
                                   (sd.Value.Lastevent != EventStates.New && sd.Value.Lastevent != EventStates.Opening) &&
                                   (!(now1 - td.Value.Eventbegin < CfgInfo.PollRate) &&
                                    (CfgInfo.PollRate != 0 || !(td.Value.ReceiveCount > 0)) && CfgInfo.PollRate != -1)
                             select new { entry = td, socket = sd.Key });//.OrderBy(x => Guid.NewGuid());

                foreach (var r in qd)
                {
                    // Break out of loop if we are exceeding connection or read limits
                    if (CfgInfo.ReadPending >= CfgInfo.Maxread && CfgInfo.Maxread != -1 || CfgInfo.ConnectionCount < CfgInfo.Minconnections)
                        break;

                    now = Now();

                    string freesocketkey = r.socket;

                    // If we did not find a useable socket for this topic/register, we move on to the next one.  Socket may not be free
                    // if it is New or Opening
                    if (freesocketkey == null) continue;


                    idlecount++;

                    triggerread++;

                    // Request Number, Recycle at 64000 because ReadRegister() accepts ushort unique request id.  
                    // Is may be possible to accumulate unreceived requests to fill up pending _requests dictionary.
                    // It would need to be purged.  If timeouts are managed correctly it should be possible to purge _requests of old entries
                    // but receivedata would still need to gracefully handle delayed response for orphaned request
                    _sequence = (ushort) (_sequence > 64000 ? 1 : (ushort) (_sequence + 1));

                    CfgInfo.SendCount++;
                    _topics[r.entry.Key].Eventbegin = now;
                    _topics[r.entry.Key].SendCount++;
                    _sockets[freesocketkey].Eventbegin = now;
                    _requests[_sequence] = new SocketTopic {SocketKey = freesocketkey, TopicKey = r.entry.Key};
                    _topics[r.entry.Key].Lastevent = EventStates.Reading;
                    _sockets[freesocketkey].Lastevent = EventStates.Reading;

                    if (CfgInfo.Simulate != 0)
                    {
                        #region simulate

                        if (CfgInfo.Simulate < 0)
                        {
                            var data = new byte[1];
                            data[0] = _sockets[freesocketkey].Mbmaster.Connected ? (byte) 1 : (byte) 0;
                            syncBuffer = data;
                        }
                        else if (CfgInfo.Simulate > 0)
                        {
                            CfgInfo.SimulateValue = CfgInfo.SimulateValue + 1;
                            var data = new byte[1];
                            data[0] = CfgInfo.SimulateValue%2 == 0 ? (byte) 0 : (byte) 1;
                            syncBuffer = data;
                            //SyncBuffer = BitConverter.GetBytes(CfgInfo.SimulateValue);
                        }

                        CfgInfo.ReadPending++;
                        _sockets[freesocketkey].QueueDepth++;
                        if (r.entry.Value.Datatype == "T")
                            MBmaster_OnResponseData(_sequence, r.entry.Value.Unit, 0x65, syncBuffer);
                        else if (r.entry.Value.Addr < 3000 && r.entry.Value.Addr >= 1000)
                            MBmaster_OnResponseData(_sequence, r.entry.Value.Unit, 0x01, syncBuffer);
                        else
                            MBmaster_OnResponseData(_sequence, r.entry.Value.Unit, 0x03, syncBuffer);

                        #endregion simulate
                    }
                    else
                    {
                        // "T" datatype is unique to Omni. Used to read ASCII text buffers.  I.e. Prover report. Few if any systems 
                        // other than OmniCom support this datatype.
                        // TODO: For Text data try to read 19 128 byte blocks. Need to program a stop once HEX 1A is seen and read more than 19 if needed
                        if (r.entry.Value.Datatype == "T")
                        {
                            #region AsciiReport

                            if ((CfgInfo.DebugLevel & LogLevel.ReadRequest) == LogLevel.ReadRequest)
                                Log.Info(
                                    "ModbusScan ReadAsciiTextBuffer topicCount:{0} idlecount:{1} triggerread:{2} _Sequence:{3}, Unit:{4}, Modicon:{5}, Addr:{6}|{7}, Length:{8}|{9}, Freesocket:{10}, DataType:{11}, References:{12}, SocketState:{13}, ScanNumber:{14}, QueueDepth:{15}",
                                    CfgInfo.TopicCount, idlecount, triggerread, _sequence, r.entry.Value.Unit,
                                    r.entry.Value.Modicon,
                                    r.entry.Value.Addr,
                                    (ushort)
                                    (r.entry.Value.Addr +
                                     (r.entry.Value.Modicon == "n" || r.entry.Value.Modicon == "s" ? 0 : -1)),
                                    r.entry.Value.Length,
                                    (ushort)
                                    (r.entry.Value.Length),
                                    freesocketkey, r.entry.Value.Datatype, _sockets[freesocketkey].References,
                                    _sockets[freesocketkey].Lastevent, _scannumber,
                                    _sockets[freesocketkey].QueueDepth);

                            _topics[r.entry.Key].Lastevent = EventStates.Reading;
                            _sockets[freesocketkey].Lastevent = EventStates.Reading;
                            _sockets[freesocketkey].QueueDepth++;

                            // TODO: Expects length of 1 in reqest. Put code in place to always enforce that for Text datatypes.
                            //a) Remove attempt to read entire text buffer for now. Figure out how to implemnt in lower level later
                            //a)                      for (int i = 0; i < 1; i++)
                            //a)                      {
                            CfgInfo.ReadPending++;
                            _sockets[freesocketkey].Mbmaster.ReadAsciiTextBuffer(_sequence, r.entry.Value.Unit,
                                                                                 (ushort)
                                                                                 (r.entry.Value.Addr +
                                                                                  (r.entry.Value.Modicon == "n" ||
                                                                                   r.entry.Value.Modicon == "s"
                                                                                       ? 0
                                                                                       : -1)),
                                                                                 (ushort) (r.entry.Value.Length));

                            #endregion AsciiReport
                        }
                        else
                        {
                            #region savelog

                            if ((CfgInfo.DebugLevel & LogLevel.ReadRequest) == LogLevel.ReadRequest)
                                Log.Info(
                                    "ModbusScan ReadRegister        topicCount:{0} idlecount:{1} triggerread:{2} _Sequence:{3}, Unit:{4}, Modicon:{5}, Addr:{6}|{7}, Length:{8}|{9}, Freesocket:{10}, DataType:{11}, References:{12}, SocketState:{13}, ScanNumber{14}, QueueDepth:{15}, RegType:{16}",
                                    CfgInfo.TopicCount, idlecount, triggerread, _sequence, r.entry.Value.Unit,
                                    r.entry.Value.Modicon,
                                    r.entry.Value.Addr,
                                    (ushort)
                                    (r.entry.Value.Addr +
                                     (r.entry.Value.Modicon == "n" || r.entry.Value.Modicon == "s" ? 0 : -1)),
                                    r.entry.Value.Length,
                                    (ushort)
                                    (r.entry.Value.Length*
                                     (r.entry.Value.Modicon == "n" || r.entry.Value.Modicon == "s"
                                          ? 1
                                          : r.entry.Value.Datatype == "D" ||
                                            r.entry.Value.Datatype == "I"
                                                ? 1
                                                : r.entry.Value.Datatype == "F" ||
                                                  r.entry.Value.Datatype == "J"
                                                      ? 2
                                                      : 4)),
                                    freesocketkey, r.entry.Value.Datatype, _sockets[freesocketkey].References,
                                    _sockets[freesocketkey].Lastevent, _scannumber,
                                    _sockets[freesocketkey].QueueDepth,
                                    r.entry.Value.Addr < 3000 && r.entry.Value.Addr > 1000 ? "Coil" : "Holding");

                            #endregion savelog

                            // Rules for Modicon=Y/N register read requirements are implemented here

                            // Sync Read request.                             
                            if ((CfgInfo.DebugLevel & LogLevel.SyncRead) == LogLevel.SyncRead)
                            {
                                #region syncread

                                // Sync reads always reading or idle

                                _topics[r.entry.Key].Lastevent = EventStates.Idle;
                                _sockets[freesocketkey].Lastevent = EventStates.Idle;
                                _sockets[freesocketkey].QueueDepth++;

                                CfgInfo.ReadPending++;
                                if (r.entry.Value.Addr < 3000 && r.entry.Value.Addr >= 1000)
                                {
                                    _sockets[freesocketkey].Mbmaster.ReadCoils(_sequence, r.entry.Value.Unit,
                                                                               (ushort)
                                                                               (r.entry.Value.Addr +
                                                                                (r.entry.Value.Modicon == "n" ||
                                                                                 r.entry.Value.Modicon == "s"
                                                                                     ? 0
                                                                                     : -1)),
                                                                               (ushort)
                                                                               (r.entry.Value.Length*
                                                                                (r.entry.Value.Modicon == "n" ||
                                                                                 r.entry.Value.Modicon == "s"
                                                                                     ? 1
                                                                                     : r.entry.Value.Datatype ==
                                                                                       "D" ||
                                                                                       r.entry.Value.Datatype ==
                                                                                       "I"
                                                                                           ? 1
                                                                                           : r.entry.Value.
                                                                                               Datatype == "F" ||
                                                                                             r.entry.Value.
                                                                                               Datatype == "J"
                                                                                                 ? 2
                                                                                                 : 4))
                                                                               // r.entry.Value.Rawdata = buffer for sync response
                                                                               , r.entry.Value.Modicon,
                                                                               ref syncBuffer);
                                    if (syncBuffer == null)
                                        MBmaster_OnException(_sequence, r.entry.Value.Unit, 0x01, 2);
                                    else
                                        MBmaster_OnResponseData(_sequence, r.entry.Value.Unit, 0x01, syncBuffer);
                                }
                                else
                                {
                                    _sockets[freesocketkey].Mbmaster.ReadHoldingRegister(_sequence,
                                                                                         r.entry.Value.Unit,
                                                                                         (ushort)
                                                                                         (r.entry.Value.Addr +
                                                                                          (r.entry.Value.Modicon ==
                                                                                           "n" ||
                                                                                           r.entry.Value.Modicon ==
                                                                                           "s"
                                                                                               ? 0
                                                                                               : -1)),
                                                                                         (ushort)
                                                                                         (r.entry.Value.Length*
                                                                                          (r.entry.Value.Modicon ==
                                                                                           "n" ||
                                                                                           r.entry.Value.Modicon ==
                                                                                           "s"
                                                                                               ? 1
                                                                                               : r.entry.Value
                                                                                                  .Datatype ==
                                                                                                 "D" ||
                                                                                                 r.entry.Value
                                                                                                  .Datatype ==
                                                                                                 "I"
                                                                                                     ? 1
                                                                                                     : r.entry
                                                                                                        .Value
                                                                                                        .
                                                                                                         Datatype ==
                                                                                                       "F" ||
                                                                                                       r.entry
                                                                                                        .Value
                                                                                                        .
                                                                                                         Datatype ==
                                                                                                       "J"
                                                                                                           ? 2
                                                                                                           : 4))
                                                                                         // r.entry.Value.Rawdata = buffer for sync response
                                                                                         , r.entry.Value.Modicon,
                                                                                         ref syncBuffer);
                                    if (syncBuffer == null)
                                        MBmaster_OnException(_sequence, r.entry.Value.Unit, 0x03, 2);
                                    else
                                        MBmaster_OnResponseData(_sequence, r.entry.Value.Unit, 0x03, syncBuffer);
                                }
                            }
                                #endregion syncread
                                // Async read request
                            else
                            {
                                #region Asyncread

                                if (r.entry.Value.Addr < 3000 && r.entry.Value.Addr >= 1000)
                                    _sockets[freesocketkey].Mbmaster.ReadCoils(_sequence, r.entry.Value.Unit,
                                                                               (ushort)
                                                                               (r.entry.Value.Addr +
                                                                                (r.entry.Value.Modicon == "n" ||
                                                                                 r.entry.Value.Modicon == "s"
                                                                                     ? 0
                                                                                     : -1)),
                                                                               (ushort)
                                                                               (r.entry.Value.Length*
                                                                                (r.entry.Value.Modicon == "n" ||
                                                                                 r.entry.Value.Modicon == "s"
                                                                                     ? 1
                                                                                     : r.entry.Value.Datatype ==
                                                                                       "D" ||
                                                                                       r.entry.Value.Datatype ==
                                                                                       "I"
                                                                                           ? 1
                                                                                           : r.entry.Value.
                                                                                               Datatype == "F" ||
                                                                                             r.entry.Value.
                                                                                               Datatype == "J"
                                                                                                 ? 2
                                                                                                 : 4))
                                                                               , r.entry.Value.Modicon);

                                else
                                    _sockets[freesocketkey].Mbmaster.ReadHoldingRegister(_sequence,
                                                                                         r.entry.Value.Unit,
                                                                                         (ushort)
                                                                                         (r.entry.Value.Addr +
                                                                                          (r.entry.Value.Modicon ==
                                                                                           "n" ||
                                                                                           r.entry.Value.Modicon ==
                                                                                           "s"
                                                                                               ? 0
                                                                                               : -1)),
                                                                                         (ushort)
                                                                                         (r.entry.Value.Length*
                                                                                          (r.entry.Value.Modicon ==
                                                                                           "n" ||
                                                                                           r.entry.Value.Modicon ==
                                                                                           "s"
                                                                                               ? 1
                                                                                               : r.entry.Value
                                                                                                  .Datatype ==
                                                                                                 "D" ||
                                                                                                 r.entry.Value
                                                                                                  .Datatype ==
                                                                                                 "I"
                                                                                                     ? 1
                                                                                                     : r.entry
                                                                                                        .Value
                                                                                                        .
                                                                                                         Datatype ==
                                                                                                       "F" ||
                                                                                                       r.entry
                                                                                                        .Value
                                                                                                        .
                                                                                                         Datatype ==
                                                                                                       "J"
                                                                                                           ? 2
                                                                                                           : 4))
                                                                                         , r.entry.Value.Modicon);
                                CfgInfo.ReadPending++;
                                _topics[r.entry.Key].Lastevent = EventStates.Reading;
                                _sockets[freesocketkey].Lastevent = EventStates.Reading;
                                _sockets[freesocketkey].QueueDepth++;

                                #endregion Asyncread
                            }
                        }
                    }
                    // Reschedule one-shot timer to callback nearly immediatly if we have work to do
                    quickcallback = true;
                }
                #endregion  foreachRtopicinQsocket
            }
            #endregion readfromsocket            
            #endregion foreachsockets


            if (quickcallback)
                callbackinterval = CfgInfo.BusyTimerInterval;
            // If we have received a Modbus read response and more than PollRate ms have elapsed since we updated Excel, 
            // we notify Excel that an update is available
            if (_updateavailable && Now() - _lastupdate > CfgInfo.PollRate)
            {
                _lastupdate = Now();
                MCallback.UpdateNotify();
            }

            _modbusscantimer.Change(callbackinterval, Timeout.Infinite);
        }

        #endregion //

        /// <summary>
        ///   Called by modbusscan to disconnect any pending topic disconnects
        /// </summary>
        private void DisconnectTopics()
        {
            foreach (var r in _topics)
            {
                bool isDisconnected;
                _disconnectList.TryRemove(r.Value.Topicid, out isDisconnected);
                if (isDisconnected)
                {

                    var qd = from sd in _sockets
                             join td in _topics on new { ip = sd.Key } equals new { ip = td.Key }
                             where td.Value.Topicid == r.Value.Topicid && sd.Value.Mbmaster.Connected
                             select sd.Value;

                    foreach (var rd in qd)
                    {
                        if (rd.References > 0) rd.References--;
                        if (rd.References == 0)
                        {
                            SocketInfo socketinfo;
                            try
                            {
                                rd.Mbmaster.Disconnect();
                                if (CfgInfo.ConnectionCount > 0) CfgInfo.ConnectionCount--;
                            }
// ReSharper disable once EmptyGeneralCatchClause
                            catch
                            {
                            }
                            _sockets.TryRemove(rd.Ip, out socketinfo);
                            if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                                Log.Info(
                                    "Topic Disconnect. Socket Disconnecting: topicid {0}, ip {1}, port {2}, unit {3}, modicon {4}, lastevent {5}",
                                    r.Value.Topicid, rd.Ip, rd.Port, rd.Unit, rd.Modicon, rd.Lastevent);
                            break; // Only disconnect last connection for this topic
                        }
                        else
                        {
                            if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                                Log.Info(
                                    "Topic Disconnect. Socket Still In Use: topicid {0}, ip {1}, port {2}, unit {3}, modicon {4}, lastevent {5}, references {6}",
                                    r.Value.Topicid, rd.Ip, rd.Port, rd.Unit, rd.Modicon, rd.Lastevent, rd.References);
                        }
                    }
                }
            }
        }

        /// <summary>
        ///   Returns millisecond of current day.  This will break when clock croses Midnight. Need to include days since 1900 or something similar
        /// </summary>
        /// <returns> </returns>
        public static double Now()
        {
            return 1000*
                   (DateTime.Now.Hour*3600 + DateTime.Now.Minute*60 + DateTime.Now.Second +
                    DateTime.Now.Millisecond*.001);
        }

        #region IRtdServer Interface

        /// <summary>
        ///   Called when Excel RTD connects to this automation component.
        /// </summary>
        /// <param name="callbackObject"> [In] This is the object through which this RTD server notifies Excel that real-time data updates are available. </param>
        /// <returns> 1 to denote "success". </returns>
        public int ServerStart(IRTDUpdateEvent callbackObject)
        {
// Uncomment to prevent Server Startup to allow fixup of problematic spreadsheet
//            return 1;

            MCallback = callbackObject;

            // Create an inactive one-shot timer with ModbusScan() as delegate
            _modbusscantimer = new Timer(delegate { ModbusScan(); }, null, Timeout.Infinite, Timeout.Infinite);
            // Now that timer is initialized, 
            // 8-8-15 Dont start it here. It will be started when client sends serverdisabled=false
            _modbusscantimer.Change(CfgInfo.IdleTimerInterval, Timeout.Infinite);

//            _lastupdate = Now() - CfgInfo.PollRate;

            if ((CfgInfo.DebugLevel & LogLevel.ServerStartStop) == LogLevel.ServerStartStop)
                Log.Info("ExcelRTD ServerStart Logger:{0}, Logfile:{1}", LogManager.GetCurrentClassLogger(),
                         LogManager.GetLogger("logfile"));

            return 1;
        }


        /// <summary>
        ///   Called when Excel RTD disconnects from this automation component.
        /// </summary>
        public void ServerTerminate()
        {
            if (_modbusscantimer != null) _modbusscantimer.Change(Timeout.Infinite, Timeout.Infinite);
            _serverterminating = true;
        }

        #region ConnectData

        /// <summary>
        ///   Called each time Excel RTD has a unique topic to request.
        /// </summary>
        /// <param name="topicId"> [In] Excel's internal topic ID. </param>
        /// <param name="strings"> [In] List of topics (parameters) passed by Excel RTD function call. </param>
        /// <param name="getNewValues"> [In/Out] On input the new_value parameter indicates whether Excel already has a value to initially display. So if Excel has a cached value to display then new_value will be false. If Excel does not have a cached value to display then new_value will be true, indicating that a new value is needed. On output the new_value parameter indicates whether Excel should use the returned value or not. If new_value is false then Excel will ignore the value returned by ConnectData(). If it doesn’t have a previously cached value then it displays the “#N/A” warning. Of course this is replaced with an actual value once it receives one via RefreshData(). If new_value is true then Excel will immediately replace whatever value it may already have with the value returned by ConnectData(). </param>
        /// <returns> Topic data value, if available. Otherwise, it returns a default message used to indicate that a computation is under way. Usually something like "RB.XL: Calculating ...". </returns>
        public object ConnectData(int topicId, ref object[] strings, ref bool getNewValues)
        {
            string p1;

            try
            {
                p1 = strings.GetValue(0).ToString().ToLower();
            }
            catch
            {
                return -1;
            }


            // If Parameter1 is a read-only configuration item rather than a poll request, save it's TopicID and return.  
            // We will merge Configuration values with poll results and return the combined list to Excel in RefreshData()
            switch (p1)
            {
                case "topiccount":
                    getNewValues = true;
                    CfgInfo.TopicCountTopicID = topicId;
                    CfgInfo.TopicCount = _topics.Count;
                    return (new object[,] {{topicId}, {CfgInfo.TopicCount}});
                case "connectioncount":
                    getNewValues = true;
                    CfgInfo.ConnectionCountTopicID = topicId;
                    CfgInfo.ConnectionCount = 0; // _sockets.Count;
                    return (new object[,] { { topicId }, { CfgInfo.ConnectionCount } });
                case "connectionpending":
                    getNewValues = true;
                    CfgInfo.ConnectionPendingTopicID = topicId;
                    CfgInfo.ConnectionPending = 0; // _sockets.Count;
                    return (new object[,] { { topicId }, { CfgInfo.ConnectionPending } });
                case "readpending":
                    getNewValues = true;
                    CfgInfo.ReadPendingTopicID = topicId;
                    CfgInfo.ReadPending = 0; 
                    return (new object[,] { { topicId }, { CfgInfo.ReadPending } });
                case "serverrunning":
                    getNewValues = true;
                    CfgInfo.ServerRunningTopicID = topicId;
                    CfgInfo.ServerRunning = !CfgInfo.ServerDisable;
                    return (new object[,] {{topicId}, {CfgInfo.ServerRunning.ToString()}});
                case "sendreceivecount":
                    getNewValues = true;
                    CfgInfo.SendReceiveCountTopicID = topicId;
 //                   CfgInfo.ReceiveCount = 0;
//                    CfgInfo.SendCount = 0;
                    return
                        (new object[,]
                             {
                                 {topicId},
                                 {
                                     CfgInfo.SendCount.ToString(CultureInfo.InvariantCulture) + "/" +
                                     CfgInfo.ReceiveCount.ToString(CultureInfo.InvariantCulture)
                                 }
                             });
            }


            // If Parameter1 is a user-configurable item, save setting provided in Parameter2, and store TopicID
            // We will merge Configuration values with poll results and return the combined list to Excel in RefreshData()
            if (p1 == "pollrate" || p1 == "excelupdaterate" || p1 == "deletecachedvalues" ||
                p1 == "sockettimeout" || p1 == "serverdisable" || p1 == "debuglevel" || p1 == "maxqueue" || p1 == "maxread" || p1 == "minconnections" || p1 == "simulate" )
                try
                {
                    string p2 = strings.GetValue(1).ToString().ToLower();
                    getNewValues = true;
                    if ((CfgInfo.DebugLevel & LogLevel.ConfigTopicRegister) == LogLevel.ConfigTopicRegister)
                        Log.Info("TopicMap ConfigChange TopicID:{0}, p1:{1}, p2:{2}, ", topicId, p1, p2);
                    switch (p1.ToLower())
                    {
                        case "pollrate":
                            CfgInfo.PollRate = Convert.ToInt16(p2);
                            // KAS Limit poll rate to minimum 1 second regardless of what user requests. Omni FC docs recommend no more than once every 10 seconds.
                            // if (CfgInfo.PollRate != -1 && CfgInfo.PollRate > 0) CfgInfo.PollRate = 1000;
                            CfgInfo.PollRateTopicID = topicId;
                            return (new object[,] {{topicId}, {CfgInfo.PollRate}});
                        case "sockettimeout":
                            CfgInfo.SocketTimeout = Convert.ToInt32(p2);
                            // Limit SocketTimeout to no less than 3 seconds regardless of what user requests. 
                            if (CfgInfo.SocketTimeout < 3000) CfgInfo.SocketTimeout = 3000;
                            CfgInfo.SocketTimeoutTopicID = topicId;
                            return (new object[,] {{topicId}, {CfgInfo.SocketTimeoutTopicID}});
                        case "excelupdaterate":
                            // Not used in this implementation. Requires Excel VBA code to set heartbeatrefresh
                            CfgInfo.ExcelUpdateRate = Convert.ToInt16(p2);
                            CfgInfo.ExcelUpdateRateTopicID = topicId;
                            return (new object[,] {{topicId}, {CfgInfo.ExcelUpdateRate}});
                        case "debuglevel":
                            CfgInfo.DebugLevel = Convert.ToInt32(p2);
                            CfgInfo.DebugLevelTopicID = topicId;
                            // Create connection to messaging to reuse later.
                            if ((CfgInfo.DebugLevel & LogLevel.ResultsToMSMQ) == LogLevel.ResultsToMSMQ && !CfgInfo.ServerDisable && _msmq == null)
                            {
                                _msmq = Messaging.GetConnection(QueueName);
                            }
                            return (new object[,] {{topicId}, {CfgInfo.DebugLevel}});
                        case "maxqueue":
                            CfgInfo.Maxqueue = Convert.ToInt32(p2);
                            CfgInfo.MaxqueueTopicID = topicId;
                            return (new object[,] {{topicId}, {CfgInfo.Maxqueue}});
                        case "maxread":
                            CfgInfo.Maxread = Convert.ToInt32(p2);
                            CfgInfo.MaxreadTopicID = topicId;
                            return (new object[,] {{topicId}, {CfgInfo.Maxread}});
                        case "minconnections":
                            CfgInfo.Minconnections = Convert.ToInt32(p2);
                            CfgInfo.MinconnectionsTopicID = topicId;
                            return (new object[,] { { topicId }, { CfgInfo.Minconnections } });
                        case "simulate":
                            CfgInfo.Simulate = Convert.ToInt16(p2);
                            CfgInfo.SimulateTopicID = topicId;
                            return (new object[,] { { topicId }, { CfgInfo.Simulate } });
                        case "deletecachedvalues":
                            CfgInfo.DeleteCachedValues = strings.GetValue(1).ToString().ToLower() != "false";
                            CfgInfo.DeleteCachedValuesTopicID = topicId;
                            return (new object[,] {{topicId}, {CfgInfo.DeleteCachedValues.ToString()}});
                        case "serverdisable":
                            // Turning Off
                            if (_sockets.Count > 0 && !CfgInfo.ServerDisable &&
                                strings.GetValue(1).ToString().ToLower() == "true")
                            {
                                foreach (var s in _sockets)
                                {
                                    if (s.Value.Mbmaster.Connected)
                                        try
                                        {
                                            s.Value.Mbmaster.Disconnect();
                                        }
// ReSharper disable once EmptyGeneralCatchClause
                                        catch
                                        {
                                        }
                                    s.Value.Lastevent = EventStates.New;
                                    s.Value.Eventbegin = Now();
                                }
                            }
                            // Turning On
                            if (_sockets.Count > 0 && CfgInfo.ServerDisable &&
                                strings.GetValue(1).ToString().ToLower() == "false")
                            {
                                // If we just enabled the server, set all EVentBegins to 100ms before regular pollrate so prevent appearance of timeouts for idle connections and topics
                                foreach (var s in _sockets)
                                {
                                    s.Value.Eventbegin = Now() - CfgInfo.PollRate + 100;
                                }
                                foreach (var s in _topics)
                                {
                                    s.Value.Eventbegin = Now() - CfgInfo.PollRate + 100;
                                }
                            }
                            CfgInfo.ServerDisable = strings.GetValue(1).ToString().ToLower() != "false";
                            CfgInfo.ServerDisableTopicID = topicId;
                            CfgInfo.ServerRunning = !CfgInfo.ServerDisable;
                            // If server is running, and maxqueue is zero, automatically change maxqueue to -1 so it will actually start processing requests
                            if (CfgInfo.ServerRunning && CfgInfo.Maxqueue == 0) CfgInfo.Maxqueue = -1;
                            // If server is running, and maxreqd is zero, automatically change maxread to -1 so it will actually start processing requests
                            if (CfgInfo.ServerRunning && CfgInfo.Maxread == 0) CfgInfo.Maxread = -1; 
                            _modbusscantimer.Change(CfgInfo.BusyTimerInterval, Timeout.Infinite);

                            // This will only disable polling after it is seen, and since it may not me the first topic connected, some topics may be polled anyway
                            return (new object[,] {{topicId}, {CfgInfo.ServerDisable.ToString()}});
                    }
                }
                catch
                {
                    return -1;
                }

            //  true  = tell excel to always get new values rather than show old cached values
            getNewValues = CfgInfo.DeleteCachedValues;

            TopicInfo ti = new TopicInfo();

            try
            {
                ti.Topicid = Convert.ToUInt16(topicId);
                ti.Ip = strings.GetValue(0).ToString().ToLower();
                ti.Unit = Convert.ToByte(strings.GetValue(1).ToString());
                ti.Modicon = Convert.ToString(strings.GetValue(2).ToString()).ToLower();
                ti.Addr = Convert.ToUInt16(strings.GetValue(3).ToString());
                ti.Length = Convert.ToByte(strings.GetValue(4).ToString());
                ti.Refreshrate = CfgInfo.PollRate; // Not used PER item right now. Global setting is PollRate and SocketTimeout
                ti.Port = Defaultport;
            }
            catch
            {
                return -1;
            }

            // Optional parameter, defaults to one connection per device
            try
            {
                ti.ConnectionType = Convert.ToString(strings.GetValue(5).ToString()).ToLower();
            }
            catch
            {
                ti.ConnectionType = "perdevice";
            }

 // Optional parameter, defaults to default port
            try
            {
                ti.Port = Convert.ToUInt16(strings.GetValue(6).ToString());
            }
            catch
            {
                try
                {
                    ti.Port = Convert.ToUInt16(strings.GetValue(5).ToString());
                }
                catch
                {
                    ti.Port = Defaultport;
                }
            }


            // Special Modicon = D setting disables polling for selected device
            if (ti.Modicon == "d") return (new object[,] {{topicId}, {"0"}});

            string topickey = ti.Ip + "|" + ti.Addr;
            TopicInfo priorti;

            _topics.TryGetValue(topickey, out priorti);
            if (priorti != null)
            {
                SocketInfo priortisi;
                string priorconnectkey = priorti.ConnectionType == "perregister"
                                             ? priorti.Ip + "|" + priorti.Addr
                                             : priorti.Ip;
                _sockets.TryGetValue(priorconnectkey, out priortisi);
                if (priortisi != null)
                {
                    // If references = 0 connection has already been disconnected by DisconnectData
                    if (priortisi.References > 0) priortisi.References--;
                    if (priortisi.References == 0)
                    {
                        SocketInfo removeconnectinfo;
                        if (!_sockets.TryRemove(priorconnectkey, out removeconnectinfo))
                        {
                            if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                                Log.Info("Remove Connection Failed:{0} Count:{1} References:{2}", priorconnectkey,
                                         _sockets.Count, priortisi.References);
                        }
                        else
                        {
                            if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                                Log.Info("Connection Removed:{0} Count:{1} References:{2}.", priorconnectkey,
                                         _sockets.Count, priortisi.References);
                        }
                    }
                    else
                    {
                        if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                            Log.Info("Connection Changed:{0} Count:{1} Old References:{2} New References:{3} ",
                                     priorconnectkey, _sockets.Count, priortisi.References + 1, priortisi.References);
                    }
                }
            }

            _topics[topickey] = ti;

            string connectkey = ti.ConnectionType == "perregister" ? ti.Ip + "|" + ti.Addr : ti.Ip;

            SocketInfo si = new SocketInfo();
            SocketInfo priorsi;

            si.Ip = ti.Ip;
            si.Port = ti.Port;
            si.Unit = ti.Unit;
            si.Modicon = ti.Modicon;
            si.References = 1;

            _sockets.TryGetValue(connectkey, out priorsi);
            if (priorsi != null)
            {
                // TODO: Remove from other processing queues before deleting? Whatever that means.
                si.References += priorsi.References;
                si.Lastevent = priorsi.Lastevent;
                si.Mbmaster = priorsi.Mbmaster;
                _sockets[connectkey] = si;
            }
            else
            {
                si.Lastevent = EventStates.New;
                si.Mbmaster = new Master();
                _sockets[connectkey] = si;
                si.Mbmaster.OnConnectComplete += MBmaster_OnConnectComplete;
                si.Mbmaster.OnResponseData += MBmaster_OnResponseData;
                si.Mbmaster.OnException += MBmaster_OnException;
            }

            if ((CfgInfo.DebugLevel & LogLevel.OtherTopicRegister) == LogLevel.OtherTopicRegister)
            {
                var q = from s in _sockets.Values
                        join t in _topics.Values on new {ip = s.Ip} equals new {ip = t.Ip}
                        where t.Topicid == ti.Topicid
                        // && s.Lastevent == "N" 
                        select
                            new
                                {
                                    topicid = t.Topicid,
                                    ip = s.Ip,
                                    port = t.Port,
                                    unit = t.Unit,
                                    modicon = t.Modicon,
                                    addr = t.Addr,
                                    length = t.Length,
                                    refreshrate = t.Refreshrate,
                                    lastevent = s.Lastevent,
                                    connectiontype = t.ConnectionType
                                };
                foreach (var r in q)
                {
                    Log.Info(
                        "TopicMap TopicID:{0}, ip:{1}, Port:{2}, Unit:{3}, Modicon:{4}, Addr:{5}, Length:{6}, RefreshRate:{7}, Socketcount:{8}, ConnectionType:{9}, Old References:{10}, New References:{11}",
                        r.topicid, r.ip, r.port, r.unit, r.modicon, r.addr, r.length, r.refreshrate, _sockets.Count,
                        r.connectiontype, priorsi != null ? priorsi.References : 0, si.References);
                }
            }

            ti.Currentvalue = "?";
            ti.Datatype = ti.Modicon == "r" ? "R" : Conversion.Datatype(ti.Addr);
            ti.Rawdata = null;
            ti.Lastevent = EventStates.Idle;
            //Force initial data to be stale and need a refresh
            ti.Eventbegin = Now() - CfgInfo.PollRate;

            return (new object[,] {{topicId}, {"0"}});
        }

        #endregion // ConnectData

        #region DisconnectData

        /// <summary>
        ///   Unsubscribe from a topic. This notification comes from
        ///   Excel when it no longer has any RTD() entries in cells
        ///   for the topic in question.  To prevent interfering with Excel UX, quickly save disconnect requests for
        ///   any Poll request topicId. Remove topics and unused sockets later when timer callback (ModbusScan()) invokes DisconnectData()
        /// </summary>
        /// <param name="topicId"> [In] Excel's internal topic ID. </param>
        public void DisconnectData(int topicId)
        {
            if (topicId != CfgInfo.TopicCountTopicID &&
                topicId != CfgInfo.ConnectionCountTopicID &&
                topicId != CfgInfo.ConnectionPendingTopicID &&
                topicId != CfgInfo.ReadPendingTopicID &&
                topicId != CfgInfo.ServerRunningTopicID &&
                topicId != CfgInfo.SendReceiveCountTopicID &&
                topicId != CfgInfo.PollRateTopicID &&
                topicId != CfgInfo.SocketTimeoutTopicID &&
                topicId != CfgInfo.ExcelUpdateRateTopicID &&
                topicId != CfgInfo.DebugLevelTopicID &&
                topicId != CfgInfo.MaxqueueTopicID &&
                topicId != CfgInfo.MaxreadTopicID &&
                topicId != CfgInfo.MinconnectionsTopicID &&
                topicId != CfgInfo.SimulateTopicID &&
                topicId != CfgInfo.DeleteCachedValuesTopicID &&
                topicId != CfgInfo.ServerDisableTopicID)
                _disconnectList[topicId] = true;
        }

        #endregion // DisconnectData

        #region RefreshData

        /// <summary>
        ///   Excel will call this method to get new data.
        /// </summary>
        /// <param name="topicCount"> [Out] Number of topics for which data is being returned. </param>
        /// <returns> Array of topic IDs and associated data. </returns>
        public object[,] RefreshData(ref int topicCount)
        {
            //Temp override to return no data to search for crash bug
            //topicCount = 0;
            //var nodata = new object[2, 0];
            //return nodata;

            int tc = 0;
            _lastupdate = Now();
            _updateavailable = false;
            if (CfgInfo.TopicCountTopicID != -1) tc++;
            if (CfgInfo.ConnectionCountTopicID != -1) tc++;
            if (CfgInfo.ConnectionPendingTopicID != -1) tc++;
            if (CfgInfo.ReadPendingTopicID != -1) tc++;
            if (CfgInfo.ServerRunningTopicID != -1) tc++;
            if (CfgInfo.SendReceiveCountTopicID != -1) tc++;
            if (CfgInfo.PollRateTopicID != -1) tc++;
            if (CfgInfo.SocketTimeoutTopicID != -1) tc++;
            if (CfgInfo.ExcelUpdateRateTopicID != -1) tc++;
            if (CfgInfo.DebugLevelTopicID != -1) tc++;
            if (CfgInfo.MaxqueueTopicID != -1) tc++;
            if (CfgInfo.MaxreadTopicID != -1) tc++;
            if (CfgInfo.MinconnectionsTopicID != -1) tc++;
            if (CfgInfo.SimulateTopicID != -1) tc++;
            if (CfgInfo.DeleteCachedValuesTopicID != -1) tc++;
            if (CfgInfo.ServerDisableTopicID != -1) tc++;

            CfgInfo.TopicCount = _topics.Count() + tc;
            var refdata = new object[2,CfgInfo.TopicCount];

            var index = 0;

            foreach (var entry in _topics)
            {
                refdata[0, index] = entry.Value.Topicid;
                refdata[1, index] = entry.Value.Currentvalue;
                // Convert.ToString(_topics[entry.Value.RRmyid].RRCurrentValue);
                ++index;
            }

            if (CfgInfo.TopicCountTopicID != -1)
            {
                refdata[0, index] = CfgInfo.TopicCountTopicID;
                refdata[1, index] = _topics.Count;
                ++index;
            }


            if (CfgInfo.ConnectionCountTopicID != -1)
            {
                refdata[0, index] = CfgInfo.ConnectionCountTopicID;
                refdata[1, index] = CfgInfo.ConnectionCount; //  _sockets.Count;
                ++index;
            }

            if (CfgInfo.ConnectionPendingTopicID != -1)
            {
                refdata[0, index] = CfgInfo.ConnectionPendingTopicID;
                refdata[1, index] = CfgInfo.ConnectionPending;
                ++index;
            }

            if (CfgInfo.ReadPendingTopicID != -1)
            {
                refdata[0, index] = CfgInfo.ReadPendingTopicID;
                refdata[1, index] = CfgInfo.ReadPending;
                ++index;
            }

            if (CfgInfo.ServerRunningTopicID != -1)
            {
                CfgInfo.ServerRunning = !CfgInfo.ServerDisable;
                refdata[0, index] = CfgInfo.ServerRunningTopicID;
                refdata[1, index] = CfgInfo.ServerRunning.ToString();
                ++index;
            }

            if (CfgInfo.SendReceiveCountTopicID != -1)
            {
                refdata[0, index] = CfgInfo.SendReceiveCountTopicID;
                refdata[1, index] =
                    CfgInfo.SendCount.ToString(CultureInfo.InvariantCulture) + "/" +
                    CfgInfo.ReceiveCount.ToString(CultureInfo.InvariantCulture);
                ++index;
            }

            if (CfgInfo.PollRateTopicID != -1)
            {
                refdata[0, index] = CfgInfo.PollRateTopicID;
                refdata[1, index] = CfgInfo.PollRate;
                ++index;
            }

            if (CfgInfo.SocketTimeoutTopicID != -1)
            {
                refdata[0, index] = CfgInfo.SocketTimeoutTopicID;
                refdata[1, index] = CfgInfo.SocketTimeout;
                ++index;
            }

            if (CfgInfo.ExcelUpdateRateTopicID != -1)
            {
                refdata[0, index] = CfgInfo.ExcelUpdateRateTopicID;
                refdata[1, index] = CfgInfo.ExcelUpdateRate;
                ++index;
            }

            if (CfgInfo.DebugLevelTopicID != -1)
            {
                refdata[0, index] = CfgInfo.DebugLevelTopicID;
                refdata[1, index] = CfgInfo.DebugLevel;
                ++index;
            }

            if (CfgInfo.MaxqueueTopicID != -1)
            {
                refdata[0, index] = CfgInfo.MaxqueueTopicID;
                refdata[1, index] = CfgInfo.Maxqueue;
                ++index;
            }

            if (CfgInfo.MaxreadTopicID != -1)
            {
                refdata[0, index] = CfgInfo.MaxreadTopicID;
                refdata[1, index] = CfgInfo.Maxread;
                ++index;
            }

            if (CfgInfo.MinconnectionsTopicID != -1)
            {
                refdata[0, index] = CfgInfo.MinconnectionsTopicID;
                refdata[1, index] = CfgInfo.Minconnections;
                ++index;
            }

            if (CfgInfo.SimulateTopicID != -1)
            {
                refdata[0, index] = CfgInfo.SimulateTopicID;
                refdata[1, index] = CfgInfo.Simulate;
                ++index;
            }

            if (CfgInfo.ServerDisableTopicID != -1)
            {
                refdata[0, index] = CfgInfo.ServerDisableTopicID;
                refdata[1, index] = CfgInfo.ServerDisable.ToString();
                ++index;
            }

            if (CfgInfo.DeleteCachedValuesTopicID != -1)
            {
                refdata[0, index] = CfgInfo.DeleteCachedValuesTopicID;
                refdata[1, index] = CfgInfo.DeleteCachedValues.ToString();
                ++index;
            }

            topicCount = index; //  CfgInfo.TopicCount;

            //if (counterHelper != null)
            //{
            //        counterHelper.RawValue(SingleInstance_PerformanceCounters.TopicCount, CfgInfo.TopicCount);
            //        counterHelper.RawValue(SingleInstance_PerformanceCounters.ConnectionCount, CfgInfo.ConnectionCount);
            //        counterHelper.RawValue(SingleInstance_PerformanceCounters.SendCount, CfgInfo.SendCount);
            //        counterHelper.RawValue(SingleInstance_PerformanceCounters.ReceiveCount, CfgInfo.ReceiveCount);
            //}

            return refdata;
        }

        #endregion // RefreshData

        #region Heartbeat

        /// <summary>
        ///   Called by Excel from time to time to check this RTD server is alive.
        /// </summary>
        /// <returns> 1 to indicate RTD server is alive. </returns>
        public int Heartbeat()
        {
            return 1;
        }

        #endregion // Heartbeat

        #endregion // IRtdServer Interface

        #region MBMasterHelpers 

        private void MBmaster_OnConnectComplete(object sender, SocketAsyncEventArgs e)
        {
            if (e.ConnectSocket == null) return;
            var endPoint = (IPEndPoint) e.RemoteEndPoint;
            string ip = endPoint.Address.ToString();
            CfgInfo.ConnectionCount++;
            CfgInfo.ConnectionPending--;
            if ((CfgInfo.DebugLevel & LogLevel.ConnectDisconnect) == LogLevel.ConnectDisconnect)
                Log.Info("Socket Connected ip:{0}, Handle:{1}, ConnectionCount:{2}, ConnectionCount:{3}", ip, ((Socket)(sender)).Handle,
                         CfgInfo.ConnectionPending, CfgInfo.ConnectionCount);
        }

        private void MBmaster_OnResponseData(ushort id, byte unit, byte function, byte[] values)
        {

            CfgInfo.ReadPending--;

            CfgInfo.ReceiveCount++;

            if (id == 0) return;

            SocketTopic removerequest;

            // Find matching requestID in _requests dictionary
            if (!_requests.TryRemove(id, out removerequest))
            {
                if ((CfgInfo.DebugLevel & LogLevel.Errors) == LogLevel.Errors)
                    Log.Info("Poll Response *Request Not found ID:{0}, function:{1}, values:{2}{3} ",
                             id, function, values[0], values[function == 1 ? 0 : 1]);
                return;
            }

            var now = Now();

            // Response received. Reduce QueueDepth for socket. Change state to Idle 
            _sockets[removerequest.SocketKey].QueueDepth--;
            _sockets[removerequest.SocketKey].Lastevent = EventStates.Idle;
            _sockets[removerequest.SocketKey].Eventbegin = now;
            TopicInfo ti = _topics[removerequest.TopicKey];

            ti.ReceiveCount++;

            // If polling rate is set to 0 we only poll each register once. 
            // Decrement reference so we will stop checking socket for timeout once all scheduled reads are complete
            if (CfgInfo.PollRate == 0 && _sockets[removerequest.SocketKey].References > 0)
                _sockets[removerequest.SocketKey].References--;

            // If we received a new value than the one we previously had, store result in CurrentValue. Conversion is based on Debug settings and datatype.
            if (ti.Rawdata != values)
            {
                StringBuilder sb = new StringBuilder();
                ti.Rawdata = values;
                if ((CfgInfo.DebugLevel & LogLevel.HexOutput) == LogLevel.HexOutput)
                    ti.Currentvalue = Conversion.Parseresult(ti.Addr, ti.Length, ti.Rawdata, ti.Modicon, "D");
                else if ((CfgInfo.DebugLevel & LogLevel.NotException) == LogLevel.NotException)
                    ti.Currentvalue = "";
                else if ((CfgInfo.DebugLevel & LogLevel.StateInfoData) == LogLevel.StateInfoData)
                    ti.Currentvalue = ti.Lastevent + (now - ti.Eventbegin).ToString() +
                                      "|" + Conversion.Parseresult(ti.Addr, ti.Length, ti.Rawdata, ti.Modicon);
                else if ((CfgInfo.DebugLevel & LogLevel.StateInfoStats) == LogLevel.StateInfoStats)
                    ti.Currentvalue = ti.Lastevent + (now - ti.Eventbegin).ToString() +
                                      "|" + ti.SendCount.ToString() + "/" +
                                      ti.ReceiveCount.ToString();
                else if ((CfgInfo.DebugLevel & LogLevel.StateInfoStatsData) == LogLevel.StateInfoStatsData)
                    ti.Currentvalue = ti.Lastevent + (now - ti.Eventbegin).ToString() +
                                      "|" + ti.SendCount.ToString() + "/" +
                                      ti.ReceiveCount.ToString() +
//                                       "(" + ti.SocketQueueDepth.ToString() + ")" +
                                      "[" + Conversion.Parseresult(ti.Addr, ti.Length, ti.Rawdata, ti.Modicon,"D") + "]";
                else ti.Currentvalue = Conversion.Parseresult(ti.Addr, ti.Length, ti.Rawdata, ti.Modicon);

                if ((CfgInfo.DebugLevel & LogLevel.ResultsToMSMQ) == LogLevel.ResultsToMSMQ && !CfgInfo.ServerDisable)
                {
                    sb.Append(ti.Addr + "," + ti.Currentvalue.ToString(CultureInfo.InvariantCulture));
                    if (_msmq != null)
                        Messaging.SendExisting(_msmq, sb.ToString());
                    else
                        Messaging.Send(QueueName, sb.ToString());
                }
                _updateavailable = true;
            }

            // Poll is complete. State is Idle.
            ti.Eventbegin = now;
            ti.Lastevent = EventStates.Idle;

            if ((CfgInfo.DebugLevel & LogLevel.ResponseReceived) == LogLevel.ResponseReceived)
                Log.Info("Poll Response Received ID:{0}, function:{1}, OnSocket:{2}, Values:{3}{4} ",
                    id, function, removerequest.SocketKey, values[0], values[ function == 1 ? 0 : 1 ]);
        }

        private void MBmaster_OnException(ushort id, byte unit, byte function, byte exception)
        {
            if (id == 0) return;

            SocketTopic removerequest;

            // TODO: Assume a response though invalid was a response to a read. May be wrong.
            CfgInfo.ReadPending--;

            // Find matching requestID in _requests dictionary
            if (!_requests.TryRemove(id, out removerequest))
            {
                if ((CfgInfo.DebugLevel & LogLevel.Errors) == LogLevel.Errors)
                    Log.Info("Poll Response *Request Not found ID:{0}, function:{1}, exception:{2} ",
                             id, function, exception);
                return;
            }

            var now = Now();
            _sockets[removerequest.SocketKey].Eventbegin = now;
            _sockets[removerequest.SocketKey].QueueDepth--;

            // If Connection lost
            if (exception == 254)
            {
                CfgInfo.ConnectionCount--; 
                _sockets[removerequest.SocketKey].Lastevent = EventStates.New;
                return;
            }

            CfgInfo.ReceiveCount++;
            // Exception Response received. Reduce QueueDepth for socket. Change state to Idle 
            _sockets[removerequest.SocketKey].Lastevent = EventStates.Idle;
            TopicInfo ti = _topics[removerequest.TopicKey];

            ti.ReceiveCount++;

            // If polling rate is set to 0 we only poll each register once. 
            // Decrement reference so we will stop checking socket for timeout once all scheduled reads are complete

            // 10-27-13 Removed check to reduce references.  If read request errored out we should not consider the request serviced.
            //if (CfgInfo.PollRate == 0 && _sockets[removerequest.SocketKey].References > 0)
            //    _sockets[removerequest.SocketKey].References--;

            // If we received a new value than the one we previously had, store result in CurrentValue. Conversion is based on Debug settings and datatype.
            if (ti.Rawdata != BitConverter.GetBytes(exception)) 
            {
                var sb = new StringBuilder();
                ti.Rawdata = BitConverter.GetBytes(exception);
                if ((CfgInfo.DebugLevel & LogLevel.HexOutput) == LogLevel.HexOutput)
                    ti.Currentvalue = "!";
                else if ((CfgInfo.DebugLevel & LogLevel.NotException) != LogLevel.NotException)
                    ti.Currentvalue = "!";
                else if ((CfgInfo.DebugLevel & LogLevel.None) == LogLevel.None)
                    ti.Currentvalue = "!";
                else if ((CfgInfo.DebugLevel & LogLevel.StateInfoData) == LogLevel.StateInfoData)
                    ti.Currentvalue = ti.Lastevent + (now - ti.Eventbegin).ToString() + "|!";
                else if ((CfgInfo.DebugLevel & LogLevel.StateInfoStats) == LogLevel.StateInfoStats)
                    ti.Currentvalue = ti.Lastevent + (now - ti.Eventbegin).ToString() + "|" + ti.SendCount.ToString() + "/" +ti.ReceiveCount.ToString();
                else if ((CfgInfo.DebugLevel & LogLevel.StateInfoStatsData) == LogLevel.StateInfoStatsData)
                    ti.Currentvalue = ti.Lastevent + (now - ti.Eventbegin).ToString() +
                                      "|" + ti.SendCount.ToString() + "/" +
                                      ti.ReceiveCount.ToString() +
                                      "[!]";
                else ti.Currentvalue = "!";

                if ((CfgInfo.DebugLevel & LogLevel.ResultsToMSMQ) == LogLevel.ResultsToMSMQ && !CfgInfo.ServerDisable && (CfgInfo.DebugLevel & LogLevel.NotException) != LogLevel.NotException)
                {
                    sb.Append(ti.Addr + ",!");
                    if (_msmq != null)
                        Messaging.SendExisting(_msmq, sb.ToString());
                    else
                        Messaging.Send(QueueName, sb.ToString());
                }
                _updateavailable = true;
            }

            // If we receive an exception, we will stop trying to retrieve data for this register.
            ti.Eventbegin = now;
            ti.Lastevent = EventStates.Exception;

            if ((CfgInfo.DebugLevel & LogLevel.ResponseReceived) == LogLevel.ResponseReceived)
                Log.Info("Poll Exception Received ID:{0}, function:{1}, OnSocket:{2}, Exception:{3} ",
                         id, function, removerequest.SocketKey, exception);

        }
        #endregion// MBMasterHelpers

        public void Dispose()
        {
            if (_modbusscantimer != null)
                _modbusscantimer.Dispose();
        }
    }
}
