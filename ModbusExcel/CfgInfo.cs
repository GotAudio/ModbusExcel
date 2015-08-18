namespace ModbusExcel
{
    /// <summary>
    ///   Stores user configurable values for poller.  Even though the client (Excel) may initalize these values, they may
    ///   not get assigned before other data collection topics have been registered and polling has potentially started, so we assign 
    ///   default values that will be used prior to receiving new values for them.
    /// </summary>
    static class CfgInfo
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
        public static ConnectionType ConnectionType;
    }
}