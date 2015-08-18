namespace ModbusExcel
{
    /// <summary>
    ///   Debug Log Level Bitmasks
    /// </summary>
    static class LogLevel
    {
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
}