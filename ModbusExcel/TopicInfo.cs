namespace ModbusExcel
{
    public class TopicInfo
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
}