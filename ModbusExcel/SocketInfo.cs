namespace ModbusExcel
{
    public class SocketInfo
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
}