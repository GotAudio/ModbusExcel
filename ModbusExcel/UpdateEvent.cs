namespace ModbusExcel
{
    /// <summary>
    /// This class is useful when testing RTD servers
    /// from clients other than Excel.
    /// 
    /// When Excel calls the RTD ServerStart() method,
    /// it passes a callback through which the RTD can
    /// notify Excel that fresh real-time data is ready
    /// -- a tap on the shoulder of Excel, so to speak.
    /// When Excel is ready, it will ask the RTD server
    /// for the data by calling the RTD's RefreshData()
    /// method.
    /// </summary>
    public partial class UpdateEvent : IRTDUpdateEvent
    {
        public int HeartbeatInterval { get; set; }

        public UpdateEvent()
        {
            // Setting HeartbeatInterval to minus
            // one for Excel means that it should
            // not call the RTD Heartbeat() method.
            // For clients other than Excel we
            // are free to interpret this value
            // how we wish.
            HeartbeatInterval = -1;
        }

        public void Disconnect()
        {
            // Do nothing.
        }

        public void UpdateNotify()
        {
        }
    }
}
