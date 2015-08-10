using System;
using System.Runtime.InteropServices;

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
    [Guid("1960436F-33A3-4146-BE75-EE90B9EDCD01"),
     ProgId("ModbusExcel.UpdateEvent"),
     ComVisible(true)]
    public partial class UpdateEvent : IRTDUpdateEvent
    {
        private const String eventName = "ModbusExcel.UpdateEvent";
    }
}
