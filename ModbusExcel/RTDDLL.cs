using System;
using System.Runtime.InteropServices;

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
        private const String rtdName = "ModbusExcel.RTD";
    }
}