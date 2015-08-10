using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ModbusExcel
{
    /// <summary>
    /// By defining the IRTDUpdateEvent interface
    /// here we do away with the need to reference
    /// the "Microsoft Excel Object Library", which
    /// is version specific. This makes deployment
    /// of the RTD server simpler as we do not have
    /// to target any one specific version of Excel
    /// in this assembly.
    ///
    /// Note: The GUID used here is that used by
    /// Microsoft to identify this interface. You
    /// must use this exact same GUID, otherwise
    /// when a client uses this RTD server you will
    /// get a run-time error.
    /// https://github.com/AJTSheppard/Andrew-Sheppard/blob/master/MyRTD/MyRTD/IRTDUpdateEvent.cs
    /// </summary>
    [ComImport,
     TypeLibType((short)0x1040),
     Guid("A43788C1-D91B-11D3-8F39-00C04F3651B8")]
    public interface IRTDUpdateEvent
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10), PreserveSig]
        void UpdateNotify();

        [DispId(11)]
        int HeartbeatInterval
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
            get;
            [param: In]
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
            set;
        }

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
        void Disconnect();
    }
}
