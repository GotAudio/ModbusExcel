using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ModbusExcel
{
    /// <summary>
    /// By defining the IRtdServer interface
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
    /// https://github.com/AJTSheppard/Andrew-Sheppard/blob/master/MyRTD/MyRTD/IRtdServer.cs
    /// </summary>
    /// 
    [ComImport,
     TypeLibType((short)0x1040),
     Guid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8")]
    public interface IRtdServer
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(10)]
        int ServerStart([In, MarshalAs(UnmanagedType.Interface)] IRTDUpdateEvent callback);

        [return: MarshalAs(UnmanagedType.Struct)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(11)]
        object ConnectData([In] int topicId, [In, MarshalAs(UnmanagedType.SafeArray,
            SafeArraySubType = VarEnum.VT_VARIANT)] ref object[] parameters, [In, Out] ref bool newValue);

        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(12)]
        object[,] RefreshData([In, Out] ref int topicCount);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(13)]
        void DisconnectData([In] int topicId);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(14)]
        int Heartbeat();

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(15)]
        void ServerTerminate();
    }
}
