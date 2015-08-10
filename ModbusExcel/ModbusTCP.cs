using System;
using System.Net;
using System.Net.Sockets;

namespace ModbusExcel
{
    /// <summary>
    ///   Modbus TCP common driver class. This class implements a modbus TCP master driver.
    ///   It supports the following commands:
    /// 
    ///   Read coils
    ///   Read discrete inputs
    ///   Write single coil
    ///   Write multiple cooils
    ///   Read holding register
    ///   Read input register
    ///   Write single register
    ///   Write multiple register
    /// 
    ///   All commands can be sent in synchronous or asynchronous mode. If a value is accessed
    ///   in synchronous mode the program will stop and wait for slave to response. If the 
    ///   slave didn't answer within a specified time a timeout exception is called.
    ///   The class uses multi threading for both synchronous and asynchronous access. For
    ///   the communication two lines are created. This is necessary because the synchronous
    ///   thread has to wait for a previous command to finish.
    /// 
    ///   This contains modificaitons to the original code here; http://www.codeproject.com/Articles/16260/Modbus-TCP-class
    /// </summary>
    public class Master : IDisposable
    {
        #region declarations
        // ------------------------------------------------------------------------
        // Constants for access

        private const byte FctError = 0;
        private const byte FctReadCoil = 1;
        private const byte FctReadDiscreteInputs = 2;
        private const byte FctReadHoldingRegister = 3;
        private const byte FctReadInputRegister = 4;
        private const byte FctWriteSingleCoil = 5;
        private const byte FctWriteSingleRegister = 6;
        private const byte FctWriteMultipleCoils = 15;
        private const byte FctWriteMultipleRegister = 16;
        private const byte FctReadWriteMultipleRegister = 23;

        private const byte FctReadAsciiTextBuffer = 65; // 3 -- fctReadHoldingRegister
//        private const byte FctWriteAsciiTextBuffer = 66; // 6 -- fctWriteSingleRegister

        /// <summary>
        ///   Constant for exception illegal function.
        /// </summary>
        private const byte ExcIllegalFunction = 1;

        /// <summary>
        ///   Constant for exception illegal data address.
        /// </summary>
        public const byte ExcIllegalDataAdr = 2;

        /// <summary>
        ///   Constant for exception illegal data value.
        /// </summary>
        public const byte ExcIllegalDataVal = 3;

        /// <summary>
        ///   Constant for exception slave device failure.
        /// </summary>
        public const byte ExcSlaveDeviceFailure = 4;

        /// <summary>
        ///   Constant for exception acknowledge.
        /// </summary>
        public const byte ExcAck = 5;

        /// <summary>
        ///   Constant for exception slave is busy/booting up.
        /// </summary>
        public const byte ExcSlaveIsBusy = 6;

        /// <summary>
        ///   Constant for exception gate path unavailable.
        /// </summary>
        public const byte ExcGatePathUnavailable = 10;

        /// <summary>
        ///   Constant for exception not connected.
        /// </summary>
        public const byte ExcExceptionNotConnected = 253;

        /// <summary>
        ///   Constant for exception connection lost.
        /// </summary>
        public const byte ExcExceptionConnectionLost = 254;

        /// <summary>
        ///   Constant for exception response timeout.
        /// </summary>
        public const byte ExcExceptionTimeout = 255;

        /// <summary>
        ///   Constant for exception wrong offset.
        /// </summary>
        private const byte ExcExceptionOffset = 128;

        /// <summary>
        ///   Constant for exception send failt.
        /// </summary>
        private const byte ExcSendFailt = 100;

        // ------------------------------------------------------------------------
        // Private declarations

        // ------------------------------------------------------------------------
        // Private declarations
        private static ushort _timeout = 5000;

        /// <summary>
        ///   .NET Socket.
        /// </summary>
        private Socket _tcpAsyCl;

        private readonly byte[] _tcpAsyClBuffer = new byte[8096];

        private Socket _tcpSynCl;


        // ------------------------------------------------------------------------
        #region Delegates

        /// <summary>
        ///   Response data event. This event is called when new data arrives
        /// </summary>
        public delegate void ResponseData(ushort id, byte unit, byte function, byte[] data);

        /// <summary>
        ///   Response data event. This event is called when new data arrives
        /// </summary>
        public event ResponseData OnResponseData;

        /// <summary>
        ///   Exception data event. This event is called when the data is incorrect
        /// </summary>
        public delegate void ExceptionData(ushort id, byte unit, byte function, byte exception);

        /// <summary>
        ///   Exception data event. This event is called when the data is incorrect
        /// </summary>
        public event ExceptionData OnException;

        /// <summary>
        ///   Connect data event. This event is called when new connection is made
        /// </summary>
        public event ConnectComplete OnConnectComplete;

        /// <summary>
        ///   ConnectComplete event. This event is called when new connection is made
        /// </summary>
        //    public event EventHandler<IPEndPointEventArgs> SocketConnected; 
        public delegate void ConnectComplete(object sender, SocketAsyncEventArgs e);

        #endregion

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Shows if a connection is active.
        /// </summary>
        public bool Connected { get; private set; }

        /// <summary>
        ///   Start connection to slave.
        /// </summary>
        /// <param name="ip"> IP adress of modbus slave. </param>
        /// <param name="port"> Port number of modbus slave. Usually port 502 is used. </param>
        public bool ConnectAsync(string ip, ushort port)
        {
            IPAddress parseip;
            if (IPAddress.TryParse(ip, out parseip) == false)
            {
                IPHostEntry hst = Dns.GetHostEntry(ip);
                ip = hst.AddressList[0].ToString();
            }

            // Connect asynchronous client
            Connected = false;
            _tcpAsyCl = new Socket(IPAddress.Parse(ip).AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            var e = new SocketAsyncEventArgs
                {
                    UserToken = _tcpAsyCl,
                    RemoteEndPoint = new IPEndPoint(IPAddress.Parse(ip), port)
                };
            e.Completed += OnConnect;

            // Returns true if the I/O operation is pending. (normal async behaviour)
            if (_tcpAsyCl.ConnectAsync(e))
            {
                SocketError errorCode = e.SocketError;
                return (errorCode == SocketError.Success);
            }
            // Otherwise it completed synchronously.  In 2000 simultaneous connections to local server this never happened once.

            if (OnConnectComplete != null) OnConnectComplete(_tcpAsyCl, e);
            Connected = true;
            return Connected;
        }

        public bool ConnectSync(string ip, ushort port)
        {
            IPAddress parseip;
            if (IPAddress.TryParse(ip, out parseip) == false)
            {
                IPHostEntry hst = Dns.GetHostEntry(ip);
                ip = hst.AddressList[0].ToString();
            }

            // Connect sync client
            _tcpAsyCl = new Socket(IPAddress.Parse(ip).AddressFamily, SocketType.Stream, ProtocolType.Tcp);
            _tcpAsyCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.SendTimeout, _timeout);
            _tcpAsyCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.ReceiveTimeout, _timeout);
            _tcpAsyCl.SetSocketOption(SocketOptionLevel.Socket, SocketOptionName.NoDelay, 1);
            try
            {
                _tcpAsyCl.Connect(new IPEndPoint(IPAddress.Parse(ip), port));
            }
            catch (Exception)
            {
                return false;
            }
            Connected = true;
            return Connected;
        }

        private void OnConnect(object sender, SocketAsyncEventArgs e)
        {
            if (OnConnectComplete != null)
            {
                OnConnectComplete(sender, e);
            }
            Connected = true;
        }

        #endregion declarations
        // ------------------------------------------------------------------------
        /// <summary>
        ///   Stop connection to slave.
        /// </summary>
        public void Disconnect()
        {
            Dispose();
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Destroy master instance.
        /// </summary>
        ~Master()
        {
            Dispose();
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Destroy master instance
        /// </summary>
        public void Dispose()
        {
            if (_tcpAsyCl != null)
            {
                if (_tcpAsyCl.Connected)
                {
                    try
                    {
                        _tcpAsyCl.Shutdown(SocketShutdown.Both);
                    }
// ReSharper disable EmptyGeneralCatchClause
                    catch
// ReSharper restore EmptyGeneralCatchClause
                    {
                    }
                    _tcpAsyCl.Close();
                }
 //               _tcpAsyCl.Dispose();
                _tcpAsyCl = null;
            }
            if (_tcpSynCl != null)
            {
                if (_tcpSynCl.Connected)
                {
                    try
                    {
                        _tcpSynCl.Shutdown(SocketShutdown.Both);
                    }
// ReSharper disable EmptyGeneralCatchClause
                    catch
// ReSharper restore EmptyGeneralCatchClause
                    {
                    }
                    _tcpSynCl.Close();
                }
//                _tcpSynCl.Dispose();
                _tcpSynCl = null;
            }
        }

        private void CallException(ushort id, byte unit, byte function, byte exception)
        {
            // Return if no connections exist on which to report
            if ((_tcpAsyCl == null) && (_tcpSynCl == null)) return;
            if (exception == ExcExceptionConnectionLost)
            {
                _tcpSynCl = null;
                _tcpAsyCl = null;
            }

            // 6-22-13 begin Stephen Added unit
            if (OnException != null) OnException(id, unit, function, exception);
        }

        private static UInt16 SwapUInt16(UInt16 inValue)
        {
            return (UInt16)(((inValue & 0xff00) >> 8) |
                     ((inValue & 0x00ff) << 8));
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read coils from slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="modicon"></param>
        public void ReadCoils(ushort id, byte unit, ushort startAddress, ushort numInputs, string modicon)
        {
            WriteAsyncData(
                modicon == "s"
                    ? CreateReadHeaderRTU(unit, startAddress, numInputs, FctReadCoil)
                    : CreateReadHeader(id, unit, startAddress, numInputs, FctReadCoil), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read coils from slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="modicon"></param>
        /// <param name="values"> Contains the result of function. </param>
        public void ReadCoils(ushort id, byte unit, ushort startAddress, ushort numInputs, string modicon, ref byte[] values)
        {
            var data = new byte[1];
            data[0] = Connected ? (byte)1 : (byte)0;

            //            Array.Copy(_tcpAsyClBuffer, 10, data, 0, 2);
            values = data;
            WriteSyncData(modicon == "s" 
            ? CreateReadHeaderRTU(unit, startAddress, numInputs, FctReadCoil) 
            : CreateReadHeader(id, unit, startAddress, numInputs, FctReadCoil), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read discrete inputs from slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        public void ReadDiscreteInputs(ushort id, byte unit, ushort startAddress, ushort numInputs)
        {
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, FctReadDiscreteInputs), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read discrete inputs from slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="values"> Contains the result of function. </param>
        public void ReadDiscreteInputs(ushort id, byte unit, ushort startAddress, ushort numInputs, ref byte[] values)
        {
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, FctReadDiscreteInputs), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read holding registers from slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="modicon">s=Serial </param>
        public void ReadHoldingRegister(ushort id, byte unit, ushort startAddress, ushort numInputs, string modicon)
        {
            WriteAsyncData(
                modicon == "s"
                    ? CreateReadHeaderRTU(unit, startAddress, numInputs, FctReadHoldingRegister)
                    : CreateReadHeader(id, unit, startAddress, numInputs, FctReadHoldingRegister), id);
        }

        /// <summary>
        ///   Read holding registers from slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id">
        ///     Unique id that marks the transaction. In asynchonous mode this id is given to the callback function.
        /// </param>
        /// <param name="unit">
        ///     Slave Unit ID to request from. 1-255 for TCP.
        /// </param>
        /// <param name="startAddress">
        ///     Address from where the data read begins.
        /// </param>
        /// <summary>
        ///   Read holding registers from slave synchronous.
        /// </summary>
        /// <param name="numInputs">
        ///     Length of data.
        /// </param>
        /// <param name="values"> Contains the result of function. </param>
        public void ReadHoldingRegister(ushort id, byte unit, ushort startAddress, ushort numInputs, string modicon, ref byte[] values)
        {
            values = WriteSyncData(modicon == "s" 
                    ? CreateReadHeaderRTU(unit, startAddress, numInputs, FctReadHoldingRegister)
                    : CreateReadHeader(id, unit, startAddress, numInputs, FctReadHoldingRegister), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read text buffer from slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        public void ReadAsciiTextBuffer(ushort id, byte unit, ushort startAddress, ushort numInputs)
        {
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, FctReadAsciiTextBuffer), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read text buffer from slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="values"> Contains the result of function. </param>
        public void ReadAsciiTextBuffer(ushort id, byte unit, ushort startAddress, ushort numInputs, ref byte[] values)
        {
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, FctReadAsciiTextBuffer), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read input registers from slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        public void ReadInputRegister(ushort id, byte unit, ushort startAddress, ushort numInputs)
        {
            WriteAsyncData(CreateReadHeader(id, unit, startAddress, numInputs, FctReadInputRegister), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read input registers from slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="values"> Contains the result of function. </param>
        public void ReadInputRegister(ushort id, byte unit, ushort startAddress, ushort numInputs, ref byte[] values)
        {
            values = WriteSyncData(CreateReadHeader(id, unit, startAddress, numInputs, FctReadInputRegister), id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write single coil in slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="onOff"> Specifys if the coil should be switched on or off. </param>
        public void WriteSingleCoils(ushort id, byte unit, ushort startAddress, bool onOff)
        {
            byte[] data;
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, FctWriteSingleCoil);
            if (onOff == true) data[10] = 255;
            else data[10] = 0;
            WriteAsyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write single coil in slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="onOff"> Specifys if the coil should be switched on or off. </param>
        /// <param name="result"> Contains the result of the synchronous write. </param>
        public void WriteSingleCoils(ushort id, byte unit, ushort startAddress, bool onOff, ref byte[] result)
        {
            byte[] data;
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, FctWriteSingleCoil);
            if (onOff == true) data[10] = 255;
            else data[10] = 0;
            result = WriteSyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write multiple coils in slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numBits"> Specifys number of bits. </param>
        /// <param name="values"> Contains the bit information in byte format. </param>
        public void WriteMultipleCoils(ushort id, byte unit, ushort startAddress, ushort numBits, byte[] values)
        {
            byte numBytes = Convert.ToByte(values.Length);
            byte[] data = CreateWriteHeader(id, unit, startAddress, numBits, (byte)(numBytes + 2), FctWriteMultipleCoils);
            Array.Copy(values, 0, data, 13, numBytes);
            WriteAsyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write multiple coils in slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address from where the data read begins. </param>
        /// <param name="numBits"> Specifys number of bits. </param>
        /// <param name="values"> Contains the bit information in byte format. </param>
        /// <param name="result"> Contains the result of the synchronous write. </param>
        public void WriteMultipleCoils(ushort id, byte unit, ushort startAddress, ushort numBits, byte[] values,
                                       ref byte[] result)
        {
            byte numBytes = Convert.ToByte(values.Length);
            byte[] data;
            data = CreateWriteHeader(id, unit, startAddress, numBits, (byte)(numBytes + 2), FctWriteMultipleCoils);
            Array.Copy(values, 0, data, 13, numBytes);
            result = WriteSyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write single register in slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address to where the data is written. </param>
        /// <param name="values"> Contains the register information. </param>
        public void WriteSingleRegister(ushort id, byte unit, ushort startAddress, byte[] values)
        {
            byte[] data;
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, FctWriteSingleRegister);
            data[10] = values[0];
            data[11] = values[1];
            WriteAsyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write single register in slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address to where the data is written. </param>
        /// <param name="values"> Contains the register information. </param>
        /// <param name="result"> Contains the result of the synchronous write. </param>
        public void WriteSingleRegister(ushort id, byte unit, ushort startAddress, byte[] values, ref byte[] result)
        {
            byte[] data;
            data = CreateWriteHeader(id, unit, startAddress, 1, 1, FctWriteSingleRegister);
            data[10] = values[0];
            data[11] = values[1];
            result = WriteSyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write multiple registers in slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address to where the data is written. </param>
        /// <param name="values"> Contains the register information. </param>
        public void WriteMultipleRegister(ushort id, byte unit, ushort startAddress, byte[] values)
        {
            ushort numBytes = Convert.ToUInt16(values.Length);
            if (numBytes % 2 > 0) numBytes++;
            byte[] data;

            data = CreateWriteHeader(id, unit, startAddress, Convert.ToUInt16(numBytes / 2),
                                     Convert.ToUInt16(numBytes + 2), FctWriteMultipleRegister);
            Array.Copy(values, 0, data, 13, values.Length);
            WriteAsyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Write multiple registers in slave synchronous.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startAddress"> Address to where the data is written. </param>
        /// <param name="values"> Contains the register information. </param>
        /// <param name="result"> Contains the result of the synchronous write. </param>
        public void WriteMultipleRegister(ushort id, byte unit, ushort startAddress, byte[] values, ref byte[] result)
        {
            ushort numBytes = Convert.ToUInt16(values.Length);
            if (numBytes % 2 > 0) numBytes++;
            byte[] data;

            data = CreateWriteHeader(id, unit, startAddress, Convert.ToUInt16(numBytes / 2),
                                     Convert.ToUInt16(numBytes + 2), FctWriteMultipleRegister);
            Array.Copy(values, 0, data, 13, values.Length);
            result = WriteSyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read/Write multiple registers in slave asynchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startReadAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="startWriteAddress"> Address to where the data is written. </param>
        /// <param name="values"> Contains the register information. </param>
        public void ReadWriteMultipleRegister(ushort id, byte unit, ushort startReadAddress, ushort numInputs,
                                              ushort startWriteAddress, byte[] values)
        {
            ushort numBytes = Convert.ToUInt16(values.Length);
            if (numBytes % 2 > 0) numBytes++;
            byte[] data;

            data = CreateReadWriteHeader(id, unit, startReadAddress, numInputs, startWriteAddress,
                                         Convert.ToUInt16(numBytes / 2));
            Array.Copy(values, 0, data, 17, values.Length);
            WriteAsyncData(data, id);
        }

        // ------------------------------------------------------------------------
        /// <summary>
        ///   Read/Write multiple registers in slave synchronous. The result is given in the response function.
        /// </summary>
        /// <param name="id"> Unique id that marks the transaction. In asynchonous mode this id is given to the callback function. </param>
        /// <param name="unit"> Slave Unit ID to request from. 1-255 for TCP. </param>
        /// <param name="startReadAddress"> Address from where the data read begins. </param>
        /// <param name="numInputs"> Length of data. </param>
        /// <param name="startWriteAddress"> Address to where the data is written. </param>
        /// <param name="values"> Contains the register information. </param>
        /// <param name="result"> Contains the result of the synchronous command. </param>
        public void ReadWriteMultipleRegister(ushort id, byte unit, ushort startReadAddress, ushort numInputs,
                                              ushort startWriteAddress, byte[] values, ref byte[] result)
        {
            ushort numBytes = Convert.ToUInt16(values.Length);
            if (numBytes % 2 > 0) numBytes++;
            byte[] data;

            data = CreateReadWriteHeader(id, unit, startReadAddress, numInputs, startWriteAddress,
                                         Convert.ToUInt16(numBytes / 2));
            Array.Copy(values, 0, data, 17, values.Length);
            result = WriteSyncData(data, id);
        }

        // ------------------------------------------------------------------------
        // Create modbus header for read action
        private byte[] CreateReadHeaderRTU(byte unit, ushort startAddress, ushort points, byte function)
        {
            //0001 	transaction
            //0000 	protocol
            //0006 	length
            //-- start at 7
            //01 	unit
            //03 	function
            //1B BE 	addr
            //00 01	length

            var data = new byte[6];

            //byte[] _id = BitConverter.GetBytes((short)id);
            //data[0] = _id[0];				// Slave id high byte
            //data[1] = _id[1];				// Slave id low byte
            //data[5] = 6;					// Message size
            data[0] = unit; // KAS Was 0 Slave address
            data[1] = function; // Function code
            byte[] adr = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)startAddress));
            data[2] = adr[0]; // Start address
            data[3] = adr[1]; // Start address
            byte[] length = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)points));
            data[4] = length[0]; // Number of data to read
            data[5] = length[1]; // Number of data to read

            byte[] crc = ModbusUtil.CalculateCrc(data);

            var ret = new byte[data.Length + crc.Length];
            Buffer.BlockCopy(data, 0, ret, 0, data.Length);
            Buffer.BlockCopy(crc, 0, ret, data.Length, crc.Length);
            return ret;
        }


        // ------------------------------------------------------------------------
        // Create modbus header for read action
        //
        private byte[] CreateReadHeader(ushort transactionId, byte unit, ushort startAddress, ushort points, byte function)
        {
            var data = new byte[12];

            byte[] id = BitConverter.GetBytes((short)transactionId);
            data[0] = id[1];			    // Slave id high byte
            data[1] = id[0];				// Slave id low byte
            data[5] = 6; // Message size
            data[6] = unit;					// Slave address
            data[7] = function; // Function code
            byte[] adr = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)startAddress));
            data[8] = adr[0]; // Start address
            data[9] = adr[1]; // Start address
            byte[] length = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)points));
            data[10] = length[0]; // Number of data to read
            data[11] = length[1]; // Number of data to read
            return data;
        }

        // ------------------------------------------------------------------------
        // Create modbus header for write action
        private byte[] CreateWriteHeader(ushort transactionId, byte unit, ushort startAddress, ushort numData, ushort numBytes,
                                         byte function)
        {
            var data = new byte[numBytes + 11];

            byte[] id = BitConverter.GetBytes((short)transactionId);
            data[0] = id[1];				// Slave id high byte
            data[1] = id[0];				// Slave id low byte
            byte[] size = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)(5 + numBytes)));
            data[4] = size[0]; // Complete message size in bytes
            data[5] = size[1]; // Complete message size in bytes
            data[6] = unit;					// Slave address
            data[7] = function; // Function code
            byte[] adr = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)startAddress));
            data[8] = adr[0]; // Start address
            data[9] = adr[1]; // Start address
            if (function >= FctWriteMultipleCoils)
            {
                byte[] cnt = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)numData));
                data[10] = cnt[0]; // Number of bytes
                data[11] = cnt[1]; // Number of bytes
                data[12] = (byte)(numBytes - 2);
            }
            return data;
        }

        // ------------------------------------------------------------------------
        // Create modbus header for write action
        private byte[] CreateReadWriteHeader(ushort transactionId, byte unit, ushort startReadAddress, ushort numRead,
                                             ushort startWriteAddress, ushort numWrite)
        {
            var data = new byte[numWrite * 2 + 17];

            byte[] id = BitConverter.GetBytes((short)transactionId);
            data[0] = id[1];						// Slave id high byte
            data[1] = id[0];						// Slave id low byte
            byte[] size = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)(11 + numWrite * 2)));
            data[4] = size[0]; // Complete message size in bytes
            data[5] = size[1]; // Complete message size in bytes
            data[6] = unit;							// Slave address
            data[7] = FctReadWriteMultipleRegister; // Function code
            byte[] adrRead = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)startReadAddress));
            data[8] = adrRead[0]; // Start read address
            data[9] = adrRead[1]; // Start read address
            byte[] cntRead = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)numRead));
            data[10] = cntRead[0]; // Number of bytes to read
            data[11] = cntRead[1]; // Number of bytes to read
            byte[] adrWrite = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)startWriteAddress));
            data[12] = adrWrite[0]; // Start write address
            data[13] = adrWrite[1]; // Start write address
            byte[] cntWrite = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)numWrite));
            data[14] = cntWrite[0]; // Number of bytes to write
            data[15] = cntWrite[1]; // Number of bytes to write
            data[16] = (byte)(numWrite * 2);

            return data;
        }

        // ------------------------------------------------------------------------
        // Write asynchronous data
        private void WriteAsyncData(byte[] writeData, ushort id)
        {
            if ((_tcpAsyCl != null) && (_tcpAsyCl.Connected))
            {
                try
                {
                    _tcpAsyCl.BeginSend(writeData, 0, writeData.Length, SocketFlags.None, OnSend, null);
                    _tcpAsyCl.BeginReceive(_tcpAsyClBuffer, 0, _tcpAsyClBuffer.Length, SocketFlags.None, OnReceive,
                                          _tcpAsyCl);
                }
                catch (SystemException)
                {
                    CallException(id, writeData[6], writeData[7], ExcExceptionConnectionLost);
                }
            }
            else CallException(id, writeData[6], writeData[7], ExcExceptionConnectionLost);
        }

        // ------------------------------------------------------------------------
        // Write asynchronous data acknowledge
        private void OnSend(IAsyncResult result)
        {
            if (result.IsCompleted == false) CallException(0xFFFF, 0xFF, 0xFF, ExcSendFailt);
        }

        // ------------------------------------------------------------------------
        // Write asynchronous data response
        private void OnReceive(IAsyncResult result)
        {
            int nBytesReceived = 0;
            try
            {
                if (_tcpAsyCl != null) nBytesReceived = _tcpAsyCl.EndReceive(result);
                if (nBytesReceived == 0)
                {
                    //              no answer
                    //              CallException(0, 0, excExceptionTimeout);
                    return;
                }
            }
            catch (ObjectDisposedException)
            {
                // Service must be shuting down. Eat the exception
                return;
            }
            catch (SocketException)
            {
                //                if (ex.SocketErrorCode == SocketError.ConnectionAborted) {do what?}
                CallException(0xFF, 0xFF, 0xFF, ExcExceptionConnectionLost);
                return;
            }

            if (result.IsCompleted == false)
            {
                CallException(0xFF, 0xFF, 0xFF, ExcExceptionConnectionLost);
                return;
            }

            // Loop thorugh received transactions and split them into Modbus responses, calling client OnResponseData for each
            int resultptr = 0;

            while (resultptr < nBytesReceived)
            {
                ushort id = SwapUInt16(BitConverter.ToUInt16(_tcpAsyClBuffer, resultptr));
                byte unit = _tcpAsyClBuffer[resultptr + 6];

                byte function = _tcpAsyClBuffer[resultptr + 7];

                // Response data is slave exception or unrecognized function is data corruption somehow
                if (function == FctError || function > ExcExceptionOffset)
                {
                    function -= (function == FctError) ? FctError : ExcExceptionOffset;
                    CallException(id, unit, function, _tcpAsyClBuffer[resultptr + 8]);
                    resultptr = resultptr + 9;
                    _tcpAsyCl.BeginReceive(_tcpAsyClBuffer, 0, _tcpAsyClBuffer.Length, SocketFlags.None, OnReceive,_tcpAsyCl);
                }
                else
                {
                    byte[] data;
                    switch (function)
                    {
        // Text Read data - always 128 bytes + 12 byte header.  T
        //  TODO: This will only be first data block. Must trigger reads for subsequent blocks by incrementing length in mbapheader[9] until hex 1A is seen.
        // Somehow this buffer needs to be saved and accumulated and returned with the transaction id given in the original request. Subsequent requests must
        // be given pseudo transaction ids that will not collide with actual requests. Or the request needs to be done synchronously. 
                        case FctReadAsciiTextBuffer:
                            data = new byte[128];
                            Array.Copy(_tcpAsyClBuffer, resultptr + 12, data, 0, 128);
                            resultptr = resultptr + 140;
                            break;
                            // Read Holding Register - Always 8 bytes at a a time - Verify that is true ...
                        case FctReadHoldingRegister:
                            data = new byte[_tcpAsyClBuffer[resultptr + 8]];
                            Array.Copy(_tcpAsyClBuffer, resultptr + 9, data, 0, _tcpAsyClBuffer[resultptr + 8]);
                            resultptr = resultptr + 9 + _tcpAsyClBuffer[resultptr + 8];
                            break;
                            // Read Holding Register - Always 1 byte at a a time. This may  not be right if more than 1 register read
                        case FctReadCoil:
                            data = new byte[1];
                            Array.Copy(_tcpAsyClBuffer, resultptr + 9, data, 0, 1);
                            resultptr = resultptr + 10;
                            break;
                        default:
                            // Any other write operation returns address written to
                            data = new byte[2];
                            Array.Copy(_tcpAsyClBuffer, resultptr + 10, data, 0, 2);
                            resultptr = resultptr + 10;
                            break;
                    }
                    //Try to read socket again in case there is more data to receive. According to examples on MSDN
                    _tcpAsyCl.BeginReceive(_tcpAsyClBuffer, 0, _tcpAsyClBuffer.Length, SocketFlags.None, OnReceive,_tcpAsyCl);
                    // Call delegate of client if assigned
                    if (OnResponseData != null) OnResponseData(id, unit, function, data);
                }
            }
        }

        // Write data and and wait for response
        private byte[] WriteSyncData(byte[] writeData, ushort id)
        {
            if ((_tcpAsyCl != null) && (_tcpAsyCl.Connected))
            {
                try
                {
                    _tcpAsyCl.Send(writeData, 0, writeData.Length, SocketFlags.None);

                    int result;

                    // Poll the socket for reception with a 80000 ms timeout.
                    if (_tcpAsyCl.Poll(60000, SelectMode.SelectRead))
                    {
                         result = _tcpAsyCl.Receive(_tcpAsyClBuffer, 0, _tcpAsyClBuffer.Length, SocketFlags.None);
                    }
                    else
                    {
                        result = 0; // Timed out
                    }

                    byte unit = _tcpAsyClBuffer[6];
                    byte function = _tcpAsyClBuffer[7];
                    byte[] data;

                    if (function == 0x7A)
                        CallException(id, unit, writeData[7], ExcIllegalFunction);

                    if (result == 0)
                    {
                        CallException(id, unit, writeData[7], ExcExceptionConnectionLost);
                        return null;
                    }

                    // ------------------------------------------------------------
                    // Response data is slave exception
                    if (function > ExcExceptionOffset)
                    {
                        function -= ExcExceptionOffset;
                        CallException(id, unit, function, _tcpAsyClBuffer[8]);
                        return null;
                    }
                    // ------------------------------------------------------------
                    
                    // Write response data
                    if ((function >= FctWriteSingleCoil) && (function != FctReadWriteMultipleRegister) &&
                        (function != FctReadAsciiTextBuffer))
                    {
                        data = new byte[2];
                        Array.Copy(_tcpAsyClBuffer, 10, data, 0, 2);
                    }
                        // ------------------------------------------------------------
                        // Read response data
                    else
                    {
                        data = new byte[_tcpAsyClBuffer[8]];
                        Array.Copy(_tcpAsyClBuffer, 9, data, 0, _tcpAsyClBuffer[8]);
                    }
                    return data;
                }
                catch (SystemException)
                {
                    CallException(id, writeData[6], writeData[7], ExcExceptionConnectionLost);
                }
            }
            else CallException(id, writeData[6], writeData[7], ExcExceptionConnectionLost);
            return null;
        }
    }
}