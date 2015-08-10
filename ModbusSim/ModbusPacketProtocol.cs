using Nito.Async;
using Nito.Async.Sockets;

namespace ModbusSim
{
    using System;

    /// <summary>
    /// Maintains the necessary buffers for applying a packet protocol over a stream-based socket.
    /// </summary>
    /// <remarks>
    /// <para>This class uses a 4-byte signed integer length prefix, which allows for packet sizes up to 2 GB with single-bit error detection.</para>
    /// <para>Keepalive packets are supported as packets with a length prefix of 0; <see cref="PacketArrived"/> is never called when keepalive packets are returned.</para>
    /// <para>Once <see cref="Start"/> is called, this class continuously reads from the underlying socket, calling <see cref="PacketArrived"/> when packets are received. To stop reading, close the underlying socket.</para>
    /// <para>Reading will also automatically stop when a read error or graceful close is detected.</para>
    /// </remarks>
    public class ModbusPacketProtocols
    {
        ///// <summary>
        ///// The buffer for the length prefix; this is always 4 bytes long.
        ///// </summary>
        //private byte[] lengthBuffer;

        /// <summary>
        /// The buffer for the data; this is null if we are receiving the length prefix buffer.
        /// </summary>
        // private byte[] dataBuffer;

        /// <summary>
        /// KAS Replaced message framed dataBuffer with tcpAsyClBuffer
        /// </summary>
        private readonly byte[] _tcpAsyClBuffer = new byte[8096];

        /// <summary>
        /// The number of bytes already read into the buffer (the length buffer if <see cref="dataBuffer"/> is null, otherwise the data buffer).
        /// </summary>
        private int _bytesReceived;

        /// <summary>
        /// Initializes a new instance of the <see cref="ModbusPacketProtocols"/> class bound to a given socket connection.
        /// </summary>
        /// <param name="socket">The socket used for communication.</param>
        public ModbusPacketProtocols(IAsyncTcpConnection socket)
        {
            this.Socket = socket;
//            this.lengthBuffer = new byte[sizeof(int)];
        }

        /// <summary>
        /// Indicates the completion of a packet read from the socket.
        /// </summary>
        /// <remarks>
        /// <para>This may be called with a null packet, indicating that the other end graciously closed the connection.</para>
        /// </remarks>
        public event Action<AsyncResultEventArgs<byte[]>> PacketArrived;

        /// <summary>
        /// Gets the socket used for communication.
        /// </summary>
        public IAsyncTcpConnection Socket { get; private set; }

        /// <overloads>
        /// <summary>Sends a packet to a socket.</summary>
        /// <remarks>
        /// <para>Generates a length prefix for the packet and writes the length prefix and packet to the socket.</para>
        /// </remarks>
        /// </overloads>
        /// <summary>Sends a packet to a socket.</summary>
        /// <remarks>
        /// <para>Generates a length prefix for the packet and writes the length prefix and packet to the socket.</para>
        /// </remarks>
        /// <param name="socket">The socket used for communication.</param>
        /// <param name="packet">The packet to send.</param>
        /// <param name="state">The user-defined state that is passed to WriteCompleted. May be null.</param>
        public static void WritePacketAsync(IAsyncTcpConnection socket, byte[] packet, object state)
        {
            //// Get the length prefix for the message
            //byte[] lengthPrefix = BitConverter.GetBytes(packet.Length);

            //// We use the special CallbackOnErrorsOnly object to tell the socket we don't want
            ////  WriteCompleted to be invoked (it would confuse socket users if they see WriteCompleted
            ////  events for writes they never started).
            //socket.WriteAsync(lengthPrefix, new CallbackOnErrorsOnly());

            // Send the actual message, this time enabling the normal callback.
            socket.WriteAsync(packet, state);
        }

        /// <inheritdoc cref="WritePacketAsync(IAsyncTcpConnection, byte[], object)" />
        /// <param name="socket">The socket used for communication.</param>
        /// <param name="packet">The packet to send.</param>
        public static void WritePacketAsync(IAsyncTcpConnection socket, byte[] packet)
        {
            WritePacketAsync(socket, packet, null);
        }

        /// <summary>
        /// Sends a keepalive (0-length) packet to the socket.
        /// </summary>
        /// <param name="socket">The socket used for communication.</param>
        public static void WriteKeepaliveAsync(IAsyncTcpConnection socket)
        {
            // We use CallbackOnErrorsOnly to indicate that the WriteCompleted callback should only be
            //  called if there was an error.
            socket.WriteAsync(BitConverter.GetBytes(0), new CallbackOnErrorsOnly());
        }

        /// <summary>
        /// Begins reading from the socket.
        /// </summary>
        public void Start()
        {
            Socket.ReadCompleted += SocketReadCompleted;
            ContinueReading();
        }

        /// <summary>
        /// Requests a read directly into the correct buffer.
        /// </summary>
        private void ContinueReading()
        {
            Socket.ReadAsync(_tcpAsyClBuffer, 0, _tcpAsyClBuffer.Length);
        }

        internal static UInt16 SwapUInt16(UInt16 inValue)
        {
            return (UInt16)(((inValue & 0xff00) >> 8) | ((inValue & 0x00ff) << 8));
        }

        /// <summary>
        /// Called when a socket read completes. Parses the received data and calls <see cref="PacketArrived"/> if necessary.
        /// </summary>
        /// <param name="e">Argument object containing the number of bytes read.</param>
        /// <exception cref="System.IO.InvalidDataException">If the data received is not a packet.</exception>
        private void SocketReadCompleted(AsyncResultEventArgs<int> e)
        {
            // Pass along read errors verbatim
            if (e.Error != null)
            {
                if (PacketArrived != null)
                {
                    PacketArrived(new AsyncResultEventArgs<byte[]>(e.Error));
                }

                return;
            }

            // Get the number of bytes read into the buffer
            _bytesReceived += e.Result;

            // If we get a zero-length read, then that indicates the remote side graciously closed the connection
            if (e.Result == 0)
            {
                if (PacketArrived != null)
                {
                    PacketArrived(new AsyncResultEventArgs<byte[]>((byte[])null));
                }
                return;
            }

            // Loop thorugh received transactions and split them into Modbus responses, calling client OnResponseData for each
            int resultptr = 0;
            while (resultptr < e.Result)
            {
                byte[] data;
                ushort id = SwapUInt16(BitConverter.ToUInt16(_tcpAsyClBuffer, resultptr));
                byte unit = _tcpAsyClBuffer[resultptr + 6];
                byte function = _tcpAsyClBuffer[resultptr + 7];

                switch (function)
                {
                    // Text Read data - always 128 bytes + 12 byte header
                    case 65:
                        data = new byte[128];
                        Array.Copy(_tcpAsyClBuffer, resultptr + 12, data, 0, 128);
                        resultptr = resultptr + 140;
                        break;
                    case 1:
                        data = new byte[_tcpAsyClBuffer[resultptr + 5] + 6];
                        Array.Copy(_tcpAsyClBuffer, resultptr, data, 0, _tcpAsyClBuffer[resultptr + 5] + 6);
                        resultptr = resultptr + _tcpAsyClBuffer[resultptr + 5] + 6;
                        break;

                        //data = new byte[1];
                        //Array.Copy(_tcpAsyClBuffer, resultptr + 9, data, 0, 1);
                        //resultptr = resultptr + 10;
                        //break;
                    case 3:
                        //00 01 00 00 00 06 01 03 00 20 00 01
                        //00 01 00 00 00 06 01 03 1B B1 00 01
                        data = new byte[_tcpAsyClBuffer[resultptr + 5] + 6];
                        Array.Copy(_tcpAsyClBuffer, resultptr, data, 0, _tcpAsyClBuffer[resultptr + 5] + 6);
                        resultptr = resultptr + _tcpAsyClBuffer[resultptr + 5] + 6;
                        break;


                        //data = new byte[tcpAsyClBuffer[resultptr + 8]];
                        //Array.Copy(tcpAsyClBuffer, resultptr + 9, data, 0, tcpAsyClBuffer[resultptr + 8]);
                        //resultptr = resultptr + 9 + tcpAsyClBuffer[resultptr + 8];
                        //break;

                    default:
                    // Any other write operation returns address written to
                        data = new byte[2];
                        Array.Copy(_tcpAsyClBuffer, resultptr + 10, data, 0, 2);
                        resultptr = resultptr + 10;
                        break;
                }
                if (PacketArrived != null)
                {
                    PacketArrived(new AsyncResultEventArgs<byte[]>(data));
                }
            }
            _bytesReceived = 0;
            ContinueReading();
        }
    }
}
