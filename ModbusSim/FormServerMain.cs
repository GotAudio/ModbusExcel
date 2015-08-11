using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Net;
using System.Net.Sockets;
using Nito.Async;
using Nito.Async.Sockets;

namespace ModbusSim
{
    public partial class FormServerMain : Form
    {
        public FormServerMain()
        {
            InitializeComponent();
            Showip();
        }

        private static Color _oddReadColor = Color.Green;
        private static Color _evenReadColor = Color.Red;
        public static int ConnectionCount = 0;
        public static int ReadRequests = 0;
        public static int SentResponses = 0;
        private static int[] Checkboxes = new int[960];
        private Bitmap _offscreenBitMap = Images._930_grid_unchecked;
        private static readonly int[] Connections = new int[40000];
        private const int GridSize = 930;

        private readonly Bitmap _930_three_state_unselected_indeterminate_grey_bg_odd = ChangeColor(Images._930_three_state_unselected_indeterminate_grey_bg, _oddReadColor.B);
        private readonly Bitmap _930_three_state_unselected_checked_grey_bg_odd = ChangeColor(Images._930_three_state_unselected_checked_grey_bg, _oddReadColor.B);
        private readonly Bitmap _930_three_state_unselected_checked_white_bg_odd = ChangeColor(Images._930_three_state_unselected_checked_white_bg, _oddReadColor.B);
        private readonly Bitmap _930_three_state_unselected_indeterminate_white_bg_odd = ChangeColor(Images._930_three_state_unselected_indeterminate_white_bg, _oddReadColor.B);

        private readonly Bitmap _930_three_state_unselected_indeterminate_grey_bg_even = ChangeColor(Images._930_three_state_unselected_indeterminate_grey_bg, _evenReadColor.R);
        private readonly Bitmap _930_three_state_unselected_checked_grey_bg_even = ChangeColor(Images._930_three_state_unselected_checked_grey_bg, _evenReadColor.R);
        private readonly Bitmap _930_three_state_unselected_checked_white_bg_even = ChangeColor(Images._930_three_state_unselected_checked_white_bg, _evenReadColor.R);
        private readonly Bitmap _930_three_state_unselected_indeterminate_white_bg_even = ChangeColor(Images._930_three_state_unselected_indeterminate_white_bg, _evenReadColor.R);

        /// <summary>
        /// The socket that listens for connections. This is null if we are not listening.
        /// </summary>
         private ServerTcpSocket _listeningSocket;

        /// <summary>
        /// The state of a child socket connection.
        /// </summary>
        private enum ChildSocketState
        {
            /// <summary>
            /// The child socket has an established connection.
            /// </summary>
            Connected,

            /// <summary>
            /// The child socket is disconnecting.
            /// </summary>
            Disconnecting
        }

        /// <summary>
        /// Keeps state information for each child socket connection.
        /// </summary>
        private class ChildSocketContext
        {
            /// <summary>
            /// The socket's packetizer, used for reading the socket.
            /// </summary>
            public ModbusPacketProtocols Protocol { get; set; }

            /// <summary>
            /// State of the socket.
            /// </summary>
            public ChildSocketState State;
            public int ConnectionNumber;
        }

        /// <summary>
        /// Size of initial _topics ConcurrentDictionary, and number of Socket Connections (MBMaster objects).  
        /// </summary>
        private const int InitialCapacity = 40000;

        /// <summary>
        /// Number of CPU processsors as queried from environment
        /// </summary>
        static readonly int NumProcs = Environment.ProcessorCount;

        /// <summary>
        /// Used by _topics ConcurrentDictionary operations. CPU * 2
        /// </summary>       
        static readonly int ConcurrencyLevel = NumProcs * 2;

        /// <summary>
        /// A mapping of sockets (with established connections) to their state.
        /// </summary>
        private readonly ConcurrentDictionary<ServerChildTcpSocket, ChildSocketContext> _childSockets = new ConcurrentDictionary<ServerChildTcpSocket, ChildSocketContext>(ConcurrencyLevel, InitialCapacity);

        /// <summary>
        /// Closes and clears the listening socket and all connected sockets, without causing exceptions.
        /// </summary>
        private void ResetListeningSocket()
        {
            // Close all child sockets
            foreach (KeyValuePair<ServerChildTcpSocket, ChildSocketContext> socket in _childSockets)
            {
                socket.Key.Close();
//                ConnectionCount--;
            }
            _childSockets.Clear();

            // Close the listening socket
            _listeningSocket.Close();
            _listeningSocket = null;
        }
        /// <summary>
        /// Closes and clears a child socket (established connection), without causing exceptions.
        /// </summary>
        /// <param name="childSocket">The child socket to close. May be null.</param>
        private void ResetChildSocket(ServerChildTcpSocket childSocket)
        {

            Connections[_childSockets[childSocket].ConnectionNumber] = 0;
            ShowDisconnected(_childSockets[childSocket]);

            // Close the child socket if possible
            ConnectionCount--;
            if (childSocket != null)
                childSocket.Close();

            // Remove it from the list of child sockets
            // kas concurrentdictionarry do not use .remove
            //            ChildSockets.Remove(childSocket);
        }

        private void RefreshDisplay()
        {
            // If the server socket is running, don't allow starting it; if it's not, then don't allow stopping it
            buttonStart.Enabled = (_listeningSocket == null);
            buttonStop.Enabled = (_listeningSocket != null);

            // Display status
            if (_listeningSocket == null)
                toolStripStatusLabel3.Text = @"Stopped";
            else
                toolStripStatusLabel3.Text = @"Listening on " + _listeningSocket.LocalEndPoint;

            toolStripStatusLabel4.Text = ConnectionCount + @" connections" + @" Reads:" + ReadRequests + @", Responses: " + SentResponses;
            
        }

        private void ListeningSocket_AcceptCompleted(AsyncResultEventArgs<ServerChildTcpSocket> e)
        {
            // Check for errors
            if (e.Error != null)
            {
                ResetListeningSocket();
                textBoxLog.AppendText("Socket error during Accept: [" + e.Error.GetType().Name + "] " + e.Error.Message + Environment.NewLine);
                RefreshDisplay();
                return;
            }

            // Always continue listening for other connections
            _listeningSocket.AcceptAsync();

            ServerChildTcpSocket socket = e.Result;
            try
            {
                // Save the new child socket connection, and create a packetizer for it
                var protocol = new ModbusPacketProtocols(socket);
                var context = new ChildSocketContext
                    {
                        Protocol = protocol,
                        State = ChildSocketState.Connected,
                        ConnectionNumber = ConnectionCount++
                    };
                _childSockets[socket] = context;
                protocol.PacketArrived += args => ChildSocket_PacketArrived(socket, args);
                socket.WriteCompleted += args => ChildSocket_WriteCompleted(socket, args);
                socket.ShutdownCompleted += args => ChildSocket_ShutdownCompleted(socket, args);

                // Display the connection information            
//                textBoxLog.AppendText("Connection established to " + socket.RemoteEndPoint.ToString() + Environment.NewLine);

                // Start reading data from the connection
                protocol.Start();
            }
            catch (Exception ex)
            {
                ResetChildSocket(socket);
                textBoxLog.AppendText("Socket error accepting connection: [" + ex.GetType().Name + "] " + ex.Message + Environment.NewLine);
            }
            finally
            {
                ShowConnection(_childSockets[socket]);
                RefreshDisplay();
            }
        }

        private void ChildSocket_PacketArrived(ServerChildTcpSocket socket, AsyncResultEventArgs<byte[]> e)
        {
            try
            {
                // Check for errors
                if (e.Error != null)
                {                                    
//                    textBoxLog.AppendText("Client socket error during Read from socket: [" + e.Error.GetType().Name + "] " + e.Error.Message + Environment.NewLine);
                    ResetChildSocket(socket);
                }
                else if (e.Result == null)
                {
                    // PacketArrived completes with a null packet when the other side gracefully closes the connection
                    //   textBoxLog.AppendText("Socket graceful close detected from " + socket.RemoteEndPoint.ToString() + Environment.NewLine);
                    // Close the socket and remove it from the list
                    ResetChildSocket(socket);
                }
                else
                {
                    // At this point, we know we actually got a message.

                    // Deserialize the message
                    //                    object message = Util.Deserialize(e.Result);
                    //                    ByteMessage byteMessage = message as ByteMessage;
                    byte[] rawdata = e.Result; // byteMessage.Message;

                    //                    StringMessage xstringMessage = message as StringMessage;

                    //var sb = new StringBuilder(rawdata.Length * 2);
                    //foreach (byte b in rawdata)
                    //{
                    //    sb.AppendFormat("{0:x2}", b);
                    //}
                    //string value = sb.ToString();
                    //     textBoxLog.AppendText("Socket read got a binary message from " + socket.RemoteEndPoint + ": " + value + Environment.NewLine);

                    try
                    {
                        if (rawdata.Length == 0)
                        {
                            return;
                        }
                    }
                    catch (SocketException)
                    {
                        return;
                    }

                    //int resultptr = 0;

                    //while (resultptr < nBytesReceived)
                    //{
                    //byte[] data;
                    //ushort id = SwapUInt16(BitConverter.ToUInt16(rawdata, resultptr));
                    //byte unit = rawdata[resultptr + 6];
                    //byte function = rawdata[resultptr + 7];        

                    SetCheck((_childSockets[socket].ConnectionNumber % GridSize), ((_childSockets[socket].ConnectionNumber) / GridSize & 1) != 0 ? "IndeterminateRead" : "CheckedRead");
                    
                    SendResponse(socket, rawdata);
                }
            }
            catch (Exception)
            {
//                textBoxLog.AppendText("Error reading from socket " + socket.RemoteEndPoint + ": [" + ex.GetType().Name + "] " + ex.Message + Environment.NewLine);
                ResetChildSocket(socket);
            }
            finally
            {
                RefreshDisplay();
            }
        }

        private void SendResponse(ServerChildTcpSocket socket, byte[] data)
        {
            byte[] bytelength;

            var adr = data[8] * 256 + data[9];
            var reqlen = (int)data[10] * 256 + data[11];
            double sizeNeeded = reqlen;
            var modbusDetails = ModbusRegisterReferenceDataProvider.GetDetailsAboutAddress(adr);
            if (modbusDetails.RegisterType == ModbusRegisterReferenceDataProvider.RegisterType.Boolean)
            {
                sizeNeeded = ((double) reqlen/(double) 8) + 1;
                if (sizeNeeded - (int) sizeNeeded == 0)
                    sizeNeeded = sizeNeeded + 1;
            }
            var mbap = new byte[9 + modbusDetails.RegisterSize * (int)(sizeNeeded)];
            mbap[0] = data[0]; // Transaction
            mbap[1] = data[1]; // Transaction
            mbap[6] = data[6]; // Unit
            mbap[7] = data[7]; // Function code

            switch (modbusDetails.RegisterType) // Function Code
            {
                case ModbusRegisterReferenceDataProvider.RegisterType.Boolean:
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize));
                    mbap[8] = bytelength[3]; // 2; // Response length
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize + 3));
                    mbap[5] = bytelength[3]; // 5; // Message size
                    for (int i = 0; i < (int)sizeNeeded; i++)
                    {
                        mbap[9 + (modbusDetails.RegisterSize * i) ] = 0x01; // 
                    }
                    break;
                case ModbusRegisterReferenceDataProvider.RegisterType.IEEE32:
                case ModbusRegisterReferenceDataProvider.RegisterType.Int32:
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize));
                    mbap[8] = bytelength[3]; // 2; // Response length
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize + 3));
                    mbap[5] = bytelength[3]; // 5; // Message size
                    for (int i = 0; i < sizeNeeded; i++)
                    {
                        mbap[9 + (modbusDetails.RegisterSize * i) ] = 0x43; // 
                        mbap[10 + (modbusDetails.RegisterSize * i)] = 0x39; // 
                        mbap[11 + (modbusDetails.RegisterSize * i)] = 0xE4; // 
                        mbap[12 + (modbusDetails.RegisterSize * i)] = 0x15; // 
                    }
                    break;
                case ModbusRegisterReferenceDataProvider.RegisterType.Int16:
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize));
                    mbap[8] = bytelength[3]; // 2; // Response length
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize + 3));
                    mbap[5] = bytelength[3]; // 5; // Message size
                    for (int i = 0; i < sizeNeeded; i++)
                    {
                        mbap[9 + (modbusDetails.RegisterSize * i) ] = 0x13; // 
                        mbap[10 + (modbusDetails.RegisterSize * i) ] = 0xF5; // 
                    }
                    break;
                case ModbusRegisterReferenceDataProvider.RegisterType.AsciiEightBytes:
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize));
                    mbap[8] = bytelength[3]; // 2; // Response length
                    bytelength = BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)sizeNeeded * modbusDetails.RegisterSize + 3));
                    mbap[5] = bytelength[3]; // 5; // Message size

                    // IsADate retister
                    if (adr == 4801 || adr == 4807 || adr == 4847 || adr == 4103 || adr == 4131 || adr == 4133 || adr == 4141 || adr == 4143 || adr == 4203 || adr == 4231 || adr == 4233 || adr == 4241 || adr == 4243 || adr == 4303 || adr == 4331 || adr == 4333 || adr == 4341 || adr == 4343 || adr == 4403 || adr == 4431 || adr == 4433 || adr == 4441 || adr == 4443 )
                    {
                        for (int i = 0; i < sizeNeeded; i++)
                        {
                            mbap[9 + (modbusDetails.RegisterSize*i)] = 0x30; // 0
                            mbap[10 + (modbusDetails.RegisterSize*i)] = 0x31; // 1
                            mbap[11 + (modbusDetails.RegisterSize*i)] = 0x2F; // /
                            mbap[12 + (modbusDetails.RegisterSize*i)] = 0x30; // 0
                            mbap[13 + (modbusDetails.RegisterSize*i)] = 0x31; // 1
                            mbap[14 + (modbusDetails.RegisterSize*i)] = 0x2F; // /
                            mbap[15 + (modbusDetails.RegisterSize*i)] = 0x30; // 0
                            mbap[16 + (modbusDetails.RegisterSize*i)] = 0x31; // 1
                        }
                    }
                    // IsATime retister
                    else if (adr == 4802 || adr == 4808 || adr == 4848 || adr == 4104 || adr == 4132 || adr == 4134 || adr == 4140 || adr == 4142 || adr == 4204 || adr == 4232 || adr == 4234 || adr == 4240 || adr == 4242 || adr == 4304 || adr == 4332 || adr == 4334 || adr == 4340 || adr == 4342 || adr == 4404 || adr == 4432 || adr == 4434 || adr == 4440 || adr == 4442)
                    {
                        for (int i = 0; i < sizeNeeded; i++)
                        {
                            mbap[9 + (modbusDetails.RegisterSize*i)] = 0x30; // 0
                            mbap[10 + (modbusDetails.RegisterSize*i)] = 0x31; // 1
                            mbap[11 + (modbusDetails.RegisterSize*i)] = 0x3A; // /
                            mbap[12 + (modbusDetails.RegisterSize*i)] = 0x30; // 0
                            mbap[13 + (modbusDetails.RegisterSize*i)] = 0x31; // 1
                            mbap[14 + (modbusDetails.RegisterSize*i)] = 0x3A; // /
                            mbap[15 + (modbusDetails.RegisterSize*i)] = 0x30; // 0
                            mbap[16 + (modbusDetails.RegisterSize*i)] = 0x31; // 1
                        }
                    }
                    else
                    {
                        for (int i = 0; i < sizeNeeded; i++)
                        {
                            mbap[9 + (modbusDetails.RegisterSize*i)] = 0x41; // 
                            mbap[10 + (modbusDetails.RegisterSize*i)] = 0x42; // 
                            mbap[11 + (modbusDetails.RegisterSize*i)] = 0x43; // 
                            mbap[12 + (modbusDetails.RegisterSize*i)] = 0x44; // 
                            mbap[13 + (modbusDetails.RegisterSize*i)] = 0x45; // 
                            mbap[14 + (modbusDetails.RegisterSize*i)] = 0x46; // 
                            mbap[15 + (modbusDetails.RegisterSize*i)] = 0x47; // 
                            mbap[16 + (modbusDetails.RegisterSize*i)] = 0x48; // 
                        }
                    }
                    break;
                default:
                    // Return register. Not sure this is acceptable.
                    mbap[5] = 5; // Message size
                    mbap[8] = 2; // Response length
                    mbap[9] = data[8]; // 
                    mbap[10] = data[9]; // 
                    break;
            }

            ReadRequests++;
            ModbusPacketProtocols.WritePacketAsync(socket, mbap);
        }

        private void ChildSocket_ShutdownCompleted(ServerChildTcpSocket socket, AsyncCompletedEventArgs e)
        {
            // Check for errors
            if (e.Error != null)
            {
                textBoxLog.AppendText("Socket error during Shutdown of " + socket.RemoteEndPoint + ": [" + e.Error.GetType().Name + "] " + e.Error.Message + Environment.NewLine);
                ResetChildSocket(socket);
            }
            else
            {
//                textBoxLog.AppendText("Socket shutdown completed on " + socket.RemoteEndPoint.ToString() + Environment.NewLine);
                // Close the socket and remove it from the list
                ResetChildSocket(socket);
            }

            RefreshDisplay();
        }

        private void ChildSocket_WriteCompleted(ServerChildTcpSocket socket, AsyncCompletedEventArgs e)
        {
            // Check for errors
            if (e.Error != null)
            {
                // Note: WriteCompleted may be called as the result of a normal write (SocketPacketizer.WritePacketAsync),
                //  or as the result of a call to SocketPacketizer.WriteKeepaliveAsync. However, WriteKeepaliveAsync
                //  will never invoke WriteCompleted if the write was successful; it will only invoke WriteCompleted if
                //  the keepalive packet failed (indicating a loss of connection).

                // If you want to get fancy, you can tell if the error is the result of a write failure or a keepalive
                //  failure by testing e.UserState, which is set by normal writes.
//                if (e.UserState is string)
                    textBoxLog.AppendText("Socket error during Write to " + socket.RemoteEndPoint + ": [" + e.Error.GetType().Name + "] " + e.Error.Message + Environment.NewLine);
  //              else
  //                  textBoxLog.AppendText("Socket error detected by keepalive to " + socket.RemoteEndPoint + ": [" + e.Error.GetType().Name + "] " + e.Error.Message + Environment.NewLine);

                ResetChildSocket(socket);
            }
            else
            {
                SentResponses++;
//            var description = (string)e.UserState;
//                textBoxLog.AppendText("Socket write completed to " + socket.RemoteEndPoint + " for message " + description + Environment.NewLine);
            }

            RefreshDisplay();
        }

        private void ButtonStartClick1(object sender, EventArgs e)
        {
            // Read the port number
            int port;
            if (!int.TryParse(textBoxPort.Text, out port))
            {
                MessageBox.Show(@"Invalid port number: " + textBoxPort.Text);
                textBoxPort.Focus();
                return;
            }

            try
            {
                _listeningSocket = new ServerTcpSocket();
                _listeningSocket.AcceptCompleted += ListeningSocket_AcceptCompleted;
                _listeningSocket.Bind(port, 40000);
                _listeningSocket.AcceptAsync();
                textBoxLog.AppendText("Listening on port " + port.ToString(CultureInfo.InvariantCulture) + Environment.NewLine);
            }
            catch (Exception ex)
            {
                ResetListeningSocket();
                textBoxLog.AppendText("Error creating listening socket on port " + port.ToString(CultureInfo.InvariantCulture) + ": [" + ex.GetType().Name + "] " + ex.Message + Environment.NewLine);
            }
            RefreshDisplay();
        }

        private void ButtonStopClick(object sender, EventArgs e)
        {
            // Close the listening socket cleanly
            ResetListeningSocket();
            RefreshDisplay();
        }

        private void Showip()
        {
            var sb = new StringBuilder();

            // Get a list of all network interfaces (usually one per network card, dialup, and VPN connection)
            NetworkInterface[] networkInterfaces = NetworkInterface.GetAllNetworkInterfaces();

            foreach (NetworkInterface network in networkInterfaces)
            {
                // Read the IP configuration for each network
                IPInterfaceProperties properties = network.GetIPProperties();

                // Each network interface may have multiple IP addresses
                foreach (UnicastIPAddressInformation address in properties.UnicastAddresses)
                {
                    // We're only interested in IPv4 addresses for now
                    if (address.Address.AddressFamily != AddressFamily.InterNetwork)
                        continue;

                    // Ignore loopback addresses (e.g., 127.0.0.1)
                    if (IPAddress.IsLoopback(address.Address))
                        continue;

                    sb.AppendLine(address.Address + " (" + network.Name + ")");
                }
            }
            textBoxLog.AppendText(sb.ToString());
        }

        //// Prevents flicker when adding controls.  Disable for now, but re-enable if slow controls become a problem.
        //// Enableing this makes all test checkboxes show up only after the loop that sets them is complete.
        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        // Activate double buffering at the form level.  All child controls will be double buffered as well.
        //        CreateParams cp = base.CreateParams;
        //        cp.ExStyle |= 0x02000000;   // WS_EX_COMPOSITED
        //        return cp;
        //    }
        //} 

        private void ShowConnection(ChildSocketContext context)
        {
            Connections[context.ConnectionNumber] = 1;
            int phase = (context.ConnectionNumber) / GridSize;
            int isodd = (phase & 1);
            SetCheck((context.ConnectionNumber % GridSize), isodd != 0 ? "Indeterminate" : "Checked");
        }

        private void ShowDisconnected(ChildSocketContext context)
        {
            Connections[context.ConnectionNumber] = 0;
            SetCheck((context.ConnectionNumber % GridSize), "Unchecked");
        }

        private void SetCheck(int connection, string setVal)
        {
            if (!checkbox_showevents.Checked) return;

            var i = connection % 30;
            var j = (connection - i) / 31;

            // textBoxLog.AppendText("Connecting: [" + (i+1).ToString() + "][" + (j+1).ToString() + "]" + Environment.NewLine);

            var toffscreenBitMap = _offscreenBitMap;
            Graphics offscreenGraphics = Graphics.FromImage(toffscreenBitMap);
            Graphics clientDc = panel1.CreateGraphics();

            int wo = i < 9 ? i * 21 : 9 * 21 + 2 + (i - 9) * 25;

            if (setVal == "CheckedRead")
                Checkboxes[connection] = Checkboxes[connection] == 1 ? 2 : 1;
            else
                if (setVal == "IndeterminateRead")
                    Checkboxes[connection] = Checkboxes[connection] == 1 ? 1 : 2;

            var newval = GetCheckbox(setVal, i % 2, Checkboxes[connection] % 2 == 1);

            offscreenGraphics.DrawImage(newval, (wo) + 47, (j * 18) + 20);
            _offscreenBitMap = toffscreenBitMap;
            clientDc.DrawImage(newval, (wo) + 47, (j * 18) + 20);
            panel1.BackgroundImage = _offscreenBitMap;
        }

        private Image GetCheckbox(string setVal, int i, bool oddEven)
        {
            if (i == 1)
            {
                switch (setVal)
                {
                    case "Checked":
                        return Images._930_three_state_unselected_checked_grey_bg;
                    case "CheckedRead":
                        return oddEven ? _930_three_state_unselected_checked_grey_bg_odd : _930_three_state_unselected_checked_grey_bg_even;
                    case "Indeterminate":
                        return Images._930_three_state_unselected_indeterminate_grey_bg;
                    case "IndeterminateRead":
                        return oddEven ? _930_three_state_unselected_indeterminate_grey_bg_odd : _930_three_state_unselected_indeterminate_grey_bg_even;
                    default:
                        return Images._930_three_state_unselected_unchecked_grey_bg;
                }
            }
            switch (setVal)
            {
                case "Checked":
                    return Images._930_three_state_unselected_checked_white_bg;
                case "CheckedRead":
                    return oddEven ? _930_three_state_unselected_checked_white_bg_odd : _930_three_state_unselected_checked_white_bg_even;
                case "Indeterminate":
                    return Images._930_three_state_unselected_indeterminate_white_bg;
                case "IndeterminateRead":
                    return oddEven ? _930_three_state_unselected_indeterminate_white_bg_odd : _930_three_state_unselected_indeterminate_white_bg_even;
                default:
                    return Images._930_three_state_unselected_unchecked_white_bg;
            }
        }

        private static Bitmap ChangeColor(Bitmap oldVal, byte oldColor)
        {
            byte r = oldColor;
            var bmp = new Bitmap(oldVal);
            for (int x = 0; x < bmp.Width; x++)
            {
                for (int y = 0; y < bmp.Height; y++)
                {
                    Color gotColor = bmp.GetPixel(x, y);
                    gotColor = Color.FromArgb(r, gotColor.G, gotColor.B);
                    bmp.SetPixel(x, y, gotColor);
                }
            }
            return bmp;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            var start = DateTime.Now;

            //Random r = new Random();
            for (int i = 0; i < 57; i++)
            {
                for (int j = 0; j < 960; j++) SetCheck(j, "Checked");
                for (int j = 0; j < 960; j++) SetCheck(j, "UnChecked");                
            }
            var stop = DateTime.Now;
            TimeSpan span = stop.Subtract(start);
            double sec = span.TotalSeconds;
            textBoxLog.AppendText(string.Format("Toggled 109,440 checkboxes in {0} seconds. At 110K reads per minute, GUI is {1:0.#}% of overhead." + Environment.NewLine, sec, (sec * 100 / 60)));
            return;
            //for (int j = 0; j < 960; j++) SetCheck(j, "Checked");
            //for (int j = 0; j < 960; j++) SetCheck(j, "UnChecked");
            //for (int j = 0; j < 960; j++) SetCheck(j, "CheckedRead");
            //for (int j = 0; j < 960; j++) SetCheck(j, "UnChecked");
            //for (int j = 0; j < 960; j++) SetCheck(j, "Indeterminate");
            //for (int j = 0; j < 960; j++) SetCheck(j, "UnChecked");
            //for (int j = 0; j < 960; j++) SetCheck(j, "IndeterminateRead");
            //for (int j = 0; j < 960; j++) SetCheck(j, "UnChecked");
        }
    }
}
