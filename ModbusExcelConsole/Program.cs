using System;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;

namespace ModbusExcelConsole
{
    /// <summary>
    /// This console application is a client for an
    /// Excel real-time data (RTD) server. It works
    /// by emulating the low level method calls
    /// and interactions that Excel makes when
    /// using a RTD.
    /// </summary>
    static class Program
    {
        // ProgIDs for COM components.

        private const String RTDProgID = "ModbusExcel.RTD";
        private const String RTDUpdateEventProgID = "ModbusExcel.UpdateEvent";

        static void Main()
        {
            Console.WriteLine("Console test RTD server.");
            TestModbusExcel(RTDProgID, RTDUpdateEventProgID);

            Console.WriteLine("Press enter to exit...");
            Console.ReadLine();
        }

        // Test harness that emulates the interaction of
        // Excel with an RTD server.
        static void TestModbusExcel(String rtdID, String eventID)
        {
            try
            {
                Type rtd = Type.GetTypeFromProgID(rtdID);
                object rtdServer = Activator.CreateInstance(rtd);
                Console.WriteLine("RTD.CreateInstance: {0}", rtdServer);

                // Create a callback event.
                Type update = Type.GetTypeFromProgID(eventID);
                var updateEvent = Activator.CreateInstance(update);
                Console.WriteLine("RTD.UpdateEvent: {0}", updateEvent);

                // Start the RTD server passing in the callback object.
                var param = new Object[1];
                param[0] = updateEvent;
                MethodInfo method = rtd.GetMethod("ServerStart");
                object ret = method.Invoke(rtdServer, param);
                Console.WriteLine("RTD.ServerStart: {0}", ret);

                // Request data from the RTD server.

                method = rtd.GetMethod("ConnectData");
                method.Invoke(rtdServer, new object[]
                {
                    2, new object[]
                    {
                        "127.0.0.1", // Localhost (ModSim.exe?)
                        "1", // UnitId
                        "N", // Modicon = R/N/Y/D/S = R?/No/Yes/Disabled/ModbusRTU with CRC
                        "4802", // Address (4802 is an 8 byte Omni FC Time register)
                        "1", // Number of registers to read (length) 
                        "PerRegister", // Make a new connection PerRegister (makes no difference for this tests single register read)
                        "1502" // PortNumber that ModbusSim or some other Modbus server/device is listening to. 502 is usually the default
                    },
                    true
                });
                // See documentation for various debug levels
                method.Invoke(rtdServer, new object[] { 2, new object[] { "DebugLevel", "512" }, true });

                // Make RTD Server poll once per second
                method.Invoke(rtdServer, new object[] { 2, new object[] { "PollRate", "1000" }, true });

                // Our own Startup service. Enable server as the last thing we do.  This is because it starts out disabled, 
                // giving the client (us or Excel) a chance to completly finish loading all topics before starting to poll.
                method.Invoke(rtdServer, new object[] { 2, new object[] { "ServerDisable", "false" }, true });

                // Loop and wait for RTD to notify (via callback) that
                // data is available.
                int count = 0;
                do
                {
                    // Wait for 2 seconds before getting
                    // more data from the RTD server. This
                    // it the default update period for Excel.
                    // This client can request data at a
                    // much higher frequency if wanted. 

                    Thread.Sleep(2000);
                    Console.WriteLine("RTD.Heartbeat: {0}", rtd.GetMethod("Heartbeat").Invoke(rtdServer, null));

                    var retval = (Object[,])rtd.GetMethod("RefreshData").Invoke(rtdServer, new object[] { 1 });

                    Console.WriteLine("RTD.RefreshData: {0}", (retval[1, 0].ToString().Length > 1) ? Encoding.Default.GetString(StringToByteArray(retval[1, 0].ToString())) : retval[1, 0]);

                } while (++count < 5); // Loop 5 times for test.

                // Our own shutdown service. Disable all polling, close all device connections.  Not required but maybe safer.
                rtd.GetMethod("ConnectData").Invoke(rtdServer, new object[] { 2, new object[] { "ServerDisable", "true" }, true });


                // Disconnect from data topic.  This is a built-in method required by the RTD protocol.
                rtd.GetMethod("DisconnectData").Invoke(rtdServer, new object[] { 1 });

                // Shutdown the RTD server. This is a built-in method required by the RTD protocol.
                rtd.GetMethod("ServerTerminate").Invoke(rtdServer, null);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0} ", e.Message);
            }
        }

        /// <summary>
        /// Helper to convert hex response to ASCII
        /// </summary>
        /// <param name="hex"></param>
        /// <returns></returns>
        private static byte[] StringToByteArray(string hex)
        {
            return Enumerable.Range(0, hex.Length).Where(x => x % 2 == 0).Select(x => Convert.ToByte(hex.Substring(x, 2), 16)).ToArray();
        }
    }
}
