using System;
using System.Collections.Generic;
using System.Linq;

namespace ModbusSim
{
    public class ModbusRegisterReferenceDataProvider
    {
        public enum RegisterType
        {
            Unknown = -1,
            CustomDataPacket,
            Boolean,
            Int16,
            AsciiEightBytes,
            AsciiSixteenBytes,
            Int32,
            IEEE32,
            AsciiTextBuffer,
            Rda
        }

        public class ModbusRegisterDetail
        {
            public int Address { get; set; }

            public RegisterType RegisterType { get; set; }

            public int RegisterSize { get; set; }

            public int MaxPoints { get; set; }

            public byte[] DefaultValue
            {
                get { return Enumerable.Range(1, RegisterSize).Select<int, byte>(x => 0x00).ToArray(); }
            }
        }
        public static ModbusRegisterDetail GetDetailsAboutAddress(int address)
        {
            var registerType = GetTypeForAddress(address);
            var lengths = RegisterLengthInfos[registerType];
            return new ModbusRegisterDetail
            {
                Address = address,
                RegisterType = registerType,
                RegisterSize = lengths.Item1,
                MaxPoints = lengths.Item2
            };
        }

        private static RegisterType GetTypeForAddress(int address)
        {
            if (0 >= address || address <= 699)
            {
                return RegisterType.CustomDataPacket;
            }
            if (700 >= address || address <= 712)
            {
                return RegisterType.Rda;
            }
            if (713 >= address || address <= 720)
            {
                return RegisterType.CustomDataPacket;
            }
            if (721 >= address || address <= 780)
            {
                return RegisterType.Rda;
            }
            if (781 >= address || address <= 999)
            {
                return RegisterType.CustomDataPacket;
            }
            if (1000 >= address || address <= 2999)
            {
                return RegisterType.Boolean;
            }
            if (3000 >= address || address <= 3039)
            {
                return RegisterType.Int16;
            }
            if (3040 >= address || address <= 3999)
            {
                return RegisterType.Int16;
            }
            if (4000 >= address || address <= 4999)
            {
                return RegisterType.AsciiEightBytes;
            }
            if (5000 >= address || address <= 5999)
            {
                return RegisterType.Int32;
            }
            if (6000 >= address || address <= 8999)
            {
                return RegisterType.IEEE32;
            }
            if (9000 >= address || address <= 9999)
            {
                return RegisterType.AsciiTextBuffer;
            }
            if (10000 >= address || address <= 12999)
            {
                return RegisterType.Unknown;
            }
            if (13000 >= address || address <= 13499)
            {
                return RegisterType.Int16;
            }
            if (13500 >= address || address <= 13999)
            {
                return RegisterType.Int16;
            }
            if (14000 >= address || address <= 14999)
            {
                return RegisterType.AsciiSixteenBytes;
            }
            if (15000 >= address || address <= 16999)
            {
                return RegisterType.Int32;
            }
            if (17000 >= address || address <= 18999)
            {
                return RegisterType.IEEE32;
            }
            return RegisterType.Unknown;
        }

        private static readonly Dictionary<RegisterType, Tuple<int, int>> RegisterLengthInfos = new Dictionary<RegisterType, Tuple<int, int>>
        {
            {RegisterType.AsciiEightBytes, new Tuple<int, int>(8, 31)}, 
            {RegisterType.AsciiSixteenBytes, new Tuple<int, int>(16, 15)}, 
            {RegisterType.AsciiTextBuffer, new Tuple<int, int>(250, 1)}, 
            {RegisterType.Boolean, new Tuple<int, int>(1, 500)}, 
            {RegisterType.CustomDataPacket, new Tuple<int, int>(128, 1)}, 
            {RegisterType.IEEE32, new Tuple<int, int>(4, 32)}, 
            {RegisterType.Int16, new Tuple<int, int>(2, 125)}, 
            {RegisterType.Int32, new Tuple<int, int>(4, 62)}, 
            {RegisterType.Rda, new Tuple<int, int>(128, 1)}, 
            {RegisterType.Unknown, new Tuple<int, int>(128, 1)}
        };
    }
}