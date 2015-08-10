using System;
using System.Globalization;
using System.Text;

namespace ModbusExcel
{
    public class Conversion
    {

        #region Datatype

        /// <summary>
        /// Accept byte array and return result as string. Will be overloaded for other result datatypes. 
        /// </summary>
        /// <returns>ASCII representation of passed value</returns>
        public static string Datatype(ushort addr)
        {
            string datatype = "?";
                 if (addr >= 1 && addr <= 1) { datatype = "M"; } // Modicon Custom Ascii 248 Map
            else if (addr >= 200 && addr <= 201) { datatype = "M"; } // Modicon Custom Ascii 248 Map
            else if (addr >= 400 && addr <= 401) { datatype = "M"; } // Modicon Custom Ascii 248 Map
            else if (addr >= 700 && addr <= 712) { datatype = "R"; } // Ascii 128 Raw Data MAP

            else if (addr >= 1000 && addr <= 2999) { datatype = "B"; } // boolean 8 bit 1 byte
            else if (addr >= 3000 && addr <= 3039) { datatype = "I"; } // int 16 2 bytes
            else if (addr >= 4000 && addr <= 4999) { datatype = "A"; } // Ascii 8 bytes
            else if (addr >= 5000 && addr <= 5999) { datatype = "J"; } // int32 4 bytes *2
            else if (addr >= 6000 && addr <= 8999) { datatype = "F"; } // IEEE32 4 Bytes

            else if (addr >= 9000 && addr <= 9999) { datatype = "T"; } // Text ASCII Report 0-8196 in multiple 249 byte block responses
            else if (addr >= 13000 && addr <= 13497) { datatype = "I"; } // Int 16 2 bytes
            else if (addr >= 13500 && addr <= 13993) { datatype = "D"; } // RDA int 16 array 128 bytes
            else if (addr >= 14000 && addr <= 15000) { datatype = "S"; } // Ascii 16 8 Bytes
            else if (addr >= 15000 && addr <= 15899) { datatype = "J"; } // int32 4 bytes *4
            else if (addr >= 17000 && addr <= 18999) { datatype = "F"; } // IEEE32 4 Bytes
            else if (addr >= 40000 && addr <= 65535) { datatype = "I"; } // Int 16 2 bytes + sequence Value for simulation
            return datatype;
        }


        #endregion
        /// <summary>
        /// Parses raw Modbus read result based on datatypes and lengths
        /// </summary>
        /// <param name="addr">Register</param>
        /// <param name="length">Read Requested Length - not used at this time. Datatypes default lengths override</param>
        /// <param name="rawdata">Variable length byte of raw data</param>
        /// <param name="datatype">Character "M, R, B, C, I, A, J, F, T, I, D, S, J, F" - See Datatype()</param>
        /// <returns>String. Multiple values will be delimited by "|"</returns>
        public static string Parseresult(ushort addr, byte length, byte[] rawdata, string modicon, string datatype )
        {
            if (addr == 5101)
            {
                addr = 5101; // for debug breakpoint
            }
            string value = "";
            switch (datatype)
            {
                case "A": // ASCII
                case "S": // ASCII
                    // Loop thorugh results, also ensure result is multiple of length Modbus response for datatype should be
                        for (int i = 0; i < rawdata.Length && rawdata.Length >= (i + 7); i = i + 8)
                    {
                        if (value != "") value += "|";
                        byte[] dta = { rawdata[i], rawdata[i + 1], rawdata[i + 2], rawdata[i + 3], rawdata[i + 4], rawdata[i + 5], rawdata[i + 6], rawdata[i + 7] };
                        value = value + Encoding.ASCII.GetString(dta);
                    }
                    break;
                case "I": // 16Bit Integer
                        for (int i = 0; i < rawdata.Length && rawdata.Length >= (i + 1); i = i + 2)
                    {
                        if (value != "") value += "|";
                        byte[] dta = { rawdata[i], rawdata[i+1] };
                        var word = new int[1];
                        byte[] bytes = { dta[1], dta[0] };
                        word[0 / 2] = BitConverter.ToInt16(bytes, 0);
                        value = value + Convert.ToString(word[0]);
                    }
                    break;
                case "J": 
                // 32Bit IEEE Floating Point
                    if (modicon == "n")
                    {
                            for (int i = 0; i < rawdata.Length && rawdata.Length >= (i + 3); i = i + 4)
                        {
                            if (value != "") value += "|";
                            var dbl =
                                (double)
                                (rawdata[i]*16777216 + rawdata[i + 1]*65536 + rawdata[i + 2]*256 + rawdata[i + 3]);
                            value = value + dbl.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < rawdata.Length && rawdata.Length >= (i + 3); i = i + 4)
                        {
                            if (value != "") value += "|";
                            var dbl =
                                (double)
                                (rawdata[i]*16777216 + rawdata[i + 1]*65536 + rawdata[i + 2]*256 + rawdata[i + 3]);
                            value = value + dbl.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    break;
                case "F":
                    // 3Bit IEEE 2s compliment
                    if (modicon == "n")
                    {
                        for (int i = 0; i < rawdata.Length && rawdata.Length >= (i + 3); i = i + 4)
                        {
                            if (value != "") value += "|";
                            byte[] nbytes = {rawdata[i + 3], rawdata[i + 2], rawdata[i + 1], rawdata[i + 0]};
                            var result1 = BitConverter.ToSingle(nbytes, 0);
                            value = value + result1.ToString(CultureInfo.InvariantCulture);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < rawdata.Length && rawdata.Length >= (i + 3); i = i + 4)
                        {
                            if (value != "") value += "|";
                            byte[] nbytes = { rawdata[i + 1], rawdata[i + 0], rawdata[i + 3], rawdata[i + 2] };
                            var result1 = BitConverter.ToSingle(nbytes, 0);
                            value = value + result1.ToString(CultureInfo.InvariantCulture);
                        }                        
                    }
                    break;
                case "D": // (Debug)Raw Hexadecimal
                case "R":
                    StringBuilder sb = new StringBuilder(rawdata.Length * 2);
                    foreach (byte b in rawdata)
                    {
                        sb.AppendFormat("{0:x2}", b);
                    }
                    value = sb.ToString();
                    break;
                case "T": // Text
                    value = Encoding.ASCII.GetString(rawdata);
                    break;
                default: // Raw Data No conversion
                    value = Convert.ToString(rawdata);
                    break;
            }
            return value;
        }

        /// <summary>
        /// Parses raw Modbus read result based on datatypes and lengths
        /// </summary>
        /// <param name="addr">Register</param>
        /// <param name="length">Read Requested Length - not used at this time. Datatypes default lengths override</param>
        /// <param name="rawdata">Variable length byte of raw data</param>
        /// <returns>String. Multiple values will be delimited by "|"</returns>
        public static string Parseresult(ushort addr, byte length, byte[] rawdata, string modicon)
        {
            var datatype = Conversion.Datatype(addr);
            return Parseresult(addr, length, rawdata, modicon, datatype);
        }
    }
}
