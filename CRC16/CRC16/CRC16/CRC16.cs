using System;
using System.Runtime.InteropServices;

namespace CRC16
{
    [ComVisible(true)]
    [Guid("1B50B6D6-596D-48FF-96B6-C21F7C084A6E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICRC16Wrapper
    {
        string ComputeCRC16(string hexData);
    }

    [ComVisible(true)]
    [Guid("51D27C72-8700-43F6-ACA2-FD9428F99102")]
    [ClassInterface(ClassInterfaceType.None)]
    public class CRC16Wrapper : ICRC16Wrapper
    {
        private static readonly ushort[] Crc16Table;

        static CRC16Wrapper()
        {
            const ushort polynomial = 0x1021; // CRC-16-CCITT polynomial
            Crc16Table = new ushort[256];

            for (int i = 0; i < 256; ++i)
            {
                ushort crc = (ushort)(i << 8);
                for (int j = 0; j < 8; ++j)
                {
                    crc = (crc & 0x8000) != 0 ? (ushort)((crc << 1) ^ polynomial) : (ushort)(crc << 1);
                }
                Crc16Table[i] = crc;
            }
        }

        public string ComputeCRC16(string hexData)
        {
            byte[] data = StringToByteArray(hexData);
            ushort crcValue = 0xFFFF; // Start value for CRC-16-CCITT

            foreach (byte by in data)
            {
                byte tableIndex = (byte)((crcValue >> 8) ^ by);
                crcValue = (ushort)((Crc16Table[tableIndex] << 8) ^ (crcValue << 8));
            }

            

            return crcValue.ToString("X4"); // CRC16 produces a 16-bit value, represented here as a 4-digit hexadecimal number.
        }

        private byte[] StringToByteArray(string hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }
    }
}
