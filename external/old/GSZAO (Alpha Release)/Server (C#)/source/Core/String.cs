using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.IO;
using System.Runtime.InteropServices;


namespace Server.Core
{
    class hString
    {
        public static string ReadField(int pos, string param, string Asc)
        {

            try
            {

                int SepCount = 1;
                int[] SepPosition = new int[32]; SepPosition[0] = 0;

                for (int i = 1; i < param.Length; i++)
                {
                    if (Asc == param.Substring(i, 1))
                    {
                        SepPosition[SepCount] = i;
                        SepCount++;
                    }
                }

                SepPosition[SepCount] = param.Length;

                if (pos == 1)
                    return param.Substring(SepPosition[0], SepPosition[1] - SepPosition[0]);
                else
                    return param.Substring(SepPosition[pos - 1] + 1, SepPosition[pos] - SepPosition[pos - 1] - 1);

            }
            
            catch
            {
                return null;
            }

        }

        public static string Left(string param, int length)
        {
            return param.Substring(0, length);
        }

        public static string Right(string param, int length)
        {
            return param.Substring(param.Length - length, length);
        }

        public static string Mid(string param, int startIndex, int length)
        {
            return param.Substring(startIndex, length);
        }

        public static string Mid(string param, int startIndex)
        {
            return param.Substring(startIndex);
        }

        // IO..
        /*public static byte[] StructToByteArray(object _oStruct)
        {
            try
            {
                byte[] buffer = new byte[Marshal.SizeOf(_oStruct)];

                GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);
                Marshal.StructureToPtr(_oStruct, handle.AddrOfPinnedObject(), false);
                handle.Free();

                return buffer;
            }

            catch (Exception Ex)
            {
                throw Ex;
            }
        }*/
    }
}
