using System;
using System.Collections.Generic;
using System.Text;

namespace Server
{

    public class ByteBuffer
    {
        #region Internal Declarations

        /* Internal _Declarations */

        // Default size of a data buffer (10 Kbs)
        private const int DATA_BUFFER = 10240;

        // The byte data
        private byte[] Data;

        // How big the data array is
        private int queueCapacity = 0;

        // How far into the data array have we written
        private int queueLength = 0;

        #endregion

        #region Constructors

        public ByteBuffer()
        {
            queueCapacity = DATA_BUFFER;
            Data = new byte[DATA_BUFFER - 1];
        }

        #endregion

        #region Properties

        public int Capacity
        {
            get { return queueLength; }
            set { queueCapacity = value; if (Length > value) { queueLength = value; } Data = new byte[queueCapacity - 1]; }
        }

        public int Length
        {
            get { return queueLength; }
            set { queueLength = value; }
        }

        #endregion

        #region Methods

        private int RemoveData(int dataLength)
        {
            int minLength = min(dataLength, queueLength);

            if (minLength != queueCapacity)
            {
                Buffer.BlockCopy(Data, minLength, Data, 0, queueLength - minLength);
            }

            queueLength -= minLength;

            return minLength;
        }

        private int min(int val1, int val2)
        {
            if (val1 < val2)
            {
                return val1;
            }
            else
            {
                return val2;
            }
        }

        private bool checkData(int LenBuf)
        {
            if (LenBuf > queueLength)
            {
                Console.WriteLine("Error in ByteBuff ReadData - Data: {0} (Not enough data)", LenBuf);
                return false;
            }

            return true;
        }

        private bool checkSpace(int LenBuf)
        {
            if (queueCapacity - queueLength - LenBuf < 0)
            {
                Console.WriteLine("Error in ByteBuff WriteData - Data: {0} (Not enough data)", LenBuf);
                return false;
            }

            return true;
        }

        public void WriteBlock(byte[] value)
        {
            if (checkSpace(value.Length) == false) { return;}

            Buffer.BlockCopy(value, 0, Data, queueLength, value.Length);

            queueLength += value.Length;
        }

        public void WriteByte(byte value)
        {
            byte[] buf = new byte[1];

            Buffer.SetByte(buf, 0, value);

            if (checkSpace(1) == false) { return; }

            Buffer.BlockCopy(buf, 0, Data, queueLength, 1);

            queueLength++;
        }

        public void WriteShort(short value)
        {
            short[] buf = new short[1];

            buf[0] = value;

            if (checkSpace(2) == false) { return; }

            Buffer.BlockCopy(buf, 0, Data, queueLength, 2);

            queueLength += 2;
        }

        public void WriteInt(int value)
        {
            int[] buf = new int[1];

            buf[0] = value;

            if (checkSpace(3) == false) { return; }

            Buffer.BlockCopy(buf, 0, Data, queueLength, 3);

            queueLength += 3;
        }

        public void WriteSingle(Single value)
        {
            Single[] buf = new Single[1];

            buf[0] = value;

            if (checkSpace(3) == false) { return; }

            Buffer.BlockCopy(buf, 0, Data, queueLength, 3);

            queueLength += 3;
        }

        public byte ReadByte()
        {
            byte[] buf = new byte[1];

            if (checkData(1) == false) { return 0; }

            Buffer.BlockCopy(Data, 0, buf, 0, 1);

            RemoveData(1);

            return buf[0];
        }

        public short ReadShort()
        {
            short[] buf = new short[2];

            if (checkData(2) == false) { return 0; }

            Buffer.BlockCopy(Data, 0, buf, 0, 2);

            RemoveData(2);

            return buf[0];
        }

        public int ReadInt()
        {
            int[] buf = new int[3];

            if (checkData(3) == false) { return 0; }

            Buffer.BlockCopy(Data, 0, buf, 0, 3);

            RemoveData(3);

            return buf[0];
        }

        public Single ReadSingle()
        {
            Single[] buf = new Single[3];

            if (checkData(3) == false) { return 0; }

            Buffer.BlockCopy(Data, 0, buf, 0, 3);

            RemoveData(3);

            return buf[0];
        }

        #endregion

    }
}
