using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

using Server.Core;

namespace Server.Game
{
    public class Time
    {

        private static bool prgRun = false;

        private static int tMultiply = 2;

        private struct gameTime
        {
            public int milliseconds, second, minute, hour, day, month, year;
        }; private static gameTime srvTime;


        public static void RequestRoutine()
        {
            int countTime = 0, precountTime = 0;

            while (prgRun)
            {
                precountTime = Environment.TickCount;

                srvTime.milliseconds = srvTime.milliseconds + (countTime * tMultiply);

                srvTime.second += srvTime.milliseconds / 1000; srvTime.milliseconds %= 1000;
                srvTime.minute += (srvTime.second / 60); srvTime.second %= 60;
                srvTime.hour += (srvTime.minute / 60); srvTime.minute %= 60;
                srvTime.day += (srvTime.hour / 24); srvTime.hour %= 24;
                srvTime.month += (srvTime.day / 31); srvTime.day %= 31;
                srvTime.year += (srvTime.month / 12); srvTime.month %= 12;

                countTime = Environment.TickCount - precountTime;

            }
        }

        public static void RequestConsultation()
        {
            Console.WriteLine("Horary: {0} {3}:{2}:{1} - {4}/{5}/{6}",
                                srvTime.milliseconds, srvTime.second, srvTime.minute, srvTime.hour,
                                 srvTime.day, srvTime.month, srvTime.year);
        }

        public static void Loader(bool load)
        {
            string FilePath = Environment.CurrentDirectory + @"\data\time.bin";

            if (load)
            {
                if (System.IO.File.Exists(FilePath) != false)
                {
                    byte[] buffer = new byte[Marshal.SizeOf(srvTime)];

                    BinaryReader Reader = new BinaryReader(new FileStream(FilePath, FileMode.Open, FileAccess.Read));
                   
                    Reader.Read(buffer, 0, buffer.Length);

                    GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);

                    try
                    {
                        srvTime = (gameTime)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(gameTime));
                    }

                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }

                    handle.Free();
                    
                    Reader.Close();
                    Reader = null;

                    GC.Collect();

                    prgRun = true;

                    Threading.ThreadRun(RequestRoutine, ThreadPriority.Normal, "Thread Time");

                    return;
                }

            }
            else
            {
                BinaryWriter Writer;

                if (System.IO.File.Exists(FilePath) != true)
                    Writer = new BinaryWriter(new FileStream(FilePath, FileMode.Create, FileAccess.Write));
                else
                    Writer = new BinaryWriter(new FileStream(FilePath, FileMode.Open, FileAccess.Write));

                byte[] buf = Helper.StructToByteArray(srvTime);

                Writer.Write(buf);

                Writer.Close(); Writer = null;

                prgRun = false;
                return;
            }

            prgRun = false;
        }
    }
}
