using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Server.Core
{
    class SrvConsole
    {
        private static bool _shouldStop = true;

        public static void RequestStop()
        {
            _shouldStop = false;
        }

        public static void RequestRoutine()
        {
            string Income, Command;

            while (_shouldStop)
            {
                Income = Console.ReadLine();
                Command = hString.ReadField(1, Income, " ");

                string arg;

                switch ((string)Command)
                {
                    case "help":

                        arg = hString.ReadField(2, Income, " ");

                        switch (arg)
                        {
                            case "1":
                                Console.WriteLine("option1");
                                break;

                            case "2":
                                Console.WriteLine("option2");
                                break;
                        }
                        
                        break;

                    case "time":
                        arg = hString.ReadField(2, Income, " ");

                        switch (arg)
                        {
                            case "-l": //load
                                //Time.Loader(true);
                                break;

                            case "-i": //info
                                //Time.RequestConsultation();
                                break;

                            case "-d": //destroy
                                //Time.Loader(false);
                                break;
                        }

                        break;

                    case "end":
                        RequestStop();
                        break;

                    default: break;
                }
            }
        }


    }
}
