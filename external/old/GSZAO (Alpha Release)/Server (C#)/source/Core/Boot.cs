using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

using Server.Game;

namespace Server.Core
{
    class Boot
    {
        static void Main(string[] args)
        {
            Threading.ThreadRun(ConsoleIO.RequestRoutine, ThreadPriority.Highest, "Therad Console");
        }

    }
}
