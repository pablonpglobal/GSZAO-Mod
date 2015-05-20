using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Server.Core
{
    class Config : BinaryStream
    {
        public Config()
            : base()
        {

        }

        public Config(string Directory)
            : base()
        {
            base.File = Directory;
        }
    }
}
