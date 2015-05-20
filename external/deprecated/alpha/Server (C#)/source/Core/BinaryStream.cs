using System;
using System.IO;

namespace Server.Core
{
    public abstract class BinaryStream
    {
        private static string Archive = Environment.CurrentDirectory;

        public BinaryStream()
        {

        }

        public string File
        {
            get { return Archive; }
            set { Archive = Environment.CurrentDirectory + value; }
        }

        protected static void Create()
        {
            try
            {
                if (System.IO.File.Exists(Archive))
                {
                    Console.WriteLine("Stream: The file has been created");
                    return;
                }

                System.IO.File.Create(Archive);
                Console.WriteLine("Stream: File created successfully in {0}", Archive);
            }

            catch
            {
                Console.WriteLine("Stream: (ERROR) Wrong directory");
            }
        }

        protected static void Delete()
        {
            if (System.IO.File.Exists(Archive))
            {
                System.IO.File.Delete(Archive);
            }
        }

        protected static string Read()
        {
            System.IO.StreamReader Reader = new System.IO.StreamReader(Archive);

            try
            {
                return null;
            }

            catch (Exception Ex)
            {
                Console.WriteLine("Stream: Error in StreamReader: {0}", Ex.Message);
                return null;
            }
        }

        protected static void Write()
        {
            System.IO.StreamWriter Writer = new System.IO.StreamWriter(Archive);

            try
            {

            }

            catch (Exception Ex)
            {
                Console.WriteLine("Stream: Error in StreamWriter: {0}", Ex.Message);
            }
        }
    }
}
