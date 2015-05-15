using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace Server.Core
{
    public abstract class Threading //Esta clase.. es un poco provisional, hay que mejorarla y hacer que vuele
    {

        public delegate void Function();

        public Threading()
        {

        }

        public static int ThreadState(ref Thread mThread)
        {
            return (int)mThread.ThreadState;
        }

        public static void ThreadRun(Function Func, ThreadPriority Priority)
        {
            Thread mThread = new Thread(new ThreadStart(Func));
            mThread.Priority = Priority;
            mThread.Start();
        }

        public static void ThreadRun(Function Func, ThreadPriority Priority, string mThreadName)
        {
            Thread mThread = new Thread(new ThreadStart(Func));
            mThread.Priority = Priority;
            mThread.Name = mThreadName;
            mThread.Start();
        }

        public void ThreadAbort(ref Thread mThread)
        {
            mThread.Abort();
        }

        public void ThreadPause(ref Thread mThread)
        {
            mThread.Suspend();
        }

        public void ThreadResume(ref Thread mThread)
        {
            mThread.Resume();
        }

        public void ThreadSleep()
        {
            Thread.Sleep(1);
        }

        public void ThreadSleep(int mSnd)
        {
            Thread.Sleep(mSnd);
        }
        
    }
}
