using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeOpenXml.Utils.RemoteCalls
{
    internal static class RemoteCallManager
    {
        public static void QueueTask(RemoteTask task)
        {
            ThreadPool.QueueUserWorkItem(_ => task.DoWork());
        }
    }
}
