using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeOpenXml.Utils.RemoteCalls
{
    internal class RemoteCallManager
    {
        static object _syncRoot = new object();
        internal List<RemoteTask> _tasks=new List<RemoteTask>();
        internal Dictionary<string, Queue<FormulaCellAddress>> _waitingToFinish = new Dictionary<string, Queue<FormulaCellAddress>>();
        public void QueueTask(RemoteTask task)
        {
            _tasks.Add(task);
            AnyTasks = true;
            if(task is HttpRemoteTask hrt)
            {
                QueueWaitingToFinish(hrt.Url, hrt.Cell);
            }
            ThreadPool.QueueUserWorkItem(_ => task.DoWork());
        }
        public void TaskComplate(RemoteTask task)
        {
            _tasks.Remove(task);
        }

        internal void QueueWaitingToFinish(string url, FormulaCellAddress currentCell)
        {
            if (_waitingToFinish.TryGetValue(url, out var queue) == false)
            {
                queue = new Queue<FormulaCellAddress>();
                _waitingToFinish[url] = queue;
            }
            queue.Enqueue(currentCell);
        }

        public bool HasRunningTasks
        {
            get { return _tasks.Count > 0; }
        }
        public bool AnyTasks { get; set; }
    }
}
