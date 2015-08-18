using System;
using System.Threading;

namespace UnitTests.Util
{
    public sealed class LoggingSynchronizationContext : SynchronizationContext
    {
        private SynchronizationContext synchronizationContext;

        public LoggingSynchronizationContext(SynchronizationContext synchronizationContext)
        {
            this.synchronizationContext = synchronizationContext;
        }

        public Action OnOperationCompleted { get; set; }
        public override void OperationCompleted()
        {
            if (OnOperationCompleted != null)
                OnOperationCompleted();
            synchronizationContext.OperationCompleted();
        }

        public Action OnOperationStarted { get; set; }
        public override void OperationStarted()
        {
            if (OnOperationStarted != null)
                OnOperationStarted();
            synchronizationContext.OperationStarted();
        }

        public Action OnPost { get; set; }
        public override void Post(SendOrPostCallback d, object state)
        {
            if (OnPost != null)
                OnPost();
            synchronizationContext.Post(d, state);
        }

        public Action OnSend { get; set; }
        public override void Send(SendOrPostCallback d, object state)
        {
            if (OnSend != null)
                OnSend();
            synchronizationContext.Send(d, state);
        }
    }
}
