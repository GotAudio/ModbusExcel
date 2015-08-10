using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Messaging;
using System.Text;

namespace ModbusExcel
{
    public class Messaging
    {
        /// <summary>
        /// Sends message to queue
        /// </summary>
        /// <param name="queue">Queue path</param>
        /// <param name="msg">Message</param>
        public static void Send(SqlString queue, SqlString msg)
        {
            using (MessageQueue msgQueue = new MessageQueue(queue.ToString(), QueueAccessMode.Send))
            {
                msgQueue.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });
                msgQueue.Send(msg.Value);
            }
        }

        public static void SendExisting(MessageQueue msgQueue, SqlString msg)
        {
//            using (MessageQueue msgQueue = new MessageQueue(queue.ToString(), QueueAccessMode.Send))
//            {
                msgQueue.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });
                msgQueue.Send(msg.Value);
//            }
        }

        public static MessageQueue GetConnection(SqlString queue)
        {
            return new MessageQueue(queue.ToString(), QueueAccessMode.Send);
        }

        /// <summary>
        /// Peeks message from queue
        /// </summary>
        /// <param name="queue">Queue path</param>
        /// <param name="msg">Message</param>
        public static void Peek(SqlString queue, out SqlString msg)
        {
            Message queueMsg = null;
            using (MessageQueue msgQueue = new MessageQueue(queue.ToString(), QueueAccessMode.Peek))
            {
                msgQueue.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });
                try
                {
                    queueMsg = msgQueue.Peek(TimeSpan.FromMilliseconds(10));
                }
                catch (MessageQueueException)
                {
                    msg = new SqlString();
                    return;
                }
                msg = new SqlString(queueMsg.Body.ToString());
            }
        }

        /// <summary>
        /// Receives message from queue
        /// </summary>
        /// <param name="queue">Queue path</param>
        /// <param name="msg">Message</param>
        public static void Receive(SqlString queue, out SqlString msg)
        {
            Message queueMsg = null;
            using (MessageQueue msgQueue = new MessageQueue(queue.ToString(), QueueAccessMode.Receive))
            {
                msgQueue.Formatter = new XmlMessageFormatter(new Type[] { typeof(string) });
                try
                {
                    queueMsg = msgQueue.Receive(TimeSpan.FromMilliseconds(10));
                }
                catch (MessageQueueException)
                {
                    msg = new SqlString();
                    return;
                }
                msg = new SqlString(queueMsg.Body.ToString());
            }
        }
    }
}
