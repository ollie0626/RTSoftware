using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Scope_Simple_tool.MyPage
{
    public sealed class MessageExt
    {
        private static readonly MessageExt instance = new MessageExt();

        private MessageExt()
        {
        }

        public static MessageExt Instance
        {
            get
            {
                return instance;
            }
        }

        /// <summary>
        /// 调用消息窗口的代理事件
        /// </summary>
        public Action<string, string> ShowDialog { get; set; }

        /// <summary>
        /// 调用消息确认窗口的代理事件
        /// </summary>
        public Action<string, string, Action> ShowYesNo { get; set; }
    }
}
