using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration.Rtd;
using Microsoft.AspNet.SignalR.Client;
using TradeOdata.Tickers;

namespace RtdClock_ExcelRtdServer
{
    class QuoteTopic : ExcelRtdServer.Topic
    {
        public string _symbol;
        public string _type;

        public QuoteTopic(ExcelRtdServer server, int topicId) :
            base(server, topicId)
        {
        }
    }

    [ComVisible(true)]                   // Required since the default template puts [assembly:ComVisible(false)] in the AssemblyInfo.cs
    [ProgId(RtdClockServer.ServerProgId)]     //  If ProgId is not specified, change the XlCall.RTD call in the wrapper to use namespace + type name (the default ProgId)
    public class RtdClockServer : ExcelRtdServer
    {
        public const string ServerProgId = "RtdClock.ClockServer";

        // Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        // (normally that would be dangeours for an RTD server, but ExcelRtdServer is thrad-safe)
        Timer _timer;
        List<QuoteTopic> _topics;

        protected override bool ServerStart()
        {
            //_timer = new Timer(timer_tick, null, 0, 1000);
            _topics = new List<QuoteTopic>();
            getCtpticks();
            return true;
        }

        protected override void ServerTerminate()
        {
            _timer.Dispose();
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            QuoteTopic topic1 = (QuoteTopic)topic;
            _topics.Add(topic1);
            return DateTime.Now.ToString("HH:mm:ss") + " (ConnectData)**********";
        }

        protected override void DisconnectData(Topic topic)
        {
            QuoteTopic topic1 = (QuoteTopic)topic;
            _topics.Remove(topic1);
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            return new QuoteTopic(this, topicId) { _symbol = topicInfo[0],_type=topicInfo[1]};
        }

    void timer_tick(object _unused_state_)
        {
            string now = DateTime.Now.ToString("HH:mm:ss");
            foreach (var topic in _topics)
            {
                var value = now + " TopicId:" + topic.TopicId ;// topic.Prefix + ":" + DateTime.Now.ToString("HH:mm:ss.fff") + ";" + _random.NextDouble().ToString("F5");
                topic.UpdateValue(value);
            }
                
        }

        private async void getCtpticks()
        {
            var hubConnection = new HubConnection("http://localhost:52842/");// ("http://www.wmleo.cc/");
            IHubProxy stockTickerHubProxy = hubConnection.CreateHubProxy("ctpQuantTicker");
            stockTickerHubProxy.On<Future>("UpdateStockPrice", stock => {
                Console.WriteLine("Future update for {0} new price {1}", stock.Symbol, stock.Price);
                foreach (var topic in _topics)
                {
                    if (topic._symbol == stock.Symbol)
                    {
                        topic.UpdateValue(stock.Price);
                    }
                }
            });
            await hubConnection.Start();
        }
    }
}
