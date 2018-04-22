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
        //TODO: add dynamic type ,maybe decemail,or int
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
        public Dictionary<string, string> _typeMapFromExcelToCtp = new Dictionary<string, string>();
        // Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        // (normally that would be dangeours for an RTD server, but ExcelRtdServer is thrad-safe)
        Timer _timer;
        //List<QuoteTopic> _topics;

        Dictionary<string, List<QuoteTopic>> _symbolTopics;

        private void loadStringMap()
        {
            String[] excel_types = new String[] { "Symbol", "Open" };
            string[] ctp_types = new String[] { "Symbol", "Open" };
            foreach (var item in excel_types)
            {

            }
        }
        protected override bool ServerStart()
        {
            //_timer = new Timer(timer_tick, null, 0, 1000);
            //_topics = new List<QuoteTopic>();
            _symbolTopics = new Dictionary<string, List<QuoteTopic>>();
            getCtpticks();
            return true;
        }

        protected override void ServerTerminate()
        {
            if (_timer != null)
            {
                _timer.Dispose();
            }
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            QuoteTopic topic1 = (QuoteTopic)topic;
            //_topics.Add(topic1);
            //topic1.UpdateValue(0);

            List<QuoteTopic> topicsForSymbol;
            if (_symbolTopics.TryGetValue(topic1._symbol, out topicsForSymbol))
            {
                if (topicsForSymbol.Find(x => x.TopicId == topic1.TopicId) == null)
                {
                    topicsForSymbol.Add(topic1);
                }
            }
            else
            {
                topicsForSymbol = new List<QuoteTopic>();
                topicsForSymbol.Add(topic1);
                _symbolTopics.Add(topic1._symbol, topicsForSymbol);

            }
            return 0;
        }

        protected override void DisconnectData(Topic topic)
        {
            QuoteTopic topic1 = (QuoteTopic)topic;
            //_topics.Remove(topic1);

            List<QuoteTopic> topicsForSymbol;
            if (_symbolTopics.TryGetValue(topic1._symbol, out topicsForSymbol))
            {
                if (topicsForSymbol.Find(x => x.TopicId == topic1.TopicId) == null)
                {
                    topicsForSymbol.Remove(topic1);
                }
            }
        }

        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            return new QuoteTopic(this, topicId) { _symbol = topicInfo[0], _type = topicInfo[1] };
        }

        void timer_tick(object _unused_state_)
        {
            string now = DateTime.Now.ToString("HH:mm:ss");
            //foreach (var topic in _topics)
            //{
            //    var value = now + " TopicId:" + topic.TopicId ;// topic.Prefix + ":" + DateTime.Now.ToString("HH:mm:ss.fff") + ";" + _random.NextDouble().ToString("F5");
            //    topic.UpdateValue(value);
            //}

        }

        public static bool IsNumericType(Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }
        private async void getCtpticks()
        {
            var hubConnection = new HubConnection("http://localhost:52842");// ("http://www.wmleo.cc/");
            //var hubConnection = new HubConnection("http://www.wmleo.cc/");
            IHubProxy stockTickerHubProxy = hubConnection.CreateHubProxy("ctpQuantTicker");
            stockTickerHubProxy.On<Future>("UpdateFuturePrice", future =>
            {
                Console.WriteLine("Future update for {0} new price {1}", future.Symbol, future.LastPrice);
                //foreach (var topic in _topics)
                //{
                //    if (topic._symbol == stock.Symbol)//&& topic._type=="Price" && stock.Price != (double)topic.Value)
                //    {
                //        topic.UpdateValue(stock.LastPrice);
                //    }

                //    //PropertyInfo[] properties = typeof(Record).GetProperties();
                //    //foreach (PropertyInfo property in properties)
                //    //{
                //    //    property.SetValue(record, value);
                //    //}
                //}

                List<QuoteTopic> topicsForSymbol = _symbolTopics[future.Symbol];
                foreach (var topic in topicsForSymbol)
                {
                    double num;
                    //if (!double.TryParse((String)topic.Value, out num))
                    //{
                    //    //throw new InvalidOperationException("Value is not a number.");
                    //    Console.WriteLine("test");
                    //}

                    if (!(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.LastPrice))
                    //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                    {

                        topic.UpdateValue(future.LastPrice);
                    }
                    //decimal? tmpvalue = ConvertTo<Decimal>(topic.Value);
                    //if (decimal.TryParse(topic.Value, out tmpvalue))
                    //    result = tmpvalue;
                    //topic.UpdateValue(stock.DayOpen);
                    //if (topic._type == "Open" && (deciml)topic.Value != stock.OpenPrice)
                    //{
                    //    topic.UpdateValue(stock.OpenPrice);
                    //}
                    //if (topic._type == "High" && (double)topic.Value != stock.DayHigh)
                    //{
                    //    topic.UpdateValue(stock.DayHigh);
                    //}
                    //if (topic._type == "Low" && (double)topic.Value != stock.DayLow)
                    //{
                    //    topic.UpdateValue(stock.DayLow);
                    //}
                    //if (topic._type == "Last" && (decimal)topic.Value != stock.LastPrice)
                    //{
                    //    topic.UpdateValue(stock.Price);
                    //}
                    //if (topic._type == "Open" && (decimal)topic.Value != stock.OpenPrice)
                    //{
                    //    topic.UpdateValue(stock.OpenPrice);
                    //}
                }

                //foreach (var topic in _topics[stock.Symbol])
                //{
                //    if(topic.value != src.GetType().GetProperty(propName).GetValue(src, null))
                //    {
                //        topic.UpdateValue(stock.Price);
                //    }
                //}

            });
            await hubConnection.Start();
        }

        public static T? ConvertTo<T>(object x) where T : struct
        {
            return x == null ? null : (T?)Convert.ChangeType(x, typeof(T));
        }
    }
}
