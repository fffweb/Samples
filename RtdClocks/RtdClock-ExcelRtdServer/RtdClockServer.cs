using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration.Rtd;
using Microsoft.AspNet.SignalR.Client;
using TradeOdata.Tickers;
using System.Threading.Tasks;

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
        public Dictionary<string, string> _typeMapFromExcelToCtp = new Dictionary<string, string>();
        // Using a System.Threading.Time which invokes the callback on a ThreadPool thread 
        // (normally that would be dangeours for an RTD server, but ExcelRtdServer is thrad-safe)
        Timer _timer;
        List<QuoteTopic> _topics;
        HubConnection _hubConnection;
        QuoteTopic _stateTopic;
        string _urlServer;
        private bool _isServerClose=false;

        //Dictionary<string, List<QuoteTopic>> _symbolTopics;

        public RtdClockServer() : base()
        {

        }

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
            _topics = new List<QuoteTopic>();
            //_symbolTopics = new Dictionary<string, List<QuoteTopic>>();
            //getCtpticks();
            return true;
        }

        protected override void ServerTerminate()
        {
            if (_hubConnection != null)
            {
                _isServerClose = true;
                _hubConnection.Dispose();
            }
            if (_timer != null)
            {
                _timer.Dispose();
            }
        }

        protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
        {
            QuoteTopic topic1 = (QuoteTopic)topic;
            _topics.Add(topic1);
            if (topic1._symbol == "Server")
            {
                if (_hubConnection != null)
                {
                    _hubConnection.Dispose();
                }
                try
                {
                    _urlServer = topic1._type;
                    var task = connectToSignalServer(topic1._type);
                    task.Wait();
                }
                catch (Exception e)
                {

                    //throw;
                    return e.Message;
                }
                _stateTopic = topic1;
                return topic1._type;
            }
            //topic1.UpdateValue(0);
            //List<QuoteTopic> topicsForSymbol;
            //if (_symbolTopics.TryGetValue(topic1._symbol, out topicsForSymbol))
            //{
            //    if (topicsForSymbol.Find(x => x.TopicId == topic1.TopicId) == null)
            //    {
            //        topicsForSymbol.Add(topic1);
            //    }
            //}
            //else
            //{
            //    topicsForSymbol = new List<QuoteTopic>();
            //    topicsForSymbol.Add(topic1);
            //    _symbolTopics.Add(topic1._symbol, topicsForSymbol);

            //}
            return 0;
        }

        protected override void DisconnectData(Topic topic)
        {
            QuoteTopic topic1 = (QuoteTopic)topic;
            _topics.Remove(topic1);

            //List<QuoteTopic> topicsForSymbol;
            //if (_symbolTopics.TryGetValue(topic1._symbol, out topicsForSymbol))
            //{
            //    if (topicsForSymbol.Find(x => x.TopicId == topic1.TopicId) == null)
            //    {
            //        topicsForSymbol.Remove(topic1);
            //    }
            //}
        }
        private string getCTPSymbol(string symbol)
        {
            if (symbol == "Server")
                return symbol;
            string FirstTwoLetter = symbol.Substring(0, 2);
            if (FirstTwoLetter == "AP" ||
                    FirstTwoLetter == "CF" ||
                    FirstTwoLetter == "FG" ||
                    FirstTwoLetter == "MA" ||
                    FirstTwoLetter == "OI" ||
                    FirstTwoLetter == "RM" ||
                    FirstTwoLetter == "SF" ||
                    FirstTwoLetter == "SM" ||
                    FirstTwoLetter == "SR" ||
                    FirstTwoLetter == "TA" ||
                    FirstTwoLetter == "ZC")
            {
                if (symbol.Length == 6)
                {
                    //remove digit ,like FG1809=>FG809
                    return symbol.Remove(2, 1);
                }
                else
                {
                    //already ctp type symbol, no need to remove 
                    return symbol;
                }
            }

            if (FirstTwoLetter != "IC" && FirstTwoLetter != "IH" && FirstTwoLetter != "IF" && FirstTwoLetter != "TF" && FirstTwoLetter != "T1")
            {
                //all dce and shfe is lowcase.
                return symbol.ToLower();
            }


            return symbol;
        }
        protected override Topic CreateTopic(int topicId, IList<string> topicInfo)
        {
            return new QuoteTopic(this, topicId) { _symbol = getCTPSymbol(topicInfo[0]), _type = topicInfo[1] };
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
        public async Task<bool> connectToSignalServer(string url)
        {
            bool connected=false;
            try
            {


                //_hubConnection = new HubConnection("http://localhost:52842");// ("http://www.wmleo.cc/");
                //var hubConnection = new HubConnection("http://www.wmleo.cc/");
                _hubConnection = new HubConnection(url);
 

                IHubProxy stockTickerHubProxy = _hubConnection.CreateHubProxy("ctpQuantTicker");
                //TODO: use RX
                //stockTickerHubProxy.Observe()
                stockTickerHubProxy.On<Future>("UpdateFuturePrice", updateTopic);
                await _hubConnection.Start();
                if (_hubConnection.State == ConnectionState.Connected)
                {
                    connected = true;
                    _hubConnection.ConnectionSlow += () =>
                    {
                        Console.WriteLine("Connection problems.");
                        _stateTopic.UpdateValue(DateTime.Now + " ConnectionSlow|" + _stateTopic.Value);
                    };
                    _hubConnection.Closed += async () =>
                    {
                        Console.WriteLine("Connection closed.");
                        _stateTopic.UpdateValue(DateTime.Now + "Closed|" + _stateTopic.Value);

                        if (!_isServerClose) // A global variable being set in "Form_closing" event of Form, check if form not closed explicitly to prevent a possible deadlock.
                        {
                            // specify a retry duration
                            TimeSpan retryDuration = TimeSpan.FromSeconds(30);

                            while (DateTime.UtcNow < DateTime.UtcNow.Add(retryDuration))
                            {
                                bool connected1 = await connectToSignalServer(_urlServer);
                                if (connected1)
                                    return;
                            }
                            Console.WriteLine("Connection closed");
                        }

                    };                    
                }

                return connected;
            }
            catch (Exception e)
            {
                return false;
                //throw;
            }
        }

        //private async void Connection_Closed()
        //{
        //    if (!IsFormClosed) // A global variable being set in "Form_closing" event of Form, check if form not closed explicitly to prevent a possible deadlock.
        //    {
        //        // specify a retry duration
        //        TimeSpan retryDuration = TimeSpan.FromSeconds(30);

        //        while (DateTime.UtcNow < DateTime.UtcNow.Add(retryDuration))
        //        {
        //            bool connected = await ConnectToSignalRServer();
        //            if (connected)
        //                return;
        //        }
        //        Console.WriteLine("Connection closed")
        //    }
        //}
        private void updateTopic(Future future)
        {
            try
            {


                Console.WriteLine("Future update for {0} new price {1}", future.Symbol, future.LastPrice);
                foreach (var topic in _topics)
                {
                    if (topic._symbol == future.Symbol)//&& topic._type=="Price" && stock.Price != (double)topic.Value)
                    {
                        //topic.UpdateValue(future.LastPrice);

                        if (topic._type == "Last" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.LastPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.LastPrice);
                        }
                        if (topic._type == "Open" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.OpenPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.OpenPrice);
                        }
                        if (topic._type == "High" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.HighestPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.HighestPrice);
                        }
                        if (topic._type == "Low" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.LowestPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.LowestPrice);
                        }

                        if (topic._type == "SettlementPrice" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.SettlementPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.SettlementPrice);
                        }
                        if (topic._type == "PreSettlementPrice" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.PreSettlementPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.PreSettlementPrice);
                        }
                        if (topic._type == "UpdateTime" &&  Convert.ToString(topic.Value)!= future.UpdateTime)
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.UpdateTime);
                        }
                        if (topic._type == "UpperLimit" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.UpperLimitPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.UpperLimitPrice);
                        }
                        if (topic._type == "LowerLimit" && !(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value) == future.LowerLimitPrice))
                        //if (!(Convert.ToDecimal(topic.Value) == future.LastPrice))
                        {
                            topic.UpdateValue(future.LowerLimitPrice);
                        }
                    }

                    //PropertyInfo[] properties = typeof(Record).GetProperties();
                    //foreach (PropertyInfo property in properties)
                    //{
                    //    property.SetValue(record, value);
                    //}
                }

                //List<QuoteTopic> topicsForSymbol = _symbolTopics[stock.Symbol];
                //foreach (var topic in topicsForSymbol)
                //{
                //    double num;
                //    //if (!double.TryParse((String)topic.Value, out num))
                //    //{
                //    //    //throw new InvalidOperationException("Value is not a number.");
                //    //    Console.WriteLine("test");
                //    //}

                //   //if (!(IsNumericType(topic.Value.GetType()) && Convert.ToDecimal(topic.Value)==stock.LastPrice) )
                //    {

                //        topic.UpdateValue(stock.LastPrice);
                //    }
                //    //decimal? tmpvalue = ConvertTo<Decimal>(topic.Value);
                //    //if (decimal.TryParse(topic.Value, out tmpvalue))
                //    //    result = tmpvalue;
                //    //topic.UpdateValue(stock.DayOpen);
                //    //if (topic._type == "Open" && (deciml)topic.Value != stock.OpenPrice)
                //    //{
                //    //    topic.UpdateValue(stock.OpenPrice);
                //    //}
                //    //if (topic._type == "High" && (double)topic.Value != stock.DayHigh)
                //    //{
                //    //    topic.UpdateValue(stock.DayHigh);
                //    //}
                //    //if (topic._type == "Low" && (double)topic.Value != stock.DayLow)
                //    //{
                //    //    topic.UpdateValue(stock.DayLow);
                //    //}
                //    //if (topic._type == "Last" && (decimal)topic.Value != stock.LastPrice)
                //    //{
                //    //    topic.UpdateValue(stock.Price);
                //    //}
                //    //if (topic._type == "Open" && (decimal)topic.Value != stock.OpenPrice)
                //    //{
                //    //    topic.UpdateValue(stock.OpenPrice);
                //    //}
                //}
                //foreach (var topic in _topics[stock.Symbol])
                //{
                //    if(topic.value != src.GetType().GetProperty(propName).GetValue(src, null))
                //    {
                //        topic.UpdateValue(stock.Price);
                //    }
                //}

            }
            catch (Exception ex)
            {

                Console.WriteLine("Error invoking GetAllStocks: {0}", ex.Message);
            }
        }

        public static T? ConvertTo<T>(object x) where T : struct
        {
            return x == null ? null : (T?)Convert.ChangeType(x, typeof(T));
        }
    }
}
