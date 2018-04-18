using System;

namespace TradeOdata.Tickers
{
    public class Future
    {
        private double _price;

        public string Symbol { get; set; }

        public decimal AskPrice1 { get; set; }

        public decimal AveragePrice { get; set; }
        public int AskVolume1 { get; set; }

        public decimal BidPrice1 { get; set; }

        public int BidVolume1 { get; set; }

        public decimal ClosePrice { get; set; }

        public decimal HighestPrice { get; set; }

        public string InstrumentID { get; set; }

        public decimal LastPrice { get; set; }

        public decimal LowerLimitPrice { get; set; }

        public decimal LowestPrice { get; set; }

        public int OpenInterest { get; set; }
        
        public decimal OpenPrice { get; private set; }

        public decimal PreClosePrice { get; set; }

        public decimal PreDelta { get; set; }

        public int PreOpenInterest { get; set; }

        public decimal PreSettlementPrice { get; set; }

        public decimal SettlementPrice { get; set; }

        public string TradingDay { get; set; }

        public int Turnover { get; set; }

        public int UpdateMillisec { get; set; }

        public string UpdateTime { get; set; }

        public decimal UpperLimitPrice { get; set; }

        public int Volume { get; set; }

        public double DayLow { get;  set; }
        
        public double DayHigh { get;  set; }

        public double LastChange { get; private set; }

        public double DayOpen { get;  set; }

        public double Change
        {
            get
            {
                return Price - DayOpen;
            }
        }

        public double PercentChange
        {
            get
            {
                return (double)Math.Round(Change / Price, 4);
            }
        }

        public Double Price
        {
            get
            {
                return _price;
            }
            set
            {
                if (_price == value)
                {
                    return;
                }

                LastChange = value - _price;
                _price = value;

                if (DayOpen == 0)
                {
                    DayOpen = _price;
                }
                if (_price < DayLow || DayLow == 0)
                {
                    DayLow = _price;
                }
                if (_price > DayHigh)
                {
                    DayHigh = _price;
                }
            }
        }
        //public decimal Change
        //{
        //    get
        //    {
        //        return Price - PreSettlementPrice;
        //    }
        //}

        //public double PercentChange
        //{
        //    get
        //    {
        //        return (double)Math.Round(Change / Price, 4);
        //    }
        //}

        //public decimal Price
        //{
        //    get
        //    {
        //        return _price;
        //    }
        //    set
        //    {
        //        if (_price == value)
        //        {
        //            return;
        //        }
                
        //        LastChange = value - _price;
        //        _price = value;
                
        //        if (OpenPrice == 0)
        //        {
        //            OpenPrice = _price;
        //        }
        //        if (_price < DayLow || DayLow == 0)
        //        {
        //            DayLow = _price;
        //        }
        //        if (_price > DayHigh)
        //        {
        //            DayHigh = _price;
        //        }
        //    }
        //}
    }
}
