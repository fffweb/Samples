using System;

namespace TradeOdata.Tickers
{
    public class Future
    {
        private decimal _price;

        public string Symbol { get; set; }

        public decimal AskPrice1 { get; set; }

        public decimal AveragePrice { get; set; }
        public int AskVolume1 { get; set; }

        public decimal BidPrice1 { get; set; }

        public int BidVolume1 { get; set; }

        public decimal ClosePrice { get; set; }

        //public decimal HighestPrice { get; set; }

        public string InstrumentID { get; set; }

        //public decimal LastPrice { get; set; }

        public decimal LowerLimitPrice { get; set; }

        //public decimal LowestPrice { get; set; }

        public double OpenInterest { get; set; }

        //public decimal OpenPrice { get; private set; }

        public decimal PreClosePrice { get; set; }

        public decimal PreDelta { get; set; }

        public double PreOpenInterest { get; set; }

        public decimal PreSettlementPrice { get; set; }

        public decimal SettlementPrice { get; set; }

        public string TradingDay { get; set; }

        public decimal Turnover { get; set; }

        public int UpdateMillisec { get; set; }

        public string UpdateTime { get; set; }

        public decimal UpperLimitPrice { get; set; }

        public int Volume { get; set; }

        public decimal LowestPrice { get; set; }

        public decimal HighestPrice { get; set; }

        public decimal LastChange { get; private set; }

        public decimal OpenPrice { get; set; }

        public decimal Change
        {
            get
            {
                return LastPrice - OpenPrice;
            }
        }

        public double PercentChange
        {
            get
            {
                return LastPrice == 0 ? 0 : (double)Math.Round(Change / LastPrice, 4);
            }
        }

        public decimal LastPrice
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

                if (OpenPrice == 0)
                {
                    OpenPrice = _price;
                }
                if (_price < LowestPrice || LowestPrice == 0)
                {
                    LowestPrice = _price;
                }
                if (_price > HighestPrice)
                {
                    HighestPrice = _price;
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
