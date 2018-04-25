using ExcelDna.Integration;

namespace RtdClock_ExcelRtdServer
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object dnaRtdClock_ExcelRtdServer()
        {
            // Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
            // Note that the topic information needs at least one string - it's not used in this sample
            return XlCall.RTD(RtdClockServer.ServerProgId, null, "");
        }

        [ExcelFunction(Description = "Provides a ticking clock id")]
        public static object dnaRtdClock_ExcelRtdServer1(string symbol, string type)
        {
            // Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
            // Note that the topic information needs at least one string - it's not used in this sample
            return XlCall.RTD(RtdClockServer.ServerProgId, null, new string[] { symbol, type });
        }

        public static object Sleep(string ms)
        {
            object result = ExcelAsyncUtil.Run("Sleep", ms, delegate
            {
                //Debug.Print("{1:HH:mm:ss.fff} Starting to sleep for {0} ms", ms, DateTime.Now);
                System.Threading.Thread.Sleep(int.Parse(ms));
                //Debug.Print("{1:HH:mm:ss.fff} Completed sleeping for {0} ms", ms, DateTime.Now);
                return "Woke Up at " + System.DateTime.Now.ToString("1:HH:mm:ss.fff");
            });
            if (Equals(result, ExcelError.ExcelErrorNA))
            {
                return "!!! Sleeping...";
            }
            return result;
        }
    }
}
