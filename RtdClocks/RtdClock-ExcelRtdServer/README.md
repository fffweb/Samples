1, need to defin server
2, get last
This project has the following NuGet package installed:
* ExcelDna.AddIn

The add-in defines:
* an RTD server based on the ExcelRtdServer base class, 
* a UDF that called `dnaRtdClock`.

why not use IExcelAddIn
namespace AsyncFunctions
{
    public class AsyncTestAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelIntegration.RegisterUnhandledExceptionHandler(ex => "!!! EXCEPTION: " + ex.ToString());
        }

        public void AutoClose()
        {
        }
    }
