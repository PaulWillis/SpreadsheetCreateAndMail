using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Collections.Specialized;

namespace Trending_Toward_Writeoff
{
    class Program
    {
        private static NameValueCollection appConfig = ConfigurationManager.AppSettings;

        static void Main(string[] args)
        {
            SpinAnimation.Start(50);

            LoadDataIntoExcel l = new LoadDataIntoExcel();
            string ReportName = appConfig["ReportName"]; 
            l.Run();

            SpinAnimation.Stop();
        }
    }
}
