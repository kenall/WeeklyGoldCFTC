using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CFTCWorkSheet
{
    class DataFetch
    {
        public bool GetGoldCommodity(out List<int> retLst, string soursePage)
        {
            retLst = new List<int>(12);
            int _openInterestCol = 88;
            int _otherInfoCol = 90;

            string[] pageLines = soursePage.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            string strOpenInterest = pageLines[_openInterestCol];
            var openIntLines = System.Text.RegularExpressions.Regex.Split(strOpenInterest, @"\s{2,}");
            strOpenInterest = openIntLines[2];
            strOpenInterest = strOpenInterest.Replace(",", "");
            int _openInterest;
            bool ret = Int32.TryParse(strOpenInterest, out _openInterest);
            if (!ret)
            {
                return false;
            }
            retLst.Add(_openInterest);

            string otherInfo = pageLines[_otherInfoCol];
            var dataList = System.Text.RegularExpressions.Regex.Split(otherInfo, @"\s{2,}").ToList();
            dataList.RemoveAt(0);
            for (int index = 0; index < dataList.Count; index++)
            {
                dataList[index] = dataList[index].Replace(",", "");
                dataList[index] = dataList[index].Replace(":", "");
                int tmpValue;
                ret = Int32.TryParse(dataList[index], out tmpValue);
                if (!ret)
                    return false;
                retLst.Add(tmpValue);
            }
            return true;
        }

        public bool FetchData(out List<int> retData)
        {
            retData = new List<int>();
            bool isSuccess = false;
            WebClient MyWebClient = new WebClient();

            MyWebClient.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于对向Internet资源的请求进行身份验证的网络凭据。
            string webAddress = @"http://www.cftc.gov/dea/futures/other_sf.htm";
            Byte[] pageData = MyWebClient.DownloadData(webAddress);//从指定网站下载数据

            //string pageHtml = Encoding.Default.GetString(pageData);  //如果获取网站页面采用的是GB2312，则使用这句             

            string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句
            isSuccess = GetGoldCommodity(out retData, pageHtml);

            return true;

        }

        public bool FetchHistoricData(out List<int> refData, DateTime date)
        {
            refData = new List<int>();
            return false;
        }
    }


}
