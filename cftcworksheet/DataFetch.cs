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
        private readonly string webAddressFormat = @"http://www.cftc.gov/files/dea/cotarchives/{0}/futures/other_sf{1}.htm";//0: year
                                                                                                                            //1: date mmddyy eg. 072517 for July 25, 2017        
        public int GetMonthIndex(string month)
        {
            int ret = 0;
            switch (month)
            {
                case "January":
                    ret = 1;
                    break;
                case "February":
                    ret = 2;
                    break;
                case "March":
                    ret = 3;
                    break;
                case "April":
                    ret = 4;
                    break;
                case "May":
                    ret = 5;
                    break;
                case "June":
                    ret = 6;
                    break;
                case "July":
                    ret = 7;
                    break;
                case "August":
                    ret = 8;
                    break;
                case "September":
                    ret = 9;
                    break;
                case "October":
                    ret = 10;
                    break;
                case "November":
                    ret = 11;
                    break;
                case "December":
                    ret = 12;
                    break;
                default:
                    break;

            }
            return ret;
        }
        public bool GetGoldCommodity(out List<int> retLst, out DateTime updateDate, string soursePage)
        {
            updateDate = new DateTime();
            retLst = new List<int>(12);
            int _openInterestCol = 88;
            int _otherInfoCol = 90;
            int _fileFirstCol = 17;
            int _monIndex = 11;
            int _dateIndex = 12;
            int _yearIndex = 13;
            string[] pageLines = soursePage.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

            string dateInfoCol = pageLines[_fileFirstCol];
            var tempCol = System.Text.RegularExpressions.Regex.Split(dateInfoCol, @"\s{1,}");
            string[] datacols = { tempCol[_monIndex], tempCol[_dateIndex].Replace(",", ""), tempCol[_yearIndex] };

            int mon = GetMonthIndex(datacols[0]);
            int date = int.Parse(datacols[1]);
            int year = int.Parse(datacols[2]);

            updateDate = new DateTime(year, mon, date);

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

        public bool FetchData(out List<int> retData, out DateTime updateDate)
        {
            retData = new List<int>();
            bool isSuccess = false;
            WebClient MyWebClient = new WebClient();

            MyWebClient.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于对向Internet资源的请求进行身份验证的网络凭据。
            string webAddress = @"http://www.cftc.gov/dea/futures/other_sf.htm";
            Byte[] pageData = MyWebClient.DownloadData(webAddress);//从指定网站下载数据

            //string pageHtml = Encoding.Default.GetString(pageData);  //如果获取网站页面采用的是GB2312，则使用这句             

            string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句
            isSuccess = GetGoldCommodity(out retData, out updateDate, pageHtml);

            return isSuccess;

        }

        public bool FetchHistoricData(out List<int> refData, DateTime date)
        {
            refData = new List<int>();
            return false;
        }
    }


}
