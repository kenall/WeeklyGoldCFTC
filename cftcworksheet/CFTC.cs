using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace CFTCWorkSheet
{
    public partial class CFTC
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.Update_Clicked);
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

        private void Update_Clicked(object sender, EventArgs e)
        {
            DataFetch df = new DataFetch();
            List<int> lstData;
            DateTime publishedDate;
            df.FetchData(out lstData, out publishedDate);
            
            ExcelOperator eo = new ExcelOperator();
            if (eo.IsNeedUpdate(publishedDate))
            {
                eo.UpdateData(ref lstData);
                MessageBox.Show("数据更新完毕");
            }
            else
            {
                MessageBox.Show("当前没有最新数据");
            }
        }
    }
}
