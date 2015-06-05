using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelToXML
{
    public partial class ThisAddIn
    {
        public static ThisAddIn Instance;
        public Excel.Workbook workBook;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Instance = this;
            this.Application.WorkbookOpen += this.MyWorkBookOpen;
        }

        public void MyWorkBookOpen(Excel.Workbook Wb)
        {
            workBook = Wb;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
