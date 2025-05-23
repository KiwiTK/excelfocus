using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private Excel.Worksheet currentSheet;
        private Excel.Range lastRow;
        private Excel.Range lastColumn;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SheetSelectionChange += Application_SheetSelectionChange;
        }

        private void Application_SheetSelectionChange(object sh, Excel.Range target)
        {
            // 清除上一次的高亮
            if (lastRow != null)
                lastRow.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            if (lastColumn != null)
                lastColumn.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;

            currentSheet = (Excel.Worksheet)target.Worksheet;
            int row = target.Row;
            int col = target.Column;

            // 高亮当前行
            lastRow = currentSheet.Rows[row];
            lastRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

            // 高亮当前列
            lastColumn = currentSheet.Columns[col];
            lastColumn.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCyan);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.SheetSelectionChange -= Application_SheetSelectionChange;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
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
