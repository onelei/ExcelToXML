using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
namespace ExcelToXML
{
    public partial class ExcelToXML
    {
        private void ExcelToXML_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            OnClickExcelToXML();  
        }

        private void About_Click(object sender, RibbonControlEventArgs e)
        {
            ShowAbout();
        }

        /// <summary>
        /// Excel to XML.
        /// </summary>
        void OnClickExcelToXML()
        {
            string fileName = ThisAddIn.Instance.workBook.Name;
            fileName = fileName.Replace(".xlsx", "");

            string text = "";
            int totalRows = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Rows.Count;
            int totalColumns = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Columns.Count;
            text += "<"+fileName+">" + "<elements>";
            text += "<oneItem>";
            int startRow = 1;
            int startColumn = 1;
            for (int i = startRow + 1; i != totalRows + 1; ++i)
            {
                for (int j = startColumn; j != totalColumns + 1; ++j)
                {
                    text += "<" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(startRow, j).Value + ">";
                    // In the last row and last column, modify text style;
                    if (i == totalRows && j == totalColumns)
                    {
                        text += "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "";
                        text += "<" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(startRow, j).Value + ">";
                    }
                    else
                    {
                        // Every Row Start, write "},";
                        if (j == totalColumns)
                        {
                            text += "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "";
                            text += "</" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(startRow, j).Value + ">";
                            text += "</oneItem><oneItem>";
                           // " },{";
                        }
                        else
                        {
                            text += "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "";
                            text += "</" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(startRow, j).Value + ">";
                        }
                    }

                }
            }
            text+="</oneItem>";
            text += "</elements>" +"</"+ fileName+">";

            // Write to file;
            string path = ThisAddIn.Instance.workBook.Path + "\\";

            
            fileName += ".xml";
            WriteToFile(path, fileName, text);
            ShowMessage(totalRows, totalColumns);
         }

        /*
       * Write to file
       * @param path
       * @param name
       * @param text
       * return
       */
        public void WriteToFile(string path, string name, string text)
        {
            lock (this)
            {
                // File stream information;
                StreamWriter sw;
                FileInfo t = new FileInfo(path + name);
                // Create File Stream;
                sw = t.CreateText();
                // Write text by line style;
                sw.WriteLine(text);
                // Close stream;
                sw.Close();
                // Destrory stream;
                sw.Dispose();
            }

        }

        /*
         * Show message.
         */
        private void ShowMessage(int totalRows, int totalColumns)
        {
            MessageBox.Show(
                "\n"
                + "  Save Successful!\n\n"
                + "  Total: " + totalRows + " Rows"
                + ", " + totalColumns + " Columns");
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                "\n"
                + "  Excel to XML\n"
                + "  Created by OneLei.\n"
                + "  Email: ahleiwolong@163.com \n"
                + "  Copyright (c) 2015 Year. All rights reserved.\n\n");
        }
    }
}
