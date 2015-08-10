using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace ExcelToXML
{
    public partial class ExcelToXML
    {
        private int KeyRow = 2;
        private int KeyColumn = 1;
        private int ValueRow = 4;
        private int ValueColumn = 1;
  
        private void ExcelToXML_Load(object sender, RibbonUIEventArgs e)
        {
            Key_R.Text = "" + KeyRow;
            Key_C.Text = "" + KeyColumn;
            Value_R.Text = "" + ValueRow;
            Value_C.Text = "" + ValueColumn;
        }

        private void Button_XML_Click(object sender, RibbonControlEventArgs e)
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
            CheckInputValue();
            Stopwatch MyCodeExeTime = new Stopwatch();
            MyCodeExeTime.Start();

            string fileName = ThisAddIn.Instance.workBook.Name;
            fileName = fileName.Replace(".xlsx", "");

            int totalRows = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Rows.Count;
            int totalColumns = ThisAddIn.Instance.workBook.ActiveSheet.UsedRange.Columns.Count;
            totalRows = FixtotalRows(totalRows);
            totalColumns = FixtotalColumns(totalColumns);

            string text = "";
            text += "<"+fileName+">" + "<elements>";
            text += "<oneItem>";      
            for (int i = ValueRow; i != totalRows + 1; ++i)
            {
                for (int j = ValueColumn; j != totalColumns + 1; ++j)
                {
                    text += "<" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(KeyRow, j).Value + ">";
                    // In the last row and last column, modify text style;
                    if (i == totalRows && j == totalColumns)
                    {
                        text += "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "";
                        text += "<" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(KeyRow, j).Value + ">";
                    }
                    else
                    {
                        // Every Row Start, write "},";
                        if (j == totalColumns)
                        {
                            text += "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "";
                            text += "</" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(KeyRow, j).Value + ">";
                            text += "</oneItem><oneItem>";
                           // " },{";
                        }
                        else
                        {
                            text += "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, j).Value + "";
                            text += "</" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(KeyRow, j).Value + ">";
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
            MyCodeExeTime.Stop();
            string myTime = MyCodeExeTime.ElapsedMilliseconds.ToString();
            ShowMessage(totalRows, totalColumns, myTime);
         }

       /**
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
        private void ShowMessage(int totalRows, int totalColumns, string _ExeTime)
        {
            MessageBox.Show(
                "\n"
                + "  Excel to XML\n"
                + "  Save Successful!\n\n"
                + "  Time:  " + _ExeTime + "毫秒.\n"
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

        void CheckInputValue()
        {
            // Get the key;
            if (JugeNumberOrNot(Key_R.Text))
            {
                KeyRow = System.Int32.Parse(Key_R.Text);
            }
            if (JugeNumberOrNot(Key_C.Text))
            {
                KeyColumn = System.Int32.Parse(Key_C.Text);
            }

            // Get the value;
            if (JugeNumberOrNot(Value_R.Text))
            {
                ValueRow = System.Int32.Parse(Value_R.Text);
            }
            if (JugeNumberOrNot(Value_C.Text))
            {
                ValueColumn = System.Int32.Parse(Value_C.Text);
            }
        }

        private bool JugeNumberOrNot(string _text)
        {
            if (String.IsNullOrEmpty(_text))
            {
                MessageBox.Show("Input value is null or empty");
                return false;
            }
            else
            {
                try
                {
                    Int32.Parse(_text);
                }
                catch
                {
                    MessageBox.Show("Input value is not number!");
                    return false;
                }
            }
            return true;
        }

        private int FixtotalRows(int totalRows)
        {
            for (int i = 1; i < totalRows; i++)
            {
                string cellValue = "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(i, 1).Value;
                if (string.IsNullOrEmpty(cellValue))
                {
                    // 当前的为空,则totalRow到此为止,修复totalRow;
                    return i - 1;
                }
            }
            return totalRows;
        }

        private int FixtotalColumns(int totalColumns)
        {
            for (int i = 1; i < totalColumns; i++)
            {
                string cellValue = "" + ThisAddIn.Instance.workBook.ActiveSheet.Cells(1, i).Value;
                if (string.IsNullOrEmpty(cellValue))
                {
                    // 当前的为空,则totalRow到此为止,修复totalRow;
                    return i - 1;
                }
            }
            return totalColumns;
        }
         
    }
}
