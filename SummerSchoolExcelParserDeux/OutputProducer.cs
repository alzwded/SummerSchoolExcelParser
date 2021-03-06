﻿/*
Copyright (c) 2014, Vlad Mesco
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace SummerSchoolExcelParserDeux
{
    /// <summary>
    /// Writes out the result spreadsheet
    /// </summary>
    class OutputProducer
    {
        private String path_;
        /// <summary>
        /// in memory representation of the settings
        /// </summary>
        private Dictionary<String, int> amounts_;

        /// <summary>
        /// </summary>
        /// <param name="path">the path to write to</param>
        public OutputProducer(String path)
        {
            path_ = path;
            amounts_ = new Dictionary<String, int>();

            // parse the settings and build the dictionary
            Regex r = new Regex("(?<name>[^:]*):(?<value>.*)");
            foreach (String s in Properties.Settings.Default.VALUES)
            {
                Match m = r.Match(s);
                int val = 0;
                if (!int.TryParse(m.Groups["value"].Value, out val)) continue;

                amounts_.Add(m.Groups["name"].Value, val);
            }
        }

        /// <summary>
        /// translate the stringy value to an integer value as specified in the application's settings
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private int Convert(String cell)
        {
            if (amounts_.Keys.Contains(cell)) return amounts_[cell];
            return 0;
        }

        /// <summary>
        /// flatten the 3d view to a 2d table by reducing the time series via summation
        /// </summary>
        /// <param name="data">raw data</param>
        /// <param name="columns">column names; mostly to keep the same order in the results</param>
        /// <returns></returns>
        private Dictionary<String, List<int>> Squash(List<List<Student>> data, String[] columns)
        {
            Dictionary<String, List<int>> squashed = new Dictionary<String, List<int>>();

            foreach (List<Student> sheet in data)
            {
                foreach (Student s in sheet)
                {
                    if (!squashed.Keys.Contains(s.name))
                    {
                        List<int> numbers = new List<int>();
                        foreach(String t in columns) numbers.Add(0);
                        squashed.Add(s.name, numbers);
                    }
                    
                    for(int i = 0; i < columns.Length; ++i) {
                        if (s.data.Keys.Contains(columns[i]))
                        {
                            squashed[s.name][i] = squashed[s.name][i] + Convert(s.data[columns[i]]);
                        }
                    }
                }
            }

            return squashed;
        }

        private Excel.XlFileFormat FormatFromExtension(String path)
        {
            Regex r = new Regex("^.*\\.(?<ext>[^.]*)$");
            var m = r.Match(path);
            if (m.Success)
            {
                switch (m.Groups["ext"].Value)
                {
                    case "xlsx":
                        return Excel.XlFileFormat.xlOpenXMLWorkbook;
                    case "xls":
                        return Excel.XlFileFormat.xlExcel8;
                    case "csv":
                        return Excel.XlFileFormat.xlCSV;
                    default:
                        return Excel.XlFileFormat.xlWorkbookDefault;
                }
            }
            else
            {
                return Excel.XlFileFormat.xlWorkbookDefault;
            }
        }

        /// <summary>
        /// parse the raw data and write out the results using columns as the table header
        /// </summary>
        /// <param name="data">the raw data</param>
        /// <param name="columns">the header columns</param>
        public void Perform(List<List<Student>> data, String[] columns)
        {
            Dictionary<String, List<int>> odata = Squash(data, columns);

            var excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Add();
            Excel.Worksheet ws = wb.Sheets.Add();

            try
            {
                ws.Activate();
                ws.Name = "Results";

                // build the header
                ws.Cells[2, 1] = "Student";
                for (int i = 0; i < columns.Length; ++i)
                {
                    ws.Cells[2, 2 + i] = columns[i];
                    Excel.Range r = ws.Cells[2, 2 + i];
                    r.WrapText = true;
                    r.Font.Bold = true;
                    r.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    r.ColumnWidth = 12;

                    Excel.Borders brs = r.Borders;
                    brs[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                    Marshal.FinalReleaseComObject(brs);

                    Marshal.FinalReleaseComObject(r);
                }

                Excel.Range addFilterRng = ws.Range["A2", ws.Cells[2, 1 + columns.Length]];
                addFilterRng.AutoFilter(1, Operator: Excel.XlAutoFilterOperator.xlAnd, VisibleDropDown: true);
                Marshal.FinalReleaseComObject(addFilterRng);

                // add the data
                int idx = 0;
                foreach (KeyValuePair<String, List<int>> kv in odata)
                {
                    Console.WriteLine("{0}: {1}", kv.Key, String.Join(",", kv.Value.ToArray()));
                    ws.Cells[3 + idx, 1] = kv.Key;
                    for (int i = 0; i < columns.Length; ++i)
                    {
                        ws.Cells[3 + idx, 2 + i] = kv.Value[i];
                    }
                    ++idx;
                }

                // remove "Sheet [123]"
                excelApp.DisplayAlerts = false;
                foreach (Excel.Worksheet w2 in wb.Worksheets)
                {
                    if (w2.Equals(ws)) continue;
                    w2.Delete();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Marshal.FinalReleaseComObject(w2);
                }
                wb.SaveAs(path_, FormatFromExtension(path_));
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(ws);

                wb.Close(false);
                Marshal.FinalReleaseComObject(wb);

                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
            }
        }
    }
}
