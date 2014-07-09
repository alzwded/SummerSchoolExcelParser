/*
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
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SummerSchoolExcelParserDeux
{
    class OutputProducer
    {
        private String path_;
        private int numRows_;
        private int numCols_;

        public OutputProducer(String path, int numRows, int numCols)
        {
            path_ = path;
            numRows_ = numRows;
            numCols_ = numCols;
        }

        private int Convert(String cell)
        {
            if (cell == "med") return 1;
            if (cell == "high") return 3;
            return 0;
        }

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
                            int oldval = squashed[s.name][i];
                            int t = Convert(s.data[columns[i]]);
                            squashed[s.name][i] = oldval + t;
                        }
                    }
                }
            }

            return squashed;
        }

        public void Perform(List<List<Student>> data, String[] columns)
        {
            Dictionary<String, List<int>> odata = Squash(data, columns);

            var excelApp = new Excel.Application();
            
            Excel.Workbook wb = excelApp.Workbooks.Add();

            Excel.Worksheet ws = wb.Sheets.Add();

            try
            {
                ws.Activate();

                ws.Cells[2, 1] = "Student";
                for (int i = 0; i < columns.Length; ++i)
                {
                    ws.Cells[2, 2 + i] = columns[i];
                }

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

                wb.SaveAs(path_);
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
