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
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SummerSchoolExcelParserDeux
{
    struct Student
    {
        public String name;
        public Dictionary<String, String> data;
    }
    class ExcelParser
    {
        private String path_;
        private int numRows_;
        private int numCols_;

        private String[] lastCols_;

        public String[] LastColumns
        {
            get { return lastCols_; }
        }

        public ExcelParser(String path, int numRows, int numCols)
        {
            path_ = path;
            numRows_ = numRows;
            numCols_ = numCols;
        }

        private List<Student> DoSheet(Excel.Worksheet ws)
        {
            List<Student> ret = new List<Student>();

            //var numRows = numRows_;//ws.Rows.Count - 2; don't work
            //var numCols = numCols_; ;// ws.Columns.Count - 1; don't work
            Excel.Range range = ws.UsedRange;
            int numRows = range.Rows.Count - 2;
            int numCols = range.Columns.Count - 1;
            if (numRows < 0 || numCols < 0) throw new Exception("Invalid format?");

            List<String> colNames = new List<String>();
            for (int i = 0; i < numCols; ++i)
            {
                const int offshot = 2;
                String name = ws.Cells[2, offshot + i].Text;

                if (name == "")
                {
                    numCols = i;
                    break;
                }
                colNames.Add(name);
            }

            String[] colNamesA = colNames.ToArray<String>();
            lastCols_ = colNamesA;

            for (int i = 0; i < numRows; ++i)
            {
                const int offshot = 3;

                StringBuilder sb = new StringBuilder("A");
                sb.Append(offshot + i);

                String name = ws.Cells[offshot + i, "A"].Text;
                if (name == "")
                {
                    numRows = i;
                    break;
                }

                Student theGuy = new Student();
                theGuy.name = name;
                theGuy.data = new Dictionary<String, String>();

                for (int j = 0; j < numCols; ++j)
                {
                    const int colOff = 2;

                    String val = ws.Cells[offshot + i, colOff + j].Text;

                    theGuy.data.Add(colNamesA[j], val);
                }

                ret.Add(theGuy);
            }

            return ret;
        }

        public List<List<Student>> Get()
        {
            List<List<Student>> ret = new List<List<Student>>();

            var excelApp = new Excel.Application();
            Excel.Workbook wb = excelApp.Workbooks.Open(path_, ReadOnly: true);
            //Excel.Worksheet ws = wb.Sheets[1];
            //ws.Activate();

            foreach (Excel.Worksheet ws in wb.Sheets)
            {
                String name = ws.Name;
                if (name == "TEMPLATE")
                {
                    continue;
                }
                ws.Activate();
                ret.Add(DoSheet(ws));

                Marshal.FinalReleaseComObject(ws);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();                   
            GC.WaitForPendingFinalizers();

            wb.Close(SaveChanges: false);
            Marshal.FinalReleaseComObject(wb); 

            excelApp.Quit();
            Marshal.FinalReleaseComObject(excelApp); 

            return ret;
        }
    }
}
