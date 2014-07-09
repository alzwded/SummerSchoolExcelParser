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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SummerSchoolExcelParserDeux
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            String path = "";
            SaveFileDialog sd = new SaveFileDialog();
            sd.DefaultExt = "xlsx";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                path = sd.FileName;
            }
            else
            {
                return;
            }

            String[] items = this.textBox1.Text.Split('\n').ToArray<String>();

            List<List<Student>> data = new List<List<Student>>();

            HashSet<String> cols = new HashSet<String>();

            foreach (String i in items)
            {
                try
                {
                    ExcelParser ep = new ExcelParser(i, Properties.Settings.Default.ROWS, Properties.Settings.Default.COLUMNS);
                    data.AddRange(ep.Get());
                    
                    foreach(String s in ep.LastColumns) cols.Add(s);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine(ex.ToString());
                }
            }

            Console.Write("Writing to {0} :: Columns: ", path);
            Console.WriteLine(String.Join(",", cols));

            OutputProducer op = new OutputProducer(path, Properties.Settings.Default.ROWS, Properties.Settings.Default.COLUMNS);
            op.Perform(data, cols.ToArray<String>());
        }
    }
}