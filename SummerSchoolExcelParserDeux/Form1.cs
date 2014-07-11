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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Threads = System.Threading;

namespace SummerSchoolExcelParserDeux
{
    internal class Wurkr
    {
        public event System.EventHandler Done;
        public delegate void Do();

        private Do do_;

        public Wurkr(Do f)
        {
            do_ = f;
        }

        private void PDo()
        {
            do_();
            Done.Invoke(this, new EventArgs());
        }

        public void Doa()
        {
            var self = this;
            Threads.ThreadStart starter = new Threads.ThreadStart(self.PDo);
            
            Threads.Thread t = new Threads.Thread(starter) { IsBackground = true };
            t.Start();
        }
    }

    public partial class Form1 : Form
    {
        public const String SAVE_BUTTON_NORMAL = "Save";
        public const String SAVE_BUTTON_WORKING = "Working...";

        public Form1()
        {
            InitializeComponent();
            this.button1.Text = SAVE_BUTTON_NORMAL;
        }

        private void Process(String path)
        {
            List<List<Student>> data = new List<List<Student>>();
            HashSet<String> cols = new HashSet<String>();

            foreach (String i in from s in this.textBox1.Lines where s.Length > 0 select s)
            {
                try
                {
                    ExcelParser ep = new ExcelParser(i);
                    data.AddRange(ep.Get());

                    foreach (String s in ep.LastColumns) cols.Add(s);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Some random error happened. Check console window for details");
                    Console.Error.WriteLine(ex.ToString());
                }
            }

            Console.Write("Writing to {0} :: Columns: ", path);
            Console.WriteLine(String.Join(",", cols));

            try
            {
                OutputProducer op = new OutputProducer(path);
                op.Perform(data, cols.ToArray<String>());
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Some random error happened. Check console window for details");
                Console.Error.WriteLine(ex.ToString());
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sd = new SaveFileDialog();
            sd.Filter = "Excel spreadsheet|*.xlsx|Old spreadsheet|*.xlsx|CSV|*.csv|All files|*.*";
            sd.DefaultExt = "xlsx";
            if (sd.ShowDialog() != DialogResult.OK) return;
            String path = sd.FileName;

            this.UseWaitCursor = true;
            this.Enabled = false;
            this.button1.Text = SAVE_BUTTON_WORKING;

            Wurkr.Do boundProcess = () => Process(path);
            Wurkr w = new Wurkr(boundProcess);
            w.Done += w_Done;
            w.Doa();
        }

        void w_Done(object sender, EventArgs e)
        {
            this.UseWaitCursor = false;
            this.Invoke((MethodInvoker)delegate {
                this.Enabled = true;
                this.button1.Text = SAVE_BUTTON_NORMAL;
            });
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excel workbooks|*.xlsx|Older excel workbook|*.xls|CSV|*.csv|All Files|*.*";
            od.CheckFileExists = true;
            od.CheckPathExists = true;
            od.Multiselect = true;
            if (od.ShowDialog() != DialogResult.OK) return;
            
            foreach (String fileName in od.FileNames)
            {
                textBox1.AppendText(Environment.NewLine + fileName);
            }
        }
    }
}
