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
using System.Windows.Forms;
using Threads = System.Threading;

namespace SummerSchoolExcelParserDeux
{
    /// <summary>
    /// Helper class that assists in running a background task and notifying it's done
    /// </summary>
    internal class Wurkr
    {
        /// <summary>
        /// Event fired when task is done
        /// </summary>
        public event System.EventHandler Done;
        public delegate void Do();

        private Do do_;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="f">the task to be run</param>
        public Wurkr(Do f)
        {
            do_ = f;
        }

        /// <summary>
        /// Run the task
        /// </summary>
        public void Doa()
        {
            Threads.ThreadStart starter = new Threads.ThreadStart(() =>
            {
                try
                {
                    do_();
                }
                finally
                {
                    Done.Invoke(this, new EventArgs());
                }
            });
            
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

            foreach (String fileName in from line in this.textBox1.Lines where line.Length > 0 select line)
            {
                try
                {
                    ExcelParser ep = new ExcelParser(fileName);
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
            // extensions are very strict
            SaveFileDialog sd = new SaveFileDialog()
            {
                Filter = "Excel spreadsheet|*.xlsx|Old spreadsheet|*.xls|CSV|*.csv",
                DefaultExt = "xlsx"
            };
            if (sd.ShowDialog() != DialogResult.OK) return;

            UseWaitCursor = !(Enabled = false); 
            button1.Text = SAVE_BUTTON_WORKING;

            // bind the path parameter because the work delegate needs to be void(void)
            Wurkr w = new Wurkr(() => Process(sd.FileName));
            // notify when done
            w.Done += (EventHandler)((sndr, evar) => {
                // UI update needs to be done on the main thread
                this.Invoke((MethodInvoker)delegate
                {
                    this.UseWaitCursor = !(Enabled = true);
                    this.button1.Text = SAVE_BUTTON_NORMAL;
                });
            });

            // start background task
            w.Doa();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog()
            {
                Filter = "Excel workbooks|*.xlsx|Older excel workbook|*.xls|CSV|*.csv|All Files|*.*",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = true
            };
            if (od.ShowDialog() != DialogResult.OK) return;
            
            foreach (String fileName in od.FileNames) textBox1.AppendText(Environment.NewLine + fileName);
        }
    }
}
