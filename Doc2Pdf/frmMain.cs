using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace Doc2Pdf
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }
        
        public void ProgressMessage(string Message)
        {
            lstConverted.Invoke(new Action(() => lstConverted.Items.Add(Message)));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (txtPassword.Text.Equals("P@ssw0rd"))
            {
                pnlLogin.Visible = false;
                this.Height = 524;
                pnlMain.Visible = true;
                this.CenterToScreen();
            }
            else
                MessageBox.Show("Invalid Password");
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            lstFiles.Items.AddRange(Directory.GetFiles(Directory.GetCurrentDirectory(), "*.docx", SearchOption.TopDirectoryOnly).ToList().Select(m=> Path.GetFileName(m)).ToArray());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Threading.Thread t = new System.Threading.Thread(new System.Threading.ThreadStart(InitiateProcess));
            t.Start();
        }

        public void InitiateProcess()
        {
            ProgressMessage("Processing Files\n");
            string output_path = Directory.CreateDirectory(Directory.GetCurrentDirectory() + "/output").FullName;
            foreach (string file in Directory.GetFiles(Directory.GetCurrentDirectory(), "*.docx", SearchOption.TopDirectoryOnly))
            {
                string filename = Path.GetFileNameWithoutExtension(file);
                try
                {
                    ProgressMessage(" -- " + Path.GetFileName(file));
                    Convert(file, output_path + "/" + filename + ".pdf", WdSaveFormat.wdFormatPDF);
                    ProgressMessage(" -- -- -- Done");
                }
                catch
                {
                    ProgressMessage(" -- -- -- Error in " + Path.GetFileName(file));
                }
            }
            ProgressMessage("Successfully Completed");
        }
        
        public static void Convert(string input, string output, WdSaveFormat format)
        {
            // Create an instance of Word.exe
            Word._Application oWord = new Word.Application();

            // Make this instance of word invisible (Can still see it in the taskmgr).
            oWord.Visible = false;

            // Interop requires objects.
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = false;
            object oInput = input;
            object oOutput = output;
            object oFormat = format;

            // Load a document into our instance of word.exe
            Word._Document oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Make this document the active document.
            oDoc.Activate();

            // Save this document in Word 2003 format.
            oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            // Always close Word.exe.
            oWord.Quit(ref oMissing, ref oMissing, ref oMissing);
        }
    }
}
