using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace Docx2Pdf
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Processing Files\n");
            string output_path = System.IO.Directory.CreateDirectory(System.IO.Directory.GetCurrentDirectory() + "/output").FullName;
            foreach (string file in System.IO.Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(),"*.docx",System.IO.SearchOption.TopDirectoryOnly))
            {
                string filename = System.IO.Path.GetFileNameWithoutExtension(file);
                try
                {
                    Console.WriteLine(" -- " + System.IO.Path.GetFileName(file));
                    Convert(file, output_path + "/" + filename + ".pdf", WdSaveFormat.wdFormatPDF);
                    Console.WriteLine(" -- -- -- Done");
                }
                catch
                {
                    Console.WriteLine(" -- -- -- Error in " + System.IO.Path.GetFileName(file));
                }
            }
            Console.WriteLine("Successfully Completed, press any key to exit");
            Console.ReadKey();
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
