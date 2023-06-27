using Microsoft.Vbe.Interop;
using Praktika2.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApp2
{
    internal class WordPaster
    {
        private FileInfo _fileInfo;

        // Constructor to initialize WordPaster with a file name
        public WordPaster(string fileName)
        {
            // Check if the file exists
            if (File.Exists(fileName))
            {
                _fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File does not exist");
            }
        }

        // Process the given items (key-value pairs) in the Word document
        internal bool Process(Dictionary<string, string> items)
        {

            Word.Application app = null;

            try
            {
                app = new Word.Application();
                app.Visible = true;
                Object file = _fileInfo.FullName;
                Object missing = Type.Missing;
                app.Documents.Open(file);

                // Find and replace the keys in the document
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object warp = Word.WdFindWrap.wdFindContinue;
                    Object replace = Word.WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Format: false,
                        ReplaceWith: missing,
                        Replace: replace);

                    // Set the font and size for the replaced text
                    Word.Range replacedRange = app.Selection.Range;
                    replacedRange.Font.Name = "Times New Roman";
                    replacedRange.Font.Size = 11;
                }


                // Save the document
                string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Документы дирекция КГЭУ");
                // Check if the folder exists, create it if it doesn't
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }
                string newFileName = Path.Combine(folderPath, DateTime.Now.ToString("dd.MM.yyyy HH.mm") + " " + _fileInfo.Name);
                app.ActiveDocument.SaveAs2(newFileName);
                MessageBox.Show(
                    "Документ сохранен на рабочем столе в папке ''Документы дирекция КГЭУ''",
                    "Успешно");
                app.Application.Documents.Open(Path.Combine(folderPath, newFileName));

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Quit the Word application
                if (app != null)
                {
                    /*app.Quit();*/
                }
            }
            return false;
        }
    }
}
