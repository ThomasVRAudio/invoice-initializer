using Invoice_Initializer;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Security.AccessControl;
using Word = Microsoft.Office.Interop.Word;

namespace Invoice_Initializer
{
    public class CreateDocument
    {
        public FindAndReplace findAndReplace;
        public ClientBook clientBook;
        private int file_count = 0;

        public CreateDocument()
        {
            findAndReplace = new FindAndReplace();
        }

        public void CreateWordDocument(object filename, object SaveAs, double cost, ClientBook clientBook, string invoicePath)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document? myWordDoc = null;

            if (File.Exists((string)filename)){
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = true;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);

                myWordDoc.Activate();

                if (clientBook == null) { Console.WriteLine("Error: No Client"); return; }

                    findAndReplace.FindReplace(wordApp,
                    "<Company>", clientBook.clientCompany);

                    findAndReplace.FindReplace(wordApp,
                    "<StreetAdress>", clientBook.clientAdress);

                    findAndReplace.FindReplace(wordApp,
                    "<Province>", clientBook.clientProvince);

                    findAndReplace.FindReplace(wordApp,
                    "<ZipCode>", clientBook.clientZipCode);

                    findAndReplace.FindReplace(wordApp,
                    "<Email>", clientBook.clientEmail);

                    findAndReplace.FindReplace(wordApp,
                    "<Cost>", cost.ToString("0.00"));

                    findAndReplace.FindReplace(wordApp,
                    "<Tax>", Math.Round(cost * 0.21f, 2).ToString("0.00"));

                    findAndReplace.FindReplace(wordApp,
                    "<Total>", Math.Round(cost * 1.21f, 2).ToString("0.00"));

                    findAndReplace.ReplaceHeader(wordApp, "<Id>", Convert.ToString(clientBook.clientID), myWordDoc);
                    findAndReplace.ReplaceHeader(wordApp, "<Date>", GetDate(), myWordDoc);
                    findAndReplace.ReplaceHeader(wordApp, "<DueDate>", GetDueDate(), myWordDoc);

                    GetDirectoryFileCount(invoicePath);

                    findAndReplace.ReplaceHeader(wordApp, "<InvoiceNumber>", file_count, myWordDoc);

            } else
            {
                Console.WriteLine("File Not Found");
            }

            myWordDoc?.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            Console.WriteLine("File Created");
        }

        private string GetDate()
        {
            return DateTime.Today.ToShortDateString();
        }

        private string GetDueDate()
        {
            return DateTime.Today.AddMonths(1).ToShortDateString();
        }

        public void GetDirectoryFileCount(string dir)
        {
            dir = dir + @"\";
            String[] all_files = Directory.GetFileSystemEntries(dir);

            foreach (string file in all_files)
            {
                if (Directory.Exists(file))
                {
                    GetDirectoryFileCount(file);
                }
                else
                {
                    file_count++;
                }
            }
        }
    }
}
