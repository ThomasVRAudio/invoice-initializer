using Invoice_Initializer;

CreateDocument createDocument = new CreateDocument();
ClientBook clientBook = new ClientBook();

string clientBookPath = @"D:\Users\Thomas\Documents\Docx\Excel\TVRA_ClientBook.xlsx";
string invoicePath = @"D:\Users\Thomas\Desktop\Facturen\Invoices";
string invoiceInitDocPath = @"D:\Users\Thomas\Documents\Docx\DOCX\Invoice_Init.docx";
string generatedDocPath = @"D:\Users\Thomas\Documents\Docx\DOCX\GeneratedInvoice.docx";

Console.Write("Company ID: ");
clientBook.GetInfo(Convert.ToInt32(Console.ReadLine()), clientBookPath);

Console.Write("Amount: ");
double amount = Convert.ToDouble(Console.ReadLine());

createDocument.CreateWordDocument(invoiceInitDocPath, 
    generatedDocPath, amount, clientBook, invoicePath);
