# Invoice Initializer

### Initial file setup:
Within Program.cs:
```
string clientBookPath = "excel path";
string invoicePath = "path where you store all your invoices";
string invoiceInitDocPath = "your default init invoice word document path";
string generatedDocPath = @"path where you want your generated document to be put";
```
### Excel file structure:
Start at the second row, second line, in the following order:

![image](https://user-images.githubusercontent.com/38658020/200164804-e8aa4255-8c6a-4f7c-948a-839a2043275e.png)

### Tags
Tags you can use in your word document:

&lt;Company>

&lt;StreetAdress>

&lt;Province>

&lt;ZipCode>

&lt;Email>

&lt;Cost>

&lt;Tax>

&lt;Total>

&lt;Id>

&lt;Date>

&lt;DueDate>

&lt;InvoiceNumber>

### Example Initial Document
![image](https://user-images.githubusercontent.com/38658020/200164768-be544353-3a14-4197-b085-425d5a7693f3.png)

### Launch
- Step One: Enter Client ID
- Step Two: Enter Cost

That's it! It will generate a filled in invoice document for you
