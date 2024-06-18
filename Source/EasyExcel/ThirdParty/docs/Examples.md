# Usage
*[LibXL website](https://www.libxl.com/)*
## generate a new spreadsheet from scratch
```
#include "libxl.h"
using namespace libxl;

int main() 
{
    Book* book = xlCreateBook(); // xlCreateXMLBook() for xlsx
    if(book)
    {
        Sheet* sheet = book->addSheet(L"Sheet1");
        if(sheet)
        {
            sheet->writeStr(2, 1, L"Hello, World !");
            sheet->writeNum(3, 1, 1000);
        }
        book->save(L"example.xls");
        book->release();
    } 
    return 0;
}
```

## extract data from an existing spreadsheet
```
Book* book = xlCreateBook();
if(book)
{
    if(book->load(L"example.xls"))
    {
        Sheet* sheet = book->getSheet(0);
        if(sheet)
        {
            const wchar_t* s = sheet->readStr(2, 1);
            if(s) wcout << s << endl;

            double d = sheet->readNum(3, 1);
            cout << d << endl;
        }
    }

    book->release();
}
```

## edit an existing spreadsheet
```
Book* book = xlCreateBook();
if(book) 
{                
    if(book->load(L"example.xls"))
    {
        Sheet* sheet = book->getSheet(0);
        if(sheet) 
        {   
            double d = sheet->readNum(3, 1);
            sheet->writeNum(3, 1, d * 2);
            sheet->writeStr(4, 1, L"new string");
        }
        book->save(L"example.xls");
    }

    book->release();   
}
```

## apply formatting options
```
Font* font = book->addFont();
font->setName(L"Impact");
font->setSize(36);        

Format* format = book->addFormat();
format->setAlignH(ALIGNH_CENTER);
format->setBorder(BORDERSTYLE_MEDIUMDASHDOTDOT);
format->setBorderColor(COLOR_RED);
format->setFont(font);
           
Sheet* sheet = book->addSheet(L"Custom");
if(sheet)
{
    sheet->writeStr(2, 1, L"Format", format);
    sheet->setCol(1, 1, 25);
}

book->save(L"format.xls");
```

## Invoice example
```
#include "libxl.h"
using namespace libxl;

int main()
{
    Book* book = xlCreateBook();
    if(book) 
    {   
        Font* boldFont = book->addFont();
        boldFont->setBold();

        Font* titleFont = book->addFont();
        titleFont->setName(L"Arial Black");
        titleFont->setSize(16);

        Format* titleFormat = book->addFormat();
        titleFormat->setFont(titleFont);

        Format* headerFormat = book->addFormat();
        headerFormat->setAlignH(ALIGNH_CENTER);
        headerFormat->setBorder(BORDERSTYLE_THIN);
        headerFormat->setFont(boldFont);        
        headerFormat->setFillPattern(FILLPATTERN_SOLID);
        headerFormat->setPatternForegroundColor(COLOR_TAN);

        Format* descriptionFormat = book->addFormat();
        descriptionFormat->setBorderLeft(BORDERSTYLE_THIN);

        Format* amountFormat = book->addFormat();
        amountFormat->setNumFormat(NUMFORMAT_CURRENCY_NEGBRA);
        amountFormat->setBorderLeft(BORDERSTYLE_THIN);
        amountFormat->setBorderRight(BORDERSTYLE_THIN);
                
        Format* totalLabelFormat = book->addFormat();
        totalLabelFormat->setBorderTop(BORDERSTYLE_THIN);
        totalLabelFormat->setAlignH(ALIGNH_RIGHT);
        totalLabelFormat->setFont(boldFont);

        Format* totalFormat = book->addFormat();
        totalFormat->setNumFormat(NUMFORMAT_CURRENCY_NEGBRA);
        totalFormat->setBorder(BORDERSTYLE_THIN);
        totalFormat->setFont(boldFont);
        totalFormat->setFillPattern(FILLPATTERN_SOLID);
        totalFormat->setPatternForegroundColor(COLOR_YELLOW);

        Format* signatureFormat = book->addFormat();
        signatureFormat->setAlignH(ALIGNH_CENTER);
        signatureFormat->setBorderTop(BORDERSTYLE_THIN);
             
        Sheet* sheet = book->addSheet(L"Invoice");
        if(sheet)
        {
            sheet->writeStr(2, 1, L"Invoice No. 3568", titleFormat);

            sheet->writeStr(4, 1, L"Name: John Smith");
            sheet->writeStr(5, 1, L"Address: San Ramon, CA 94583 USA");

            sheet->writeStr(7, 1, L"Description", headerFormat);
            sheet->writeStr(7, 2, L"Amount", headerFormat);

            sheet->writeStr(8, 1, L"Ball-Point Pens", descriptionFormat);
            sheet->writeNum(8, 2, 85, amountFormat);
            sheet->writeStr(9, 1, L"T-Shirts", descriptionFormat);
            sheet->writeNum(9, 2, 150, amountFormat);
            sheet->writeStr(10, 1, L"Tea cups", descriptionFormat);
            sheet->writeNum(10, 2, 45, amountFormat);

            sheet->writeStr(11, 1, L"Total:", totalLabelFormat);
            sheet->writeNum(11, 2, 280, totalFormat);

            sheet->writeStr(14, 2, L"Signature", signatureFormat);

            sheet->setCol(1, 1, 40);
            sheet->setCol(2, 2, 15);
          }

          book->save(L"invoice.xls");       
          book->release();   
    }

    return 0;
}
```
