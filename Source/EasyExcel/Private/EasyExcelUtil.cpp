// Fill out your copyright notice in the Description page of Project Settings.


#include "EasyExcelUtil.h"

#include "EasyExcel.h"


using namespace std;

FString UEasyExcelUtil::ReadSheetValue(const FString& ExcelPath, int32 Row, int32 Column, bool& bResult,
                                  const FString& SheetName)
{
	FString Result;
	bResult = false;

	//xlCreateXMLBook() for xlsx; xlCreateBook() for xls
	if (Book* Book = GetBookFromFile(ExcelPath))
	{
		RegisterKey(Book);
		// do something
		// FString to const wchar_t *
		if (Book->load(*ExcelPath))
		{
			Sheet* Sheet;
			if (SheetName.IsEmpty())
			{
				Sheet = Book->getSheet(0);
			}
			else
			{
				Sheet = GetSheetByName(Book,TCHAR_TO_WCHAR(*SheetName));
			}

			if (Sheet)
			{
				const wchar_t* str = Sheet->readStr(Row, Column);
				if (str)
				{
					Result = WCHAR_TO_TCHAR(str);
					//UE_LOG(LogEasyExcel,Display,TEXT("Read value from sheet: %s"),*Result);
					bResult = true;
				}
				else
				{
					UE_LOG(LogEasyExcel, Warning, TEXT("Can not read value from Row[%d],Column[%d]"), Row, Column);
				}
			}
		}

		// release
		Book->release();
	}

	return Result;
}

bool UEasyExcelUtil::TestCreateExcel()
{
	Book* Book = xlCreateBook();
	if(Book)
	{
		RegisterKey(Book);

		Sheet* sheet = Book->addSheet(L"Sheet1");
		if(sheet)
		{
			sheet->writeStr(2, 1, L"Hello, World !");
			sheet->writeNum(3, 1, 1000);
		}
		Book->save(L"D:\\Projects\\GitProjects\\EasyExcelPJ\\Plugins\\EasyExcel\\xls\\example.xls");
		Book->release();
	}
	
	return false;
}

void UEasyExcelUtil::RegisterKey(Book* Book)
{
	if(Book)
	{
		Book->setKey(L"liangbochao",L"windows-2e26240202c4e9026db76066a6z6mev7");
		//Book->setKey(L"Za0Shu1",L"windows-2e2521040cc6ea0666b86a6bads4s4id");
	}
}

Book* UEasyExcelUtil::GetBookFromFile(const FString& ExcelPath)
{
	if(FPaths::FileExists(ExcelPath))
	{
		const FString Extension = FPaths::GetExtension(ExcelPath).ToLower();
		if (Extension == "xls")
		{
			return xlCreateBook();
		}
		else if(Extension == "xlsx" || Extension == "xlsm")
		{
			return xlCreateXMLBook();
		}
	}

	UE_LOG(LogEasyExcel, Warning, TEXT("Invalid file format [%s],only support 'xls,xlsx/xlsm'"), *ExcelPath);
	return nullptr;
}

Sheet* UEasyExcelUtil::GetSheetByName(const Book* Book, const wchar_t* SheetName)
{
	const int SheetCount = Book->sheetCount();
	for (int i = 0; i < SheetCount; ++i)
	{
		Sheet* Sheet = Book->getSheet(i);
		if (Sheet && wcscmp(Sheet->name(), SheetName) == 0)
		{
			return Sheet;
		}
	}

	UE_LOG(LogEasyExcel, Warning, TEXT("Can not found sheet named [%s]"), WCHAR_TO_TCHAR(SheetName));
	return nullptr;
}
