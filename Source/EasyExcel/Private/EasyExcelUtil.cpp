// Fill out your copyright notice in the Description page of Project Settings.


#include "EasyExcelUtil.h"

#include "EasyExcel.h"


using namespace std;

FString UEasyExcelUtil::ReadExcelCellData(const FString& ExcelPath, int32 Row, int32 Column, bool& bResult,
                                          const FString& SheetName)
{
	FString Result;
	bResult = false;

	//xlCreateXMLBook() for xlsx; xlCreateBook() for xls
	if (Book* Book = GetBookFromFile(ExcelPath))
	{
		RegisterKey(Book);

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
				if (const wchar_t* str = Sheet->readStr(Row, Column))
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

TArray<FString> UEasyExcelUtil::ReadExcelRow(const FString& ExcelPath, int32 Row, bool& bResult,
                                             const FString& SheetName)
{
	TArray<FString> Result;
	bResult = false;

	//xlCreateXMLBook() for xlsx; xlCreateBook() for xls
	if (Book* Book = GetBookFromFile(ExcelPath))
	{
		RegisterKey(Book);

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
				if (Row >= Sheet->firstRow() && Row < Sheet->lastRow())
				{
					for (int Column = Sheet->firstCol(); Column < Sheet->lastCol(); ++Column)
					{
						if (const wchar_t* str = Sheet->readStr(Row, Column))
						{
							Result.Add(WCHAR_TO_TCHAR(str));
						}
					}
					bResult = true;
				}
				else
				{
					UE_LOG(LogEasyExcel, Warning, TEXT("Invalid Row Index [%d]"), Row);
				}
			}
		}

		// release
		Book->release();
	}

	return Result;
}

TArray<FString> UEasyExcelUtil::ReadExcelColumn(const FString& ExcelPath, int32 Column, bool& bResult,
                                                const FString& SheetName)
{
	TArray<FString> Result;
	bResult = false;

	//xlCreateXMLBook() for xlsx; xlCreateBook() for xls
	if (Book* Book = GetBookFromFile(ExcelPath))
	{
		RegisterKey(Book);

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
				if (Column >= Sheet->firstCol() && Column < Sheet->lastCol())
				{
					for (int Row = Sheet->firstRow(); Row < Sheet->lastRow(); ++Row)
					{
						if (const wchar_t* str = Sheet->readStr(Row, Column))
						{
							Result.Add(WCHAR_TO_TCHAR(str));
						}
					}
					bResult = true;
				}
				else
				{
					UE_LOG(LogEasyExcel, Warning, TEXT("Invalid Column Index [%d]"), Column);
				}
			}
		}

		// release
		Book->release();
	}

	return Result;
}

bool UEasyExcelUtil::FindExcelCellData(const FString& ExcelPath, const FString& CellData, int32& Row, int32& Column,
                                       const FString& SheetName, FString& R1C1)
{
	Row = -1;
	Column = -1;
	R1C1 = "";

	//xlCreateXMLBook() for xlsx; xlCreateBook() for xls
	if (Book* Book = GetBookFromFile(ExcelPath))
	{
		RegisterKey(Book);

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
				for (int m_Row = Sheet->firstRow(); m_Row < Sheet->lastRow(); ++m_Row)
				{
					for (int m_Column = Sheet->firstCol(); m_Column < Sheet->lastCol(); ++m_Column)
					{
						if (const wchar_t* str = Sheet->readStr(m_Row, m_Column))
						{
							if (CellData.Equals(WCHAR_TO_TCHAR(str)))
							{
								Row = m_Row;
								Column = m_Column;
								R1C1 = ToR1C1(m_Column, m_Row);
								return true;
							}
						}
					}
				}
			}
		}
	}

	UE_LOG(LogEasyExcel, Warning, TEXT("Can not find [%s]"), *CellData);
	return false;
}

bool UEasyExcelUtil::FindExcelCellDataByName(const FString& ExcelPath, FString& CellData, const FString& RowName,
                                             const FString& ColumnName, const FString& SheetName)
{
	//xlCreateXMLBook() for xlsx; xlCreateBook() for xls
	if (Book* Book = GetBookFromFile(ExcelPath))
	{
		RegisterKey(Book);

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
				int32 ColumnIndex = -1, RowIndex = -1;
				// find column index
				for (int m_Column = Sheet->firstCol(); m_Column < Sheet->lastCol(); ++m_Column)
				{
					if (const wchar_t* str = Sheet->readStr(Sheet->firstRow(), m_Column))
					{
						if (WCHAR_TO_TCHAR(str) == ColumnName)
						{
							ColumnIndex = m_Column;
							break;
						}
					}
				}

				for (int m_Row = Sheet->firstRow(); m_Row < Sheet->lastRow(); ++m_Row)
				{
					if (const wchar_t* str = Sheet->readStr(m_Row, Sheet->firstCol()))
					{
						if (WCHAR_TO_TCHAR(str) == RowName)
						{
							RowIndex = m_Row;
							break;
						}
					}
				}

				if (RowIndex >= 0 && ColumnIndex >= 0)
				{
					if (const wchar_t* str = Sheet->readStr(RowIndex, ColumnIndex))
					{
						CellData = WCHAR_TO_TCHAR(str);
						return true;
					}
				}
			}
		}
	}

	UE_LOG(LogEasyExcel, Warning, TEXT("Can not find target celldata."));
	return false;
}

void UEasyExcelUtil::FromR1C1(const FString& R1C1, int32& OutColumn, int32& OutRow)
{
	FString ColumnPart, RowPart;
	bool bColumnPartFinished = false;
	for (const TCHAR Character : R1C1.ToUpper())
	{
		if (FChar::IsAlpha(Character) && !bColumnPartFinished)
		{
			ColumnPart.AppendChar(Character);
		}
		else if (FChar::IsDigit(Character))
		{
			bColumnPartFinished = true;
			RowPart.AppendChar(Character);
		}
	}

	//UE_LOG(LogEasyExcel, Warning, TEXT("ColumnPart [%s]"), *ColumnPart);

	if (ColumnPart.Len() > 0)
	{
		OutColumn = 0;
		for (int32 i = 0; i < ColumnPart.Len(); ++i)
		{
			// UE_LOG(LogEasyExcel, Warning, TEXT("OutColumn =[%d],(ColumnPart[i] - 'A' + 1) = [%d]"), OutColumn,
			//        (ColumnPart[i] - 'A' + 1));
			OutColumn = OutColumn * 26 + (ColumnPart[i] - 'A' + 1);
		}
		OutColumn -= 1; // 1->base => 0->base
	}
	else
	{
		OutColumn = -1;
		UE_LOG(LogEasyExcel, Warning, TEXT("Invalid column part."));
	}


	OutRow = -1;
	if (!RowPart.IsEmpty())
	{
		OutRow = FCString::Atoi(*RowPart) - 1; // 1->base => 0->base
	}
	else
	{
		UE_LOG(LogEasyExcel, Warning, TEXT("Invalid row part."));
	}
}

FString UEasyExcelUtil::ToR1C1(int32 Column, int32 Row)
{
	if (Column < 0 || Row < 0)
	{
		UE_LOG(LogEasyExcel, Warning, TEXT("Invalid Index Column[%d],Row[%d]"), Column, Row);
		return FString();
	}

	FString ColumnPart;
	const FString RowPart = FString::FromInt(Row + 1); // 0->base => 1->base

	while (Column >= 0)
	{
		ColumnPart.InsertAt(0, (Column % 26) + 'A');
		Column = Column / 26 - 1;
	}

	return ColumnPart + RowPart;
}

bool UEasyExcelUtil::CreateExcel(const FString& SaveDirectory, const FString& FileName,
                                 ESupportExcelFileExtension FileExtension, FString SheetName,
                                 TMap<FString, FString> Content)
{
	if (Content.IsEmpty())
	{
		UE_LOG(LogEasyExcel, Warning, TEXT("Empty content to create excel."));
		return false;
	}

	if (!FPaths::DirectoryExists(SaveDirectory))
	{
		UE_LOG(LogEasyExcel, Warning, TEXT("No such directory [%s] while create excel."), *SaveDirectory);
		return false;
	}

	if (SheetName.IsEmpty())
	{
		SheetName = "Sheet1";
	}

	Book* Book;
	FString FullFilePath;
	FString FullDirectory = SaveDirectory;
	if (!FullDirectory.EndsWith("\\"))
	{
		FullDirectory += "\\";
	}

	if (FileExtension == ESupportExcelFileExtension::XLS)
	{
		Book = xlCreateBook();
		FullFilePath = FullDirectory + FileName + ".xls";
	}
	else
	{
		Book = xlCreateXMLBook();
		FullFilePath = FullDirectory + FileName + ".xlsx";
	}

	if (Book)
	{
		RegisterKey(Book);

		Sheet* sheet = Book->addSheet(TCHAR_TO_WCHAR(*SheetName));
		if (sheet)
		{
			int32 Row, Column;
			for (auto Pair : Content)
			{
				FromR1C1(Pair.Key, Column, Row);
				if (Row >= 0 && Column >= 0)
				{
					sheet->writeStr(Row, Column, TCHAR_TO_WCHAR(*Pair.Value));
				}
			}
		}
		Book->save(TCHAR_TO_WCHAR(*FullFilePath));
		Book->release();
		return true;
	}

	return false;
}

void UEasyExcelUtil::RegisterKey(Book* Book)
{
	if (Book)
	{
		Book->setKey(L"liangbochao", L"windows-2e26240202c4e9026db76066a6z6mev7");
		//Book->setKey(L"Za0Shu1",L"windows-2e2521040cc6ea0666b86a6bads4s4id");
	}
}

Book* UEasyExcelUtil::GetBookFromFile(const FString& ExcelPath)
{
	if (FPaths::FileExists(ExcelPath))
	{
		const FString Extension = FPaths::GetExtension(ExcelPath).ToLower();
		if (Extension == "xls")
		{
			return xlCreateBook();
		}
		else if (Extension == "xlsx" /*|| Extension == "xlsm"*/)
		{
			return xlCreateXMLBook();
		}
	}

	UE_LOG(LogEasyExcel, Warning, TEXT("Invalid file format [%s],only support 'xls,xlsx'"), *ExcelPath);
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
