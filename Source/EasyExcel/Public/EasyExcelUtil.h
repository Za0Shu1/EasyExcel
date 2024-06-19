// Fill out your copyright notice in the Description page of Project Settings.

#pragma once

#include "CoreMinimal.h"
#include "Kismet/BlueprintFunctionLibrary.h"
#include "libxl.h"
#include "EasyExcelUtil.generated.h"

/**
 * 
 */

UENUM(BlueprintType)
enum class ESupportExcelFileExtension : uint8
{
	XLS,
	XLSX
};

using namespace libxl;

UCLASS()
class EASYEXCEL_API UEasyExcelUtil : public UBlueprintFunctionLibrary
{
	GENERATED_BODY()

public:
	UFUNCTION(BlueprintCallable, Category = "Easy Excel | Read", meta = (AdvancedDisplay = "SheetName"))
	static FString ReadExcelCellData(const FString& ExcelPath, int32 Row, int32 Column, bool& bResult,
	                                 const FString& SheetName);

	UFUNCTION(BlueprintCallable, Category = "Easy Excel | Read", meta = (AdvancedDisplay = "SheetName"))
	static TArray<FString> ReadExcelRow(const FString& ExcelPath, int32 Row, bool& bResult, const FString& SheetName);

	UFUNCTION(BlueprintCallable, Category = "Easy Excel | Read", meta = (AdvancedDisplay = "SheetName"))
	static TArray<FString> ReadExcelColumn(const FString& ExcelPath, int32 Column, bool& bResult,
	                                       const FString& SheetName);

	UFUNCTION(BlueprintCallable, Category = "Easy Excel | Find", meta = (AdvancedDisplay = "SheetName"))
	static bool FindExcelCellData(const FString& ExcelPath, const FString& CellData, int32& Row, int32& Column,
	                              const FString& SheetName, FString& R1C1);

	UFUNCTION(BlueprintCallable, Category = "Easy Excel | Find", meta = (AdvancedDisplay = "SheetName"))
	static bool FindExcelCellDataByName(const FString& ExcelPath, FString& CellData, const FString& RowName, const FString& ColumnName,
	                                    const FString& SheetName);

	UFUNCTION(BlueprintPure, Category = "Easy Excel | Helper")
	static void FromR1C1(const FString& R1C1, int32& OutColumn, int32& OutRow);

	UFUNCTION(BlueprintPure, Category = "Easy Excel | Helper")
	static FString ToR1C1(int32 Column, int32 Row);

	// ====Test Start====
	UFUNCTION(BlueprintCallable, Category = "Easy Excel | Create", meta = (AdvancedDisplay = "SheetName"))
	static bool CreateExcel(const FString& SaveDirectory, const FString& FileName,
	                            ESupportExcelFileExtension FileExtension, FString SheetName,TMap<FString,FString> Content);
	// ====Test End======

	// helper functions
	static void RegisterKey(Book* Book);
	static Book* GetBookFromFile(const FString& ExcelPath);
	static Sheet* GetSheetByName(const Book* Book, const wchar_t* SheetName);

	// library functions

	// int firstRow() const;
	// int lastRow() const;
	// int firstCol() const;
	// int lastCol() const;

	// CellType cellType(int row,int col) const;
	// enum CellType {CELLTYPE_EMPTY, CELLTYPE_NUMBER, CELLTYPE_STRING, CELLTYPE_BOOLEAN, CELLTYPE_BLANK, CELLTYPE_ERROR, CELLTYPE_STRICTDATE};

	// bool isFormula(int row,int col) const;
	// const wchar_t* readFormula(int row,int col,Format** format = 0);

	// Book::errorMessage() // get error message

	// R1C1
	// (Column)
	// A->XFD =>24*26^2 + 6*26 + 4 = 16384 2^14
	// wps:A->IV = 9*26 + 22 = 256 = 2^8
	// (Row)
	// 1->1048576 2^20
	// wps:1->65536 = 2^16
};
