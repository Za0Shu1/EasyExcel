// Fill out your copyright notice in the Description page of Project Settings.

#pragma once

#include "CoreMinimal.h"
#include "Kismet/BlueprintFunctionLibrary.h"
#include "libxl.h"
#include "EasyExcelUtil.generated.h"

/**
 * 
 */

using namespace libxl;

UCLASS()
class EASYEXCEL_API UEasyExcelUtil : public UBlueprintFunctionLibrary
{
	GENERATED_BODY()

public:
	UFUNCTION(BlueprintCallable,Category = "Easy Excel | Access",meta = (AdvancedDisplay = "SheetName"))
	static FString ReadSheetValue(const FString& ExcelPath, int32 Row, int32 Column, bool& bResult,const FString& SheetName);


	// ====Test Start====
	static bool TestCreateExcel();
	// ====Test End======

	// helper functions
	static void RegisterKey(Book* Book);
	static Book* GetBookFromFile(const FString& ExcelPath);
	static Sheet* GetSheetByName(const Book* Book,const wchar_t* SheetName);
};
