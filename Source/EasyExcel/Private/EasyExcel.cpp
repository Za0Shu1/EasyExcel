// Copyright Epic Games, Inc. All Rights Reserved.

#include "EasyExcel.h"

#include "Interfaces/IPluginManager.h"

#define LOCTEXT_NAMESPACE "FEasyExcelModule"
DEFINE_LOG_CATEGORY(LogEasyExcel);

void FEasyExcelModule::StartupModule()
{
	// This code will execute after your module is loaded into memory; the exact timing is specified in the .uplugin file per-module

	const TSharedPtr<IPlugin> PluginInterface = IPluginManager::Get().FindPlugin("EasyExcel");
	if (PluginInterface)
	{
		// const FString LibXlPath = FPaths::Combine(PluginInterface->GetBaseDir(), TEXT("Source"),TEXT("EasyExcel"),
		// 								TEXT("ThirdParty"), TEXT("libxl-win-4.2.0"),TEXT("bin"),
		// 								TEXT("libxl.dll"));

		const FString LibXlPath = FPaths::Combine(PluginInterface->GetBaseDir(), TEXT("Binaries"),TEXT("Win64"),
		                                    TEXT("libxl.dll"));
		UE_LOG(LogEasyExcel, Warning, TEXT("%s"), *LibXlPath);
		// load dll
		LibXlHandle = FPlatformProcess::GetDllHandle(*LibXlPath);
		if (!LibXlHandle)
		{
			UE_LOG(LogEasyExcel, Warning, TEXT("Can not load libxl64 dll libary"));
		}
	}
}

void FEasyExcelModule::ShutdownModule()
{
	// This function may be called during shutdown to clean up your module.  For modules that support dynamic reloading,
	// we call this function before unloading the module.

	// free dll
	if (LibXlHandle)
	{
		FPlatformProcess::FreeDllHandle(LibXlHandle);
	}
}

#undef LOCTEXT_NAMESPACE

IMPLEMENT_MODULE(FEasyExcelModule, EasyExcel)
