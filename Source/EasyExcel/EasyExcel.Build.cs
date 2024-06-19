// Copyright Epic Games, Inc. All Rights Reserved.

using System.IO;
using UnrealBuildTool;

public class EasyExcel : ModuleRules
{
	public EasyExcel(ReadOnlyTargetRules Target) : base(Target)
	{
		PCHUsage = ModuleRules.PCHUsageMode.UseExplicitOrSharedPCHs;

		if (Target.Platform == UnrealTargetPlatform.Win64)
		{
			// Add the import library
			PublicAdditionalLibraries.Add(Path.Combine(ModuleDirectory, "ThirdParty", "libxl-win-4.2.0", "lib",
				"libxl.lib"));

			// Delay-load the DLL, so we can load it from the right place first
			PublicDelayLoadDLLs.Add("libxl.dll");

			// Ensure that the DLL is staged along with the executable
			RuntimeDependencies.Add("$(BinaryOutputDir)/libxl.dll",
				"$(ModuleDir)/ThirdParty/libxl-win-4.2.0/bin/libxl.dll");
			
			RuntimeDependencies.Add(Path.Combine("$(PluginDir)", "xls", "..."),StagedFileType.NonUFS);
		}

		PublicIncludePaths.Add(Path.Combine(ModuleDirectory, "ThirdParty", "libxl-win-4.2.0", "include"));

		PublicIncludePaths.AddRange(
			new string[]
			{
			}
		);


		PrivateIncludePaths.AddRange(
			new string[]
			{
				// ... add other private include paths required here ...
			}
		);


		PublicDependencyModuleNames.AddRange(
			new string[]
			{
				"Core",
				// ... add other public dependencies that you statically link with here ...
			}
		);


		PrivateDependencyModuleNames.AddRange(
			new string[]
			{
				"CoreUObject",
				"Engine",
				"Slate",
				"SlateCore",
				"Projects"
				// ... add private dependencies that you statically link with here ...	
			}
		);


		DynamicallyLoadedModuleNames.AddRange(
			new string[]
			{
				// ... add any modules that your module loads dynamically here ...
			}
		);
	}
}