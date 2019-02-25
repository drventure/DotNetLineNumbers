==============================================================================
DotNetLineNumbers and GenerateLineMap utility
Copyright (c) 2010-2019 Darin Higgins 

The MSBuild Task portion of this project was based on
the excellent "ILMerge_MSBuild_Task" project by 
Emerson Brito and without which, I'd probably never have
figured out how to get the build task functioning!

==============================================================================
ABOUT
==============================================================================

This nuget installs the GenerateLineMap utility and MSBuild task
to automatically pull line number information from an assembly's PDB file
and create a resource from it that can be used for determining line number
information when rendering a stack trace in a production release build.

Access those line numbers by referencing the included DotNetLineNumbers.dll
or by including the DotNetLineNumbers.cs file in you project.

The process uses information from the PDB, but then embeds that info into the
executable as a resource so that you do NOT have to distribute PDBs with your 
release-build application.

Works with DLL or EXE projects.

Important Note:
Due to the nature of the way the resource has to be inserted into the 
compiled assembly, strong names and signing may get removed in the process.

The best workaround I can come up with for now is to use one of the signing 
tools to "resign" the assembly via a Post Build Task.

==============================================================================
