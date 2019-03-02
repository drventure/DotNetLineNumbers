==============================================================================
DotNetLineNumbers and GenerateLineMap utility
Copyright (c) 2010-2019 Darin Higgins 

The MSBuild Task portion of this project was based on
the excellent "ILMerge_MSBuild_Task" project by 
Emerson Brito and without which, I'd probably never have
figured out how to get the build task functioning!

==============================================================================
INSTALLING
==============================================================================

VS Package Manager

	PM> Install-Package DotNetLineNumbers -Version 1.0.0 


==============================================================================
REQUIREMENTS
==============================================================================

Currently, the nuget is built with .net framework 4.5.1.
It has not been tested with any other versions.


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
BUILDING
==============================================================================

Once you've add the nuget to your project, you should be able to build as 
normal. 

However, it is critical that for both DEBUG and RELEASE builds, you enable
PDB generation.

Do this as follows: 
	1. Open the PROPERTIES page for your project
	2. Click the BUILD tab
	3. Click the ADVANCED button
	4. For DEBUGGING INFORMATION, choose FULL or PDB-ONLY
	5. Repeat the above for both DEBUG and RELEASE configurations

Note that once the build is complete, you can delete the *.PDB files.

You DO NOT need to distribute the PDB files with your application.


==============================================================================
QUICK START USAGE
==============================================================================

When you've added the package to your project, the nuget will automatically
add a reference to DotNetLineNumbers.Utilities.dll.

You can use this reference directly, or, if you prefer, delete that reference
and add the file: 

	DotNetLineNumbers.cs 

located in .\packages\DotNetLineNumbers[version] to your project.
The DLL is directly compiled from this .cs file.

To render an exception with additional stack trace information and line numbers,
call:

	var str = ex.ToStringEnhanced();
	System.Diagnostics.Debug.WriteLine(str);

To retrieve a stack trace for the exception that contains enhanced stack frame
objects, call:

	var trace = ex.GetEnhancedStackTrace();

Once you have an EnhancedStackTrace object, retrieve frames as you would
with a normal StackTrace object:

	var frame = trace.GetFrame(0);

With an EnhancedStackFrame object, you can query for specific properties, 
in particular, the source filename and the line number in that file
corresponding to the stack frame:

	var file = frame.GetFileName();
	var line = frame.GetFileLineNumber();
	var col = frame.GetFileColumnNumber();

For more details, see the help file.


==============================================================================
HELP FILE
==============================================================================

For detailed help, see the locally installed help file here:

	.\packages\DotNetLineNumbers[version]\DotNetLineNumbers.chm

