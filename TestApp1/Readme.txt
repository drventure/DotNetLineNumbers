==============================================================================
GenerateLineMap Line Map insertion utility
Darin Higgins 2010-2017
==============================================================================
This nuget installs the GenerateLineMap utility and MSBuild task
to automatically pull line number information from an assembly's PDB file
and create a resource from it that can be used for determining line number
information when rendering a stack trace in a production release build.

The process uses information from the PDB, but then embeds that info into the
executable so that you do NOT have to distribute PDBs with your release
build application.
==============================================================================
