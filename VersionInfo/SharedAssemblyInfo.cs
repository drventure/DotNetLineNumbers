using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// Shared Assembly info for GenerateLineMap Projects
[assembly: AssemblyCompany("Darin Higgins")]
[assembly: AssemblyProduct("GenerateLineMap")]
[assembly: AssemblyCopyright("Copyright © 2017 Darin Higgins")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

#if DEBUG
[assembly: AssemblyConfiguration("Debug")]
#else
[assembly: AssemblyConfiguration("Release")]
#endif

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
[assembly: AssemblyVersion(VersionInfo.Version)]
[assembly: AssemblyFileVersion(VersionInfo.Version)]
//this attrib supports semantic versioning
[assembly: AssemblyInformationalVersion(VersionInfo.Version)]

/// <summary>
/// This allows us to have a single place for version info
/// </summary>
internal struct VersionInfo
{
	public const string Version = "0.0.0.6";
	//this accommodates Nuget Semantic versioning	
}
