using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Resources;

// Shared Assembly info for All Projects
#region "Static Assembly Version info
[assembly: AssemblyCompany("Darin Higgins")]
[assembly: AssemblyProduct("GenerateLineMap")]
[assembly: AssemblyCopyright("Copyright © 2018 Darin Higgins")]
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

[assembly: NeutralResourcesLanguageAttribute( "en-US" )]

[assembly: AssemblyVersion(VersionInfo.Version)]
[assembly: AssemblyFileVersion(VersionInfo.Version)]
//this attrib supports semantic versioning
[assembly: AssemblyInformationalVersion(VersionInfo.Version)]
#endregion

/// <summary>
/// This allows us to have a single place for version info
///	This also accommodates Nuget Semantic versioning
/// </summary>
internal struct VersionInfo
{
    public const string Version = "0.0.0.32";
}
