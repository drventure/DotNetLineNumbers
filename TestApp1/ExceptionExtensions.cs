#region MIT License
/*
    MIT License

    Copyright (c) 2017 Darin Higgins

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
 */
#endregion

using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;

using InternalExtensionMethods;

namespace ExceptionExtensions
{
	public static class ExceptionExtensions
	{
		/// <summary>
		/// Force the use of the PDB if available
		/// This is primarily used when debugging/Unit testing
		/// </summary>
		public static bool AllowUseOfPDB {get; set;}


		/// <summary>
		/// Used to identify stack frames that relate to "this" class so they can be skipped
		/// </summary>
		private static string CLASSNAME = "ExceptionExtensions";


		/// <summary>
		/// turns a single stack frame object into an informative string
		/// </summary>
		/// <param name="sf"></param>
		/// <returns></returns>
		private static string StackFrameToString(StackFrame sf)
		{
			StringBuilder sb = new StringBuilder();
			int intParam = 0;
			MemberInfo mi = sf.GetMethod();

			//build method name
			sb.Append("   ");
			string MethodName = mi.DeclaringType.Namespace + "." + mi.DeclaringType.Name + "." + mi.Name;
			sb.Append(MethodName);

			//build method params
			ParameterInfo[] objParameters = sf.GetMethod().GetParameters();
			sb.Append("(");
			intParam = 0;
			foreach (ParameterInfo objParameter in objParameters)
			{
				intParam += 1;
				if (intParam > 1)
					sb.Append(", ");
				sb.Append(objParameter.Name);
				sb.Append(" As ");
				sb.Append(objParameter.ParameterType.Name);
			}
			sb.AppendLine(")");

			// if source code is available, append location info
			sb.Append("       ");
			if (sf.GetFileName() != null && sf.GetFileName().Length != 0 && ExceptionExtensions.AllowUseOfPDB)
			{
				// the PDB appears to be available, since the above elements are 
				// not blank, so just use it's information

				sb.Append(System.IO.Path.GetFileName(sf.GetFileName()));
				var Line = sf.GetFileLineNumber();
				if (Line != 0)
				{
					sb.Append(": line ");
					sb.Append(string.Format("{0}", Line));
				}
				var col = sf.GetFileColumnNumber();
				if (col != 0)
				{
					sb.Append(", col ");
					sb.Append(string.Format("{0:#00}", sf.GetFileColumnNumber()));
				}
				// if IL is available, append IL location info
				if (sf.GetILOffset() != StackFrame.OFFSET_UNKNOWN)
				{
					sb.Append(", IL ");
					sb.Append(string.Format("{0:#0000}", sf.GetILOffset()));
				}
			}
			else
			{
				// the PDB is not available, so attempt to retrieve 
				// any embedded linemap information
				string Filename;
				if (ParentAssembly != null)
				{
					Filename = System.IO.Path.GetFileName(ParentAssembly.CodeBase);
				}
				else
				{
					Filename = "Unable to determine Assembly Filename";
				}

				sb.Append(Filename);
				// Get the native code offset and convert to a line number
				// first, make sure our linemap is loaded
				try
				{
					LineMap.AssemblyLineMaps.Add(sf.GetMethod().DeclaringType.Assembly);

					var sl = MapStackFrameToSourceLine(sf);
					if (sl.Line != 0)
					{
						sb.Append(": Source File - ");
						sb.Append(sl.SourceFile);
						sb.Append(": line ");
						sb.Append(string.Format("{0}", sl.Line));
					}

				}
				catch (Exception ex)
				{
					// any problems in loading the Linemap, just write to debugger and call it a day
					Debug.WriteLine(string.Format("Unable to load line map information. Error: {0}", ex.ToString()));
				}
				finally
				{
					// native code offset is always available
					var IL = sf.GetILOffset();
					if (IL != StackFrame.OFFSET_UNKNOWN)
					{
						sb.Append(": IL ");
						sb.Append(string.Format("{0:#00000}", IL));
					}
				}
			}
			sb.AppendLine();
			return sb.ToString();
		}


		private static Assembly _parentAssembly = null;
		/// <summary>
		/// Retrieve the root assembly of the executing assembly
		/// </summary>
		/// <returns></returns>
		private static Assembly ParentAssembly
		{
			get
			{
				if (_parentAssembly == null)
				{
					if (Assembly.GetEntryAssembly() != null)
					{
						_parentAssembly = Assembly.GetEntryAssembly();
					}
					else if (Assembly.GetCallingAssembly() != null)
					{
						_parentAssembly = Assembly.GetCallingAssembly();
					}
					else
					{
						//TODO questionable
						_parentAssembly = Assembly.GetExecutingAssembly();
					}
				}
				return _parentAssembly;
			}
		}


		/// <summary>
		/// Map an address offset from a stack frame entry to a linenumber
		/// using the Method name, the base address of the method and the
		/// IL offset from the base address
		/// </summary>
		/// <param name="sf"></param>
		/// <param name="Line"></param>
		/// <param name="SourceFile"></param>
		/// <remarks></remarks>
		private static SourceLine MapStackFrameToSourceLine(StackFrame sf)
		{
			// first, get the base addr of the method
			// if possible
			var sl = new SourceLine();

			// you have to have symbols to do this
			if (LineMap.AssemblyLineMaps.Count == 0)
				return sl;

			// first, check if for symbols for the assembly for this stack frame
			if (!LineMap.AssemblyLineMaps.Keys.Contains(sf.GetMethod().DeclaringType.Assembly.CodeBase))
				return sl;

			// retrieve the cache
			var alm = LineMap.AssemblyLineMaps[sf.GetMethod().DeclaringType.Assembly.CodeBase];

			// does the symbols list contain the metadata token for this method?
			MemberInfo mi = sf.GetMethod();
			// Don't call this mdtoken or PostSharp will barf on it! Jeez
			long mdtokn = mi.MetadataToken;
			if (!alm.Symbols.ContainsKey(mdtokn))
				return sl;

			// all is good so get the line offset (as close as possible, considering any optimizations that
			// might be in effect)
			var ILOffset = sf.GetILOffset();
			if (ILOffset != StackFrame.OFFSET_UNKNOWN)
			{
				Int64 Addr = alm.Symbols[mdtokn].Address + ILOffset;

				// now start hunting down the line number entry
				// use a simple search. LINQ might make this easier
				// but I'm not sure how. Also, a binary search would be faster
				// but this isn't something that's really performance dependent
				int i = 1;
				for (i = alm.AddressToLineMap.Count - 1; i >= 0; i += -1)
				{
					if (alm.AddressToLineMap[i].Address <= Addr)
					{
						break;
					}
				}
				// since the address may end up between line numbers,
				// always return the line num found
				// even if it's not an exact match
				sl.Line = alm.AddressToLineMap[i].Line;
				sl.SourceFile = alm.Names[alm.AddressToLineMap[i].SourceFileIndex];
			}

			return sl;
		}



		/// <summary>
		/// translate exception object to string, with additional system info
		/// </summary>
		/// <param name="Ex"></param>
		/// <returns></returns>
		public static string ToStringExtended(this Exception ex)
		{
			try
			{
				StringBuilder sb = new StringBuilder();

				if ((ex.InnerException != null))
				{
					// sometimes the original exception is wrapped in a more relevant outer exception
					// the detail exception is the "inner" exception
					// see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/exceptdotnet.asp
					sb.AppendLine("(Inner Exception)");
					sb.AppendLine(ex.InnerException.ToString());
					sb.AppendLine("(Outer Exception)");
				}
				// get general system and app information
				sb.Append(SysInfoToString());

				// get exception-specific information
				sb.AppendLine("Exception Source:      ", () => { return ex.Source; });
				sb.AppendLine("Exception Type:        ", () => { return ex.GetType().FullName; });
				sb.AppendLine("Exception Message:     ", () => { return ex.Message; });
				sb.AppendLine("Exception Target Site: ", () => { return ex.TargetSite.Name; });
				sb.AppendLine(() => { return EnhancedStackTrace(ex); });

				return sb.ToString();

			}
			catch (Exception ex2)
			{
				return string.Format("Error '{0}' while generating exception description", ex2.Message);
			}
		}



		/// <summary>
		/// gather some system information that is helpful to diagnosing
		/// exception
		/// </summary>
		/// <param name="Ex"></param>
		/// <returns></returns>
		private static string SysInfoToString(Exception ex = null)
		{
			StringBuilder sb = new StringBuilder();

			//If Err.Number <> 0 Then
			//   .Append("Error code:            ")
			//   .Append(Err.Number)
			//   .Append(Environment.NewLine)
			//End If

			//If Len(Err.Description) <> 0 Then
			//   .Append("Error Description:     ")
			//   .Append(Err.Description)
			//   .Append(Environment.NewLine)
			//End If

			//'---- report the line or ERL location as available
			//If Err.Line <> 0 Then
			//   .Append("Error Line:            ")
			//   .Append(Err.Line)
			//   If Err.Erl <> 0 AndAlso Err.Erl <> Err.Line Then
			//   .Append("  (Location " & Err.Erl.ToString & ")")
			//   End If
			//   .Append(Environment.NewLine)
			//ElseIf Err.Erl <> 0 Then
			//   .Append("Error Location:        ")
			//   .Append(Err.Erl)
			//   .Append(Environment.NewLine)
			//End If

			//If Err.Column <> 0 Then
			//   .Append("Error Column:          ")
			//   .Append(Err.Column)
			//   .Append(Environment.NewLine)
			//End If

			//If Len(Err.FileName) <> 0 Then
			//   .Append("Error Module:          ")
			//   .Append(Err.FileName)
			//   .Append(Environment.NewLine)
			//End If

			//If Len(Err.Method) <> 0 Then
			//   .Append("Error Method:          ")
			//   .Append(Err.Method)
			//   .Append(Environment.NewLine)
			//End If

			sb.AppendLine("Date and Time:         ", () => { return DateTime.Now.ToString(); });
			sb.AppendLine("Machine Name:          ", () => { return Environment.MachineName; });
			sb.AppendLine("IP Address:            ", () => { return GetCurrentIP(); });
			sb.AppendLine("Current User:          ", () => { return UserIdentity(); });
			sb.AppendLine();
			sb.AppendLine("Application Domain:    ", () => { return System.AppDomain.CurrentDomain.FriendlyName; });
			sb.AppendLine("Assembly Codebase:     ", () => { return ParentAssembly.CodeBase; });
			sb.AppendLine("Assembly Full Name:    ", () => { return ParentAssembly.FullName; });
			sb.AppendLine("Assembly Version:      ", () => { return ParentAssembly.GetName().Version.ToString(); });
			sb.AppendLine("Assembly Build Date:   ", () => { return AssemblyBuildDate(ParentAssembly).ToString(); });
			sb.AppendLine();
			if (ex != null) sb.AppendLine(EnhancedStackTrace(ex));

			return sb.ToString();
		}



		/// <summary>
		/// get IP address of this machine
		/// not an ideal method for a number of reasons (guess why!)
		/// but the alternatives are very ugly
		/// </summary>
		/// <returns></returns>
		private static string GetCurrentIP()
		{
			try
			{
				return System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName()).AddressList[0].ToString();
			}
			catch 
			{
				// just provide a default value
				return "127.0.0.1";
			}
		}



		/// <summary>
		/// retrieve user identity with fallback on error to safer method
		/// </summary>
		/// <returns></returns>
		private static string UserIdentity()
		{
			string strTemp = CurrentWindowsIdentity();
			if (string.IsNullOrEmpty(strTemp))
			{
				strTemp = CurrentEnvironmentIdentity();
			}
			return strTemp;
		}



		/// <summary>
		/// exception-safe WindowsIdentity.GetCurrent retrieval returns "domain\username"
		/// per MS, this sometimes randomly fails with "Access Denied" particularly on NT4
		/// </summary>
		/// <returns></returns>
		private static string CurrentWindowsIdentity()
		{
			try
			{
				return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
			}
			catch
			{
				//just provide a default value
				return string.Empty;
			}
		}


		/// <summary>
		/// exception-safe "domain\username" retrieval from Environment
		/// </summary>
		/// <returns></returns>
		private static string CurrentEnvironmentIdentity()
		{
			try
			{
				return System.Environment.UserDomainName + "\\" + System.Environment.UserName;
			}
			catch 
			{
				//just provide a default value
				return string.Empty;
			}
		}



		/// <summary>
		/// returns build datetime of assembly
		/// assumes default assembly value in AssemblyInfo:
		/// {Assembly: AssemblyVersion("1.0.*")}
		///
		/// filesystem create time is used, if revision and build were overridden by user
		/// </summary>
		/// <param name="objAssembly"></param>
		/// <param name="bForceFileDate"></param>
		/// <returns></returns>
		private static DateTime AssemblyBuildDate(System.Reflection.Assembly objAssembly, bool bForceFileDate = false)
		{
			System.Version objVersion = objAssembly.GetName().Version;
			DateTime dtBuild = default(DateTime);

			if (bForceFileDate)
			{
				dtBuild = AssemblyFileTime(objAssembly);
			}
			else
			{
				dtBuild = ((DateTime.Parse("01/01/2000")).AddDays(objVersion.Build).AddSeconds(objVersion.Revision * 2));
				if (TimeZone.IsDaylightSavingTime(DateTime.Now, TimeZone.CurrentTimeZone.GetDaylightChanges(DateTime.Now.Year)))
				{
					dtBuild = dtBuild.AddHours(1);
				}
				if (dtBuild > DateTime.Now | objVersion.Build < 730 | objVersion.Revision == 0)
				{
					dtBuild = AssemblyFileTime(objAssembly);
				}
			}

			return dtBuild;
		}


		/// <summary>
		/// exception-safe file attrib retrieval; we don't care if this fails
		/// </summary>
		/// <param name="objAssembly"></param>
		/// <returns></returns>
		private static DateTime AssemblyFileTime(System.Reflection.Assembly objAssembly)
		{
			try
			{
				return System.IO.File.GetLastWriteTime(objAssembly.Location);
			}
			catch
			{
				//just provide a default value
				return DateTime.MinValue;
			}
		}


		//--
		//enhanced stack trace generator (exception)
		//--
		private static string InternalEnhancedStackTrace(Exception objException)
		{
			if (objException == null)
			{
				return EnhancedStackTrace(new StackTrace(true), true);
			}
			else
			{
				return EnhancedStackTrace(new StackTrace(objException, true));
			}
		}



		/// <summary>
		/// enhanced stack trace generator (no params)
		/// </summary>
		/// <returns></returns>
		private static string EnhancedStackTrace()
		{
			return EnhancedStackTrace(new StackTrace(true), true);
		}



		/// <summary>
		/// Provides enhanced stack tracing output that includes line numbers, IL Position, and even column
		/// numbers when PDB files are available
		/// </summary>
		/// <param name="thisException"></param>
		/// <returns></returns>
		public static string EnhancedStackTrace(this Exception thisException)
		{
			return InternalEnhancedStackTrace(thisException);
		}

		/// <summary>
		/// enhanced stack trace generator
		/// </summary>
		/// <param name="stackTrace"></param>
		/// <param name="skipLocalFrames"></param>
		/// <returns></returns>
		private static string EnhancedStackTrace(StackTrace stackTrace, bool skipLocalFrames = false)
		{
			StringBuilder sb = new StringBuilder();

			sb.AppendLine();
			sb.Append("---- Stack Trace ----");
			sb.AppendLine();

			for (int intFrame = 0; intFrame <= stackTrace.FrameCount - 1; intFrame++)
			{
				StackFrame sf = stackTrace.GetFrame(intFrame);

				if (skipLocalFrames && sf.GetMethod().DeclaringType.Name.IndexOf(CLASSNAME) > -1)
				{
					// don't include frames related to this class
					// this lets of keep any class frames related to this class out of
					// the strack trace, they'd just be clutter anyway
				}
				else
				{
					sb.Append(StackFrameToString(sf));
				}
			}
			sb.AppendLine();

			return sb.ToString();
		}


		/// <summary>
		/// Used to pass linemap information to the stackframe renderer
		/// </summary>
		private class SourceLine
		{
			public string SourceFile = string.Empty;
			public int Line = 0;
		}
	}
}

namespace InternalExtensionMethods
{
	public static class StringBuilderExtensions
	{
		public static void AppendLine(this StringBuilder sb, string caption, Func<string> getValue)
		{
			sb.Append(caption);
			try
			{
				string text = getValue();
				sb.AppendLine(text);
			}
			catch (Exception ex)
			{
				sb.AppendLine(ex.Message);
			}
		}

		public static void AppendLine(this StringBuilder sb, Func<string> getValue)
		{
			sb.AppendLine(getValue);
		}

	}
}
