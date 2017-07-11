using System;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;


namespace ExceptionExtensions
{
	public static class ExceptionExtensions
	{
		private const string ERRCLASSNAME = "ExceptionExtensions";
		public static bool UsePDB {get; set;}

		/// <summary>
		/// turns a single stack frame object into an informative string
		/// </summary>
		/// <param name="FrameNum"></param>
		/// <param name="sf"></param>
		/// <returns></returns>
		private static string StackFrameToString(int FrameNum, StackFrame sf)
		{
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			int intParam = 0;
			MemberInfo mi = sf.GetMethod();

			//-- build method name
			sb.Append("   ");
			string MethodName = mi.DeclaringType.Namespace + "." + mi.DeclaringType.Name + "." + mi.Name;
			sb.Append(MethodName);
			//If FrameNum = 1 Then rCachedErr.Method = MethodName

			//-- build method params
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
			sb.Append(")");
			sb.Append(Environment.NewLine);

			//-- if source code is available, append location info
			sb.Append("       ");
			if (sf.GetFileName() != null && sf.GetFileName().Length != 0 && ExceptionExtensions.UsePDB)
			{
				//---- the PDB appears to be available, since the above elements are 
				//     not blank, so just use it's information

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
				//-- if IL is available, append IL location info
				if (sf.GetILOffset() != StackFrame.OFFSET_UNKNOWN)
				{
					sb.Append(", IL ");
					sb.Append(string.Format("{0:#0000}", sf.GetILOffset()));
				}
			}
			else
			{
				//---- the PDB is not available, so attempt to retrieve 
				//     any embedded linemap information
				string Filename;
				if (ParentAssembly != null)
				{
					Filename = System.IO.Path.GetFileName(ParentAssembly.CodeBase);
				}
				else
				{
					Filename = "Unable to determine Assembly Filename";
				}
				//If FrameNum = 1 Then rCachedErr.FileName = FileName
				sb.Append(Filename);
				//---- Get the native code offset and convert to a line number
				//     first, make sure our linemap is loaded
				try
				{
					LineMap.AssemblyLineMaps.Add(sf.GetMethod().DeclaringType.Assembly);

					int Line = 0;
					string SourceFile = string.Empty;
					MapStackFrameToSourceLine(sf, ref Line, ref SourceFile);
					if (Line != 0)
					{
						sb.Append(": Source File - ");
						sb.Append(SourceFile);
						sb.Append(": Line ");
						sb.Append(string.Format("{0:#0000}", Line));
					}

				}
				catch (Exception ex)
				{
					//---- just catch any exception here, if we can't load the linemap
					//     oh well, we tried

					//TODO Fixup
					//OnWriteLog(ex, "Unable to load Line Map Resource");
				}
				finally
				{
					//-- native code offset is always available
					var IL = sf.GetILOffset();
					if (IL != StackFrame.OFFSET_UNKNOWN)
					{
						sb.Append(": IL ");
						sb.Append(string.Format("{0:#00000}", IL));
					}
				}
			}
			sb.Append(Environment.NewLine);
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
		private static void MapStackFrameToSourceLine(StackFrame sf, ref int Line, ref string SourceFile)
		{
			//---- first, get the base addr of the method
			//     if possible
			Line = 0;
			SourceFile = string.Empty;

			//---- you have to have symbols to do this
			if (LineMap.AssemblyLineMaps.Count == 0)
				return;

			//---- first, check if for symbols for the assembly for this stack frame
			if (!LineMap.AssemblyLineMaps.Keys.Contains(sf.GetMethod().DeclaringType.Assembly.CodeBase))
				return;

			//---- retrieve the cache
			var alm = LineMap.AssemblyLineMaps[sf.GetMethod().DeclaringType.Assembly.CodeBase];

			//---- does the symbols list contain the metadata token for this method?
			MemberInfo mi = sf.GetMethod();
			//---- Don't call this mdtoken or PostSharp will barf on it! Jeez
			long mdtokn = mi.MetadataToken;
			if (!alm.Symbols.ContainsKey(mdtokn))
				return;

			//---- all is good so get the line offset (as close as possible, considering any optimizations that
			//     might be in effect)
			var ILOffset = sf.GetILOffset();
			if (ILOffset != StackFrame.OFFSET_UNKNOWN)
			{
				Int64 Addr = alm.Symbols[mdtokn].Address + ILOffset;

				//---- now start hunting down the line number entry
				//     use a simple search. LINQ might make this easier
				//     but I'm not sure how. Also, a binary search would be faster
				//     but this isn't something that's really performance dependent
				int i = 1;
				for (i = alm.AddressToLineMap.Count - 1; i >= 0; i += -1)
				{
					if (alm.AddressToLineMap[i].Address <= Addr)
					{
						break;
					}
				}
				//---- since the address may end up between line numbers,
				//     always return the line num found
				//     even if it's not an exact match
				Line = alm.AddressToLineMap[i].Line;
				SourceFile = alm.Names[alm.AddressToLineMap[i].SourceFileIndex];
			}
			else
			{
				return;
			}
		}



		/// <summary>
		/// translate exception object to string, with additional system info
		/// </summary>
		/// <param name="Ex"></param>
		/// <returns></returns>
		public static string ToStringExtended(this Exception Ex)
		{
			try
			{
				StringBuilder sb = new StringBuilder();

				if ((Ex.InnerException != null))
				{
					//-- sometimes the original exception is wrapped in a more relevant outer exception
					//-- the detail exception is the "inner" exception
					//-- see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/exceptdotnet.asp
					sb.Append("(Inner Exception)");
					sb.Append(Environment.NewLine);
					sb.Append(Ex.InnerException.ToString());
					sb.Append(Environment.NewLine);
					sb.Append("(Outer Exception)");
					sb.Append(Environment.NewLine);
				}
				//-- get general system and app information
				sb.Append(SysInfoToString());

				//-- get exception-specific information
				sb.Append("Exception Source:      ");
				try
				{
					sb.Append(Ex.Source);
				}
				catch (Exception e)
				{
					sb.Append(e.Message);
				}
				sb.Append(Environment.NewLine);

				sb.Append("Exception Type:        ");
				try
				{
					sb.Append(Ex.GetType().FullName);
				}
				catch (Exception e)
				{
					sb.Append(e.Message);
				}
				sb.Append(Environment.NewLine);

				sb.Append("Exception Message:     ");
				try
				{
					sb.Append(Ex.Message);
				}
				catch (Exception e)
				{
					sb.Append(e.Message);
				}
				sb.Append(Environment.NewLine);

				sb.Append("Exception Target Site: ");
				try
				{
					sb.Append(Ex.TargetSite.Name);
				}
				catch (Exception e)
				{
					sb.Append(e.Message);
				}
				sb.Append(Environment.NewLine);

				try
				{
					string x = EnhancedStackTrace(Ex);
					sb.Append(x);
				}
				catch (Exception e)
				{
					sb.Append(e.Message);
				}
				sb.Append(Environment.NewLine);

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
		private static string SysInfoToString(Exception Ex = null)
		{
			StringBuilder sb = new StringBuilder();

			//If Err.Number <> 0 Then
			//    .Append("Error code:            ")
			//    .Append(Err.Number)
			//    .Append(Environment.NewLine)
			//End If

			//If Len(Err.Description) <> 0 Then
			//    .Append("Error Description:     ")
			//    .Append(Err.Description)
			//    .Append(Environment.NewLine)
			//End If

			//'---- report the line or ERL location as available
			//If Err.Line <> 0 Then
			//    .Append("Error Line:            ")
			//    .Append(Err.Line)
			//    If Err.Erl <> 0 AndAlso Err.Erl <> Err.Line Then
			//        .Append("  (Location " & Err.Erl.ToString & ")")
			//    End If
			//    .Append(Environment.NewLine)
			//ElseIf Err.Erl <> 0 Then
			//    .Append("Error Location:        ")
			//    .Append(Err.Erl)
			//    .Append(Environment.NewLine)
			//End If

			//If Err.Column <> 0 Then
			//    .Append("Error Column:          ")
			//    .Append(Err.Column)
			//    .Append(Environment.NewLine)
			//End If

			//If Len(Err.FileName) <> 0 Then
			//    .Append("Error Module:          ")
			//    .Append(Err.FileName)
			//    .Append(Environment.NewLine)
			//End If

			//If Len(Err.Method) <> 0 Then
			//    .Append("Error Method:          ")
			//    .Append(Err.Method)
			//    .Append(Environment.NewLine)
			//End If

			sb.Append("Date and Time:         ");
			sb.Append(DateTime.Now);
			sb.Append(Environment.NewLine);

			sb.Append("Machine Name:          ");
			try
			{
				sb.Append(Environment.MachineName);
			}
			catch (Exception e)
			{
				sb.Append(e.Message);
			}
			sb.Append(Environment.NewLine);

			sb.Append("IP Address:            ");
			sb.Append(GetCurrentIP());
			sb.Append(Environment.NewLine);

			sb.Append("Current User:          ");
			sb.Append(UserIdentity());
			sb.Append(Environment.NewLine);
			sb.Append(Environment.NewLine);

			sb.Append("Application Domain:    ");
			try
			{
				sb.Append(System.AppDomain.CurrentDomain.FriendlyName);
			}
			catch (Exception e)
			{
				sb.Append(e.Message);
			}


			sb.Append(Environment.NewLine);
			sb.Append("Assembly Codebase:     ");
			try
			{
				sb.Append(ParentAssembly.CodeBase);
			}
			catch (Exception e)
			{
				sb.Append(e.Message);
			}
			sb.Append(Environment.NewLine);

			sb.Append("Assembly Full Name:    ");
			try
			{
				sb.Append(ParentAssembly.FullName);
			}
			catch (Exception e)
			{
				sb.Append(e.Message);
			}
			sb.Append(Environment.NewLine);

			sb.Append("Assembly Version:      ");
			try
			{
				sb.Append(ParentAssembly.GetName().Version.ToString());
			}
			catch (Exception e)
			{
				sb.Append(e.Message);
			}
			sb.Append(Environment.NewLine);

			sb.Append("Assembly Build Date:   ");
			try
			{
				sb.Append(AssemblyBuildDate(ParentAssembly).ToString());
			}
			catch (Exception e)
			{
				sb.Append(e.Message);
			}
			sb.Append(Environment.NewLine);
			sb.Append(Environment.NewLine);

			if (Ex != null)
			{
				sb.Append(EnhancedStackTrace(Ex));
			}

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
			catch (Exception ex)
			{
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
		///  exception-safe WindowsIdentity.GetCurrent retrieval returns "domain\username"
		///  per MS, this sometimes randomly fails with "Access Denied" particularly on NT4
		/// </summary>
		/// <returns></returns>
		private static string CurrentWindowsIdentity()
		{
			try
			{
				return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
			}
			catch (Exception ex)
			{
				return "";
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
			catch (Exception ex)
			{
				return "";
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
			catch (Exception ex)
			{
				return DateTime.MinValue;
			}
		}


		//--
		//-- enhanced stack trace generator (exception)
		//--
		private static string InternalEnhancedStackTrace(Exception objException)
		{
			if (objException == null)
			{
				return EnhancedStackTrace(new StackTrace(true), ERRCLASSNAME);
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
			return EnhancedStackTrace(new StackTrace(true), ERRCLASSNAME);
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
		/// <param name="objStackTrace"></param>
		/// <param name="SkipClassNameToSkip"></param>
		/// <returns></returns>
		private static string EnhancedStackTrace(StackTrace objStackTrace, string SkipClassNameToSkip = "")
		{

			int intFrame = 0;

			System.Text.StringBuilder sb = new System.Text.StringBuilder();

			sb.Append(Environment.NewLine);
			sb.Append("---- Stack Trace ----");
			sb.Append(Environment.NewLine);

			int FrameNum = 0;
			for (intFrame = 0; intFrame <= objStackTrace.FrameCount - 1; intFrame++)
			{
				StackFrame sf = objStackTrace.GetFrame(intFrame);

				if (SkipClassNameToSkip.Length > 0 && sf.GetMethod().DeclaringType.Name.IndexOf(SkipClassNameToSkip) > -1)
				{
					//---- don't include frames with this name
					//     this lets of keep any ERR class frames out of
					//     the strack trace, they'd just be clutter
				}
				else
				{
					FrameNum += 1;
					sb.Append(StackFrameToString(FrameNum, sf));
				}
			}
			sb.Append(Environment.NewLine);

			return sb.ToString();
		}

	}
}
