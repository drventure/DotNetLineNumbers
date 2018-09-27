#region MIT License
/*
    MIT License

    Copyright (c) 2018 Darin Higgins

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
using System.Security;
using System.Text;


public static class ExceptionExtensions
{
	#region Public Members
	public static bool UsePDB { get; set; }


	/// <summary>
	/// translate exception object to string, with additional system info
	/// </summary>
	/// <param name="Ex"></param>
	/// <returns></returns>
	public static string ToStringExtended(this Exception Ex)
	{
		//-- sometimes the original exception is wrapped in a more relevant outer exception
		//-- the detail exception is the "inner" exception
		//-- see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/exceptdotnet.asp
		string innertemplate = "";
		if ((Ex.InnerException != null))
		{
			innertemplate = Cleanup($@"
				(Inner Exception)
				{render(() => Ex.InnerException.ToString())} 
				(Outer Exception)
				");
		}

		string template = Cleanup($@"
			{innertemplate}{SysInfoToString()}
			Exception Source:      {render(() => Ex.Source)} 
			Exception Type:        {render(() => Ex.GetType().FullName)}
			Exception Message:     {render(() => Ex.Message)}
			Exception Target Site: {render(() => Ex.TargetSite.Name)}
			{render(() => EnhancedStackTrace(Ex))}");

		return template;
	}


	/// <summary>
	/// Provides enhanced stack tracing output that includes line numbers, IL Position, and even column
	/// numbers when PDB files are available
	/// </summary>
	/// <param name="ex"></param>
	/// <returns></returns>
	public static string EnhancedStackTrace(this Exception ex)
	{
		return InternalEnhancedStackTrace(ex);
	}
	#endregion


	#region Private Members
	private const string INDENT = "[[3]]";

	/// <summary>
	/// turns a single stack frame object into an informative string
	/// </summary>
	/// <param name="FrameNum"></param>
	/// <param name="sf"></param>
	/// <returns></returns>
	private static string StackFrameToString(int FrameNum, StackFrame sf)
	{
		StringBuilder sb = new StringBuilder();
		int intParam = 0;
		MemberInfo mi = sf.GetMethod();

		//-- build method name
		sb.Append(INDENT);
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
		sb.Append(INDENT + INDENT);
		if (sf.GetFileName() != null && sf.GetFileName().Length != 0 && ExceptionExtensions.UsePDB)
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
			//-- if IL is available, append IL location info
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
			//If FrameNum = 1 Then rCachedErr.FileName = FileName
			sb.Append(Filename);
			// Get the native code offset and convert to a line number
			// first, make sure our linemap is loaded
			try
			{
				LineMap.AssemblyLineMaps.Add(sf.GetMethod().DeclaringType.Assembly);

				sf = new MappedStackFrame(sf);
				if (sf.GetFileLineNumber() != 0)
				{
					sb.Append(": Source File - ");
					sb.Append(sf.GetFileName());
					sb.Append(": line ");
					sb.Append(string.Format("{0}", sf.GetFileLineNumber()));
				}

			}
			catch (Exception ex)
			{
				//any problems in loading the Linemap, just write to debugger and call it a day
				Debug.WriteLine(string.Format("Unable to load line map information. Error: {0}", ex.ToString()));
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
	/// A remapped stack frame with Line# and source file 
	/// </summary>
	public class MappedStackFrame : StackFrame
	{
		StackFrame _stackFrame = null;
		int _line = 0;
		string _sourceFile = "";

		public MappedStackFrame(StackFrame stackFrame)
		{
			_stackFrame = stackFrame;
			// first, get the base addr of the method
			// if possible

			// you have to have symbols to do this
			if (LineMap.AssemblyLineMaps.Count == 0)
				return;

			// first, check if for symbols for the assembly for this stack frame
			if (!LineMap.AssemblyLineMaps.Keys.Contains(stackFrame.GetMethod().DeclaringType.Assembly.CodeBase))
				return;

			// retrieve the cache
			var alm = LineMap.AssemblyLineMaps[stackFrame.GetMethod().DeclaringType.Assembly.CodeBase];

			// does the symbols list contain the metadata token for this method?
			MemberInfo mi = stackFrame.GetMethod();
			// Don't call this mdtoken or PostSharp will barf on it! Jeez
			long mdtokn = mi.MetadataToken;
			if (!alm.Symbols.ContainsKey(mdtokn))
				return;

			// all is good so get the line offset (as close as possible, considering any optimizations that
			// might be in effect)
			var ILOffset = stackFrame.GetILOffset();
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
				_line = alm.AddressToLineMap[i].Line;
				_sourceFile = alm.Names[alm.AddressToLineMap[i].SourceFileIndex];
				return;
			}
			else
			{
				return;
			}
		}


		public override int GetFileLineNumber()
		{
			return (_line > 0) ? _line : _stackFrame.GetFileLineNumber();
		}
	   	[SecuritySafeCritical]
		public override string GetFileName()
		{
			return (!string.IsNullOrEmpty(_sourceFile) ? _sourceFile : _stackFrame.GetFileName());
		}
		public override int GetFileColumnNumber()
		{
			return _stackFrame.GetFileColumnNumber();
		}
		public override int GetILOffset()
		{
			return _stackFrame.GetILOffset();
		}
		public override MethodBase GetMethod()
		{
			return _stackFrame.GetMethod();
		}
		public override int GetNativeOffset()
		{
			return _stackFrame.GetNativeOffset();
		}
		[SecuritySafeCritical]
		public override string ToString()
		{
			return _stackFrame.ToString();
		}
	}


	/// <summary>
	/// gather some system information that is helpful to diagnosing
	/// exception
	/// </summary>
	/// <param name="ex"></param>
	/// <returns></returns>
	private static string SysInfoToString(Exception Ex = null)
	{
		string template = Cleanup($@"
			Date and Time:         {date}
			Machine Name:          {machineName}
			IP Address:            {currentIP}
			Current User:          {userIdentity}
			Application Domain:    {currentDomain}
			Assembly Codebase:     {codeBase}
			Assembly Full Name:    {assemblyName}
			Assembly Version:      {assemblyVersion}
			Assembly Build Date:   {assemblyBuildDate}");

		return template;
	}


	/// <summary>
	/// return the current date for logging purposes
	/// </summary>
	private static string date
	{
		get
		{
			return DateTime.Now.ToString();
		}
	}


	/// <summary>
	/// get IP address of this machine
	/// not an ideal method for a number of reasons (guess why!)
	/// but the alternatives are very ugly
	/// </summary>
	/// <returns></returns>
	private static string currentIP
		{
			get
			{
				try
				{
					return System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName()).AddressList[0].ToString();
				}
				catch
				{
					//just provide a default value
					return "127.0.0.1";
				}
			}
		}


	/// <summary>
	/// Return the current machine name for logging purposes
	/// </summary>
	private static string machineName
	{
		get
		{
			try
			{
				return Environment.MachineName;
			}
			catch 
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// retrieve user identity with fallback on error to safer method
	/// </summary>
	/// <returns></returns>
	private static string userIdentity
	{
		get
		{
			try
			{
				string ident = currentWindowsIdentity;
				if (string.IsNullOrEmpty(ident))
				{
					ident = currentEnvironmentIdentity;
				}
				return ident;
			}
			catch
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// return the current app domain name for logging
	/// </summary>
	private static string currentDomain
	{
		get
		{
			try
			{
				return System.AppDomain.CurrentDomain.FriendlyName;
			}
			catch 
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// Return the codebase location for logging
	/// </summary>
	private static string codeBase
	{
		get
		{
			try
			{
				return ParentAssembly.CodeBase;
			}
			catch 
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// return the assembly version for logging
	/// </summary>
	private static string assemblyVersion
	{
		get
		{
			try
			{
				return ParentAssembly.GetName().Version.ToString();
			}
			catch 
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// Return the Assembly Build date for logging
	/// </summary>
	private static string assemblyBuildDate
	{
		get
		{
			try
			{
				return GetAssemblyBuildDate(ParentAssembly).ToString();
			}
			catch 
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// return the assemblyName for logging
	/// </summary>
	private static string assemblyName
	{
		get
		{
			try
			{
				return ParentAssembly.FullName;
			}
			catch 
			{
				return "unavailable";
			}
		}
	}


	/// <summary>
	/// exception-safe WindowsIdentity.GetCurrent retrieval returns "domain\username"
	/// per MS, this sometimes randomly fails with "Access Denied" particularly on NT4
	/// </summary>
	/// <returns></returns>
	private static string currentWindowsIdentity
	{
		get
		{
			try
			{
				return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
			}
			catch
			{
				//just provide a default value
				return "";
			}
		}
	}


	/// <summary>
	/// exception-safe "domain\username" retrieval from Environment
	/// </summary>
	/// <returns></returns>
	private static string currentEnvironmentIdentity
	{
		get
		{
			try
			{
				return System.Environment.UserDomainName + "\\" + System.Environment.UserName;
			}
			catch
			{
				//just provide a default value
				return "";
			}
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
	private static DateTime GetAssemblyBuildDate(System.Reflection.Assembly objAssembly, bool bForceFileDate = false)
	{
		System.Version objVersion = objAssembly.GetName().Version;
		DateTime dtBuild = default(DateTime);

		if (bForceFileDate)
		{
			dtBuild = GetAssemblyFileTime(objAssembly);
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
				dtBuild = GetAssemblyFileTime(objAssembly);
			}
		}

		return dtBuild;
	}


	/// <summary>
	/// exception-safe file attrib retrieval; we don't care if this fails
	/// </summary>
	/// <param name="objAssembly"></param>
	/// <returns></returns>
	private static DateTime GetAssemblyFileTime(System.Reflection.Assembly objAssembly)
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


	/// <summary>
	/// enhanced stack trace generator (no params)
	/// </summary>
	/// <returns></returns>
	private static string EnhancedStackTrace()
	{
		//use the type of THIS class and ignore any stackframes in it
		return EnhancedStackTrace(new StackTrace(true), MethodBase.GetCurrentMethod().DeclaringType.Name);
	}


	/// <summary>
	/// enhanced stack trace generator
	/// </summary>
	/// <param name="stackTrace"></param>
	/// <param name="SkipClassNameToSkip"></param>
	/// <returns></returns>
	private static string EnhancedStackTrace(StackTrace stackTrace, string classNameToSkip = "")
	{
		try
		{
			int intFrame = 0;

			StringBuilder sb = new StringBuilder();

			sb.Append("---- Stack Trace ----");
			sb.Append(Environment.NewLine);

			int FrameNum = 0;
			for (intFrame = 0; intFrame <= stackTrace.FrameCount - 1; intFrame++)
			{
				StackFrame sf = stackTrace.GetFrame(intFrame);

				if (classNameToSkip.Length > 0 && sf.GetMethod().DeclaringType.Name.IndexOf(classNameToSkip) > -1)
				{
					// don't include frames with this name
					// this lets of keep any ERR class frames out of
					// the strack trace, they'd just be clutter
				}
				else
				{
					FrameNum += 1;
					sb.Append(StackFrameToString(FrameNum, sf));
				}
			}

			return sb.ToString();
		}
		catch
		{
			return "stack trace unavailable";
		}

	}


	/// <summary>
	/// enhanced stack trace generator (exception)
	/// </summary>
	/// <param name="ex"></param>
	/// <returns></returns>
	private static string InternalEnhancedStackTrace(Exception ex)
	{
		if (ex == null)
		{
			//ignore any stackframes from THIS class
			return EnhancedStackTrace(new StackTrace(true), MethodBase.GetCurrentMethod().DeclaringType.Name);
		}
		else
		{
			return EnhancedStackTrace(new StackTrace(ex, true));
		}
	}


	/// <summary>
	/// Protected rendering of Exception information
	/// </summary>
	/// <param name="render"></param>
	/// <returns></returns>
	private static string render(Func<string> render)
	{
		try
		{
			return render.Invoke();
		}
		catch (Exception ex)
		{
			return ex.Message;
		}
	}


	/// <summary>
	/// Cleanup the templated strings.
	/// A bit brute force, but this doesn't have to be performant
	/// </summary>
	/// <param name="buf"></param>
	/// <returns></returns>
	private static string Cleanup(string buf)
	{
		buf = buf.Trim(new char[] { '\r', '\n' });
		buf = buf.Replace("\t", "");
		buf = buf.Replace("\r\n    ", "\r\n ");
		buf = buf.Replace("\r\n   ", "\r\n ");
		buf = buf.Replace("\r\n  ", "\r\n ");
		buf = buf.Replace("[[3]]", "   ");
		return buf;
	}

	#endregion
}
