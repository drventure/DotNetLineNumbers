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
using System.Reflection;
using System.Text;


public static class ExceptionExtensions
{
	#region Public Members

	/// <summary>
	/// translate exception object to string, with additional system info
	/// </summary>
	/// <param name="Ex"></param>
	/// <returns></returns>
	public static string ToStringExtended(this Exception Ex)
	{
		try
		{
			//-- sometimes the original exception is wrapped in a more relevant outer exception
			//-- the detail exception is the "inner" exception
			//-- see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/exceptdotnet.asp
			string innertemplate = "";
			if ((Ex.InnerException != null))
			{
				innertemplate = Cleanup($@"
					(Inner Exception)
					{Ex.ToStringExtended()} 
					(Outer Exception)
					");
			}

			string template = Cleanup($@"
				{innertemplate}{SysInfoToString()}
				Exception Source:      {Ex.Source} 
				Exception Type:        {Ex.GetType().FullName}
				Exception Message:     {Ex.Message}
				Exception Target Site: {Ex.TargetSite?.Name}
				{EnhancedStackTrace(Ex)}");

			return template;
		}
		catch (Exception ex)
		{
			//there was a problem rendering the extended exception, fall back to generic ToString()
			return Ex.ToString() + "\r\nProblem rendering Exception: " + ex.ToString();
		}
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
		//build method name
		MemberInfo mi = sf.GetMethod();
		string methodName = mi.DeclaringType.Namespace + "." + mi.DeclaringType.Name + "." + mi.Name;

		//build method params
		ParameterInfo[] objParameters = sf.GetMethod().GetParameters();
		int param = 0;
		string parameters = "";
		foreach (ParameterInfo objParameter in objParameters)
		{
			param++;
			if (param > 1) parameters += ", ";
			parameters += objParameter.ParameterType.Name + " " + objParameter.Name;
		}

		string filename = "";
		if (sf.GetFileName() != null && sf.GetFileName().Length != 0)
		{
			//file is available via PDB so use it
			filename = System.IO.Path.GetFileName(sf.GetFileName());
		}
		else if (ParentAssembly != null)
		{
			filename = System.IO.Path.GetFileName(ParentAssembly.CodeBase);
		}
		else
		{
			filename = "Unable to determine Assembly Filename";
		}

		string line = "";
		if (sf.GetFileLineNumber() != 0)
		{
			//line is available via pdb so use it
			line = ": line " + string.Format("{0}", sf.GetFileLineNumber());
		}

		string column = "";
		if (sf.GetFileColumnNumber() != 0)
		{
			//column is available via pdb so use it
			column = ", col " + string.Format("{0:#00}", sf.GetFileColumnNumber());
		}

		string il = "";
		if (sf.GetILOffset() != StackFrame.OFFSET_UNKNOWN)
		{
			//IL is available, append IL location info
			il = ", IL " + string.Format("{0:#0000}", sf.GetILOffset());
		}

		if (filename == "" || line == "")
		{ 
			// Get the native code offset and convert to a line number
			// first, make sure our linemap is loaded
			try
			{
				sf = new MappedStackFrame(sf);
				if (sf.GetFileLineNumber() != 0)
				{
					filename = ": Source File -" + sf.GetFileName();
					line = ": line " + string.Format("{0}", sf.GetFileLineNumber());
				}
			}
			catch (Exception ex)
			{
				//any problems in loading the Linemap, just write to debugger and call it a day
				line = $"Unable to load line map information. Error: {ex.ToString()}";
			}
		}

		return $@"{INDENT}{methodName}({parameters})
			{INDENT}{INDENT}{filename}{line}{column}{il}";
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
	/// gather some system information that is helpful to diagnosing
	/// exception
	/// </summary>
	/// <param name="ex"></param>
	/// <returns></returns>
	private static string SysInfoToString(Exception Ex = null)
	{
		string template = Cleanup($@"
			Date and Time:         {Date}
			Machine Name:          {MachineName}
			IP Address:            {CurrentIP}
			Current User:          {UserIdentity}
			Application Domain:    {CurrentDomain}
			Assembly Codebase:     {CodeBase}
			Assembly Full Name:    {AssemblyName}
			Assembly Version:      {AssemblyVersion}
			Assembly Build Date:   {AssemblyBuildDate}");

		return template;
	}


	/// <summary>
	/// return the current date for logging purposes
	/// </summary>
	private static string Date
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
	private static string CurrentIP
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
	private static string MachineName
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
	private static string UserIdentity
	{
		get
		{
			try
			{
				string ident = CurrentWindowsIdentity;
				if (string.IsNullOrEmpty(ident))
				{
					ident = CurrentEnvironmentIdentity;
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
	private static string CurrentDomain
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
	private static string CodeBase
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
	private static string AssemblyVersion
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
	private static string AssemblyBuildDate
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
	private static string AssemblyName
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
	private static string CurrentWindowsIdentity
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
	private static string CurrentEnvironmentIdentity
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
	/// Cleanup the templated strings.
	/// A bit brute force, but this doesn't have to be performant
	/// </summary>
	/// <param name="buf"></param>
	/// <returns></returns>
	private static string Cleanup(string buf)
	{
		//trim leading and trailing crlfs
		buf = buf.Trim(new char[] { '\r', '\n' });
		//remove tabs (they could end up in the string via the IDE)
		buf = buf.Replace("\t", "");
		//remove leading line spacing (as a result of @"" strings in the IDE)
		buf = buf.Replace("\r\n    ", "\r\n ");
		buf = buf.Replace("\r\n   ", "\r\n ");
		buf = buf.Replace("\r\n  ", "\r\n ");
		//put spaces back for this marker
		buf = buf.Replace("[[3]]", "   ");
		return buf;
	}

	#endregion
}
