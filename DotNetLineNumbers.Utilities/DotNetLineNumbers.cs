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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Xml;


namespace DotNetLineNumbers
{
	/// <summary>
	/// Root name space for the DotNetLineNumbers classes and utilities for resolving
	/// line numbers from line map resources embedded within .net executables.
	/// </summary>
	internal class NamespaceDoc
	{ }


	#region " Enhanced Stack Classes"
	/// <summary>
	/// A StackFrame subclass that includes logic to resolve
	/// a line number from an existing line map (either a resource or a separate line map file)
	/// If no line map resource can be found, this automatically falls back to using
	/// whatever information is available via a normal StackFrame
	/// <code>
	/// var sf = new EnhancedStackFrame(ExistingStackFrame);
	/// var ln = sf.GetLineNumber();
	/// </code>
	/// </summary>
	public class EnhancedStackFrame : StackFrame
	{
		readonly StackFrame _stackFrame = null;
		readonly int _line = 0;
		readonly string _sourceFile = string.Empty;

		/// <summary>
		/// Retrieve an EnhancedStackFrame from the passed in StackFrame object
		/// </summary>
		/// <param name="stackFrame">A normal StackFrame object to enhance.</param>
		public EnhancedStackFrame(StackFrame stackFrame)
		{
			//attempt to load any line maps from the source assembly
			Internals.Bookkeeping.AssemblyLineMaps.Add(stackFrame.GetMethod().DeclaringType.Assembly);

			_stackFrame = stackFrame;
			// first, get the base address of the method
			// if possible. We have to have symbols to do this...
			if (Internals.Bookkeeping.AssemblyLineMaps.Count == 0)
				return;

			// first, check if for symbols for the assembly for this stack frame
			if (!Internals.Bookkeeping.AssemblyLineMaps.Keys.Contains(stackFrame.GetMethod().DeclaringType.Assembly.CodeBase))
				return;

			// retrieve the cache
			var alm = Internals.Bookkeeping.AssemblyLineMaps[stackFrame.GetMethod().DeclaringType.Assembly.CodeBase];

			// does the symbols list contain the meta data token for this method?
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
				// always return the line number found
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


		/// <summary>
		/// Retrieve the line number recorded for this stack frame
		/// </summary>
		/// <returns>An integer line number in the source file that this stack frame points to.</returns>
		public override int GetFileLineNumber()
		{
			return (_line > 0) ? _line : _stackFrame.GetFileLineNumber();
		}


		/// <summary>
		/// Retrieve the filename recorded for this stack frame
		/// </summary>
		/// <returns>A string containing the filename of the source file this stack frame points to.</returns>
		[SecuritySafeCritical]
		public override string GetFileName()
		{
			return (!string.IsNullOrEmpty(_sourceFile) ? _sourceFile : _stackFrame.GetFileName());
		}


		/// <summary>
		/// Retrieve the column number recorded for this stack frame
		/// </summary>
		/// <returns>An integer containing the column number of the source file this stack frame points to.</returns>
		public override int GetFileColumnNumber()
		{
			return _stackFrame.GetFileColumnNumber();
		}


		/// <summary>
		/// Retrieve the Intermediate Language (IL) offset recorded for this stack frame
		/// </summary>
		/// <returns>An integer containing the IL offset of the source file this stack frame points to.</returns>
		public override int GetILOffset()
		{
			return _stackFrame.GetILOffset();
		}


		/// <summary>
		/// Retrieve the MethodBase recorded for this stack frame
		/// </summary>
		/// <returns>The MethodBase object that this stack frame points to.</returns>
		public override MethodBase GetMethod()
		{
			return _stackFrame.GetMethod();
		}


		/// <summary>
		/// Retrieve the native code offset recorded for this stack frame
		/// </summary>
		/// <returns>An integer containing the native code offset that this stack frame points to.</returns>
		public override int GetNativeOffset()
		{
			return _stackFrame.GetNativeOffset();
		}


		/// <summary>
		/// Generate a human readable string description of this stack frame.
		/// </summary>
		/// <returns>The string description of this stack frame.</returns>
		[SecuritySafeCritical]
		public override string ToString()
		{
			//build method name
			MemberInfo mi = this.GetMethod();
			string methodName = mi.DeclaringType.Namespace + "." + mi.DeclaringType.Name + "." + mi.Name;

			//build method parameters
			ParameterInfo[] objParameters = this.GetMethod().GetParameters();
			int param = 0;
			string parameters = "";
			foreach (ParameterInfo objParameter in objParameters)
			{
				param++;
				if (param > 1) parameters += ", ";
				parameters += objParameter.ParameterType.Name + " " + objParameter.Name;
			}

			string filename = "";
			if (this.GetFileName() != null && this.GetFileName().Length != 0)
			{
				//file is available via PDB so use it
				filename = System.IO.Path.GetFileName(this.GetFileName());
			}
			else if (Internals.ParentAssembly != null)
			{
				filename = System.IO.Path.GetFileName(Internals.ParentAssembly.CodeBase);
			}
			else
			{
				filename = "Unable to determine Assembly Filename";
			}

			string line = "";
			if (this.GetFileLineNumber() != 0)
			{
				//line is available via pdb so use it
				line = ": line " + string.Format("{0}", this.GetFileLineNumber());
			}

			string column = "";
			if (this.GetFileColumnNumber() != 0)
			{
				//column is available via pdb so use it
				column = ", col " + string.Format("{0:#00}", this.GetFileColumnNumber());
			}

			string il = "";
			if (this.GetILOffset() != StackFrame.OFFSET_UNKNOWN)
			{
				//IL is available, append IL location info
				il = ", IL " + string.Format("{0:#0000}", this.GetILOffset());
			}

			if (filename == "" || line == "")
			{
				// Get the native code offset and convert to a line number
				// first, make sure our line map is loaded
				try
				{
					if (this.GetFileLineNumber() != 0)
					{
						filename = "Source File -" + this.GetFileName();
						line = ": line " + string.Format("{0}", this.GetFileLineNumber());
					}
				}
				catch (Exception ex)
				{
					//any problems in loading the Line map, just write to debugger and call it a day
					line = $"Unable to load line map information. Error: {ex.ToString()}";
				}
			}

			return $@"{Internals.Constants.INDENT}{methodName}({parameters}) :
			{Internals.Constants.INDENT}{Internals.Constants.INDENT}{filename}{line}{column}{il}";
		}
	}


	/// <summary>
	/// A StackTrace subclass that includes logic to resolve
	/// a line number from an existing line map (either a resource or a separate line map file)
	/// for all contained stack frames.
	/// If no line map resource can be found, this automatically falls back to using
	/// whatever information is available via a normal StackFrame
	/// <code>
	/// var st = new EnhancedStackTrace();
	/// var ln = st.GetFrame(0).GetLineNumber();
	/// </code>
	/// </summary>
	public class EnhancedStackTrace : StackTrace
	{
		private EnhancedStackFrame[] _frames = null;
		private HashSet<string> _skipClassNames = null;

		#region " Constructors"
		/// <summary>
		/// Initializes a new instance of an EnhancedStackTrace class from the
		/// caller's frame.
		/// </summary>
		public EnhancedStackTrace() : base()
		{ }


		/// <summary>
		/// Initializes a new instance of an EnhancedStackTrace class from the
		/// caller's frame, skipping any frame pointed to the class with the passed class name.
		/// </summary>
		/// <param name="skipClassName">A string containing the class name used to skip frames.</param>
		public EnhancedStackTrace(string skipClassName) : base()
		{
			_skipClassNames = new HashSet<string>(new string[] { skipClassName });
		}


		/// <summary>
		/// Initializes a new instance of an EnhancedStackTrace class from the
		/// caller's frame, skipping any frame pointed to any of the named classes.
		/// </summary>
		/// <param name="skipClassNames">An array or list of class names used to skip frames.</param>
		public EnhancedStackTrace(IEnumerable<string> skipClassNames) : base()
		{
			_skipClassNames = new HashSet<string>(skipClassNames);
		}


		/// <summary>
		///  Initializes a new instance of the System.Diagnostics.StackTrace class from the
		///  caller's frame, optionally capturing source information.
		/// </summary>
		/// <param name="fNeedFileInfo">True to capture the file name, line number, and column number; otherwise, false.</param>
		public EnhancedStackTrace(bool fNeedFileInfo) : base(fNeedFileInfo)
		{ }


		/// <summary>
		///  Initializes a new instance of the System.Diagnostics.StackTrace class from the
		///  caller's frame, optionally capturing source information.
		/// </summary>
		/// <param name="fNeedFileInfo">True to capture the file name, line number, and column number; otherwise, false.</param>
		/// <param name="skipClassName">A string containing the class name used to skip frames.</param>
		public EnhancedStackTrace(bool fNeedFileInfo, string skipClassName) : base(fNeedFileInfo)
		{
			_skipClassNames = new HashSet<string>(new string[] { skipClassName });
		}


		/// <summary>
		///  Initializes a new instance of the System.Diagnostics.StackTrace class from the
		///  caller's frame, optionally capturing source information.
		/// </summary>
		/// <param name="fNeedFileInfo">True to capture the file name, line number, and column number; otherwise, false.</param>
		/// <param name="skipClassNames">An array or list of class names used to skip frames.</param>
		public EnhancedStackTrace(bool fNeedFileInfo, IEnumerable<string> skipClassNames) : base(fNeedFileInfo)
		{
			_skipClassNames = new HashSet<string>(skipClassNames);
		}


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class from the
		/// caller's frame, skipping the specified number of frames.
		/// </summary>
		/// <param name="skipFrames">The number of frames up the stack from which to start the trace.</param>
		/// <exception cref="System.ArgumentOutOfRangeException">The skipFrames parameter is negative.</exception>
		public EnhancedStackTrace(int skipFrames) : base(skipFrames)
		{ }


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class using the
		/// provided exception object.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		public EnhancedStackTrace(Exception e) : base(e)
		{ }
				   

		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class using the
		/// provided exception object, skipping frames of the given named class.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="skipClassName">A string containing the class name used to skip frames.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		public EnhancedStackTrace(Exception e, string skipClassName) : base(e)
		{
			_skipClassNames = new HashSet<string>(new string[] { skipClassName });
		}


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class using the
		/// provided exception object, skipping frames of the given named classes.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="skipClassNames">An array or list of class names used to skip frames.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		public EnhancedStackTrace(Exception e, IEnumerable<string> skipClassNames) : base(e)
		{
			_skipClassNames = new HashSet<string>(skipClassNames);
		}


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class that contains
		/// a single frame.
		/// </summary>
		/// <param name="frame">The frame that the System.Diagnostics.StackTrace object should contain.</param>
		public EnhancedStackTrace(EnhancedStackFrame frame) : base(frame)
		{ }


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class from the
		/// caller's frame, skipping the specified number of frames and optionally capturing
		/// source information.
		/// </summary>
		/// <param name="skipFrames">The number of frames up the stack from which to start the trace.</param>
		/// <param name="fNeedFileInfo">true to capture the file name, line number, and column number; otherwise, false.</param>
		/// <exception cref="System.ArgumentOutOfRangeException">The skipFrames parameter is negative.</exception>
		public EnhancedStackTrace(int skipFrames, bool fNeedFileInfo) : base(skipFrames, fNeedFileInfo)
		{ }


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class, using
		/// the provided exception object and optionally capturing source information.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="fNeedFileInfo">true to capture the file name, line number, and column number; otherwise, false.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		public EnhancedStackTrace(Exception e, bool fNeedFileInfo) : base(e, fNeedFileInfo)
		{ }


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class, using
		/// the provided exception object and optionally capturing source information.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="fNeedFileInfo">true to capture the file name, line number, and column number; otherwise, false.</param>
		/// <param name="skipClassName">A string containing the class name used to skip frames.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		public EnhancedStackTrace(Exception e, bool fNeedFileInfo, string skipClassName) : base(e, fNeedFileInfo)
		{
			_skipClassNames = new HashSet<string>(new string[] { skipClassName });
		}


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class, using
		/// the provided exception object and optionally capturing source information.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="fNeedFileInfo">true to capture the file name, line number, and column number; otherwise, false.</param>
		/// <param name="skipClassNames">An array or list of class names used to skip frames.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		public EnhancedStackTrace(Exception e, bool fNeedFileInfo, IEnumerable<string> skipClassNames) : base(e, fNeedFileInfo)
		{
			_skipClassNames = new HashSet<string>(skipClassNames);
		}


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class using the
		/// provided exception object and skipping the specified number of frames.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="skipFrames">The number of frames up the stack from which to start the trace.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		/// <exception cref="System.ArgumentOutOfRangeException">The skipFrames parameter is negative.</exception>
		public EnhancedStackTrace(Exception e, int skipFrames) : base(e, skipFrames)
		{ }


		/// <summary>
		/// Initializes a new instance of the System.Diagnostics.StackTrace class using the
		/// provided exception object, skipping the specified number of frames and optionally
		/// capturing source information.
		/// </summary>
		/// <param name="e">The Exception from which to generate this stack trace.</param>
		/// <param name="skipFrames">The number of frames up the stack from which to start the trace.</param>
		/// <param name="fNeedFileInfo">true to capture the file name, line number, and column number; otherwise, false.</param>
		/// <exception cref="System.ArgumentNullException">The e parameter is null.</exception>
		/// <exception cref="System.ArgumentOutOfRangeException">The skipFrames parameter is negative.</exception>
		public EnhancedStackTrace(Exception e, int skipFrames, bool fNeedFileInfo) : base(e, skipFrames, fNeedFileInfo)
		{ }
		#endregion


		/// <summary>
		/// Retrieve a specific frame from the collection of frames that make up the trace.
		/// </summary>
		/// <param name="index">0-based index of the frame to retrieve.</param>
		/// <returns>An EnhancedStackFrame object</returns>
		public new EnhancedStackFrame GetFrame(int index)
		{
			if (_frames == null) return null;
			if (index >= 0 && index < _frames.Length) return _frames[index];
			return null;
		}


		/// <summary>
		/// Override the FrameCount property so that it takes into account skipped frames as well.
		/// </summary>
		public new int FrameCount
		{
			get
			{
				return this.GetFrames().Length;
			}
		}


		/// <summary>
		/// Returns a copy of all stack frames in the current stack trace.
		/// </summary>
		/// <returns>
		/// An array of type System.Diagnostics.StackFrame representing the function calls in the stack trace.
		/// </returns>
		[ComVisible(false)]
		public new EnhancedStackFrame[] GetFrames()
		{
			if (_frames == null)
			{
				var list = new List<EnhancedStackFrame>();

				for (var frameIdx = 0; frameIdx <= base.FrameCount - 1; frameIdx++)
				{
					var frame = new EnhancedStackFrame(base.GetFrame(frameIdx));

					if (_skipClassNames != null)
					{
						var className = frame.GetMethod().DeclaringType.Name;
						if (!_skipClassNames.Contains(className))
						{
							list.Add(frame);
						}
					}
					else
					{
						list.Add(frame);
					}
				}
				_frames = list.ToArray();
			}
			return _frames;
		}


		/// <summary>
		/// Convert this Stack Trace into a readable string representation.
		/// </summary>
		/// <returns>A string representation of this stack trace.</returns>
		public override string ToString()
		{
			try
			{
				int intFrame = 0;

				StringBuilder sb = new StringBuilder();

				sb.Append("---- Stack Trace ----");
				sb.Append(Environment.NewLine);

				for (intFrame = 0; intFrame <= this.FrameCount - 1; intFrame++)
				{
					EnhancedStackFrame sf = this.GetFrame(intFrame);
					sb.Append(sf.ToString());
				}

				return sb.ToString();
			}
			catch
			{
				return "stack trace unavailable";
			}
		}
	}
	#endregion


	#region " Exception Extension Methods"
	/// <summary>
	/// Static class that contains extension methods to the Exception base class
	/// for assistance with resolving line numbers using line map resources.
	/// </summary>
	public static class ExceptionExtensions
	{
		#region Public Members

		/// <summary>
		/// translate exception object to string, with additional system info
		/// </summary>
		/// <param name="ex">An Exception object to resolve into a detailed stack trace string representation.</param>
		/// <returns>A detailed string rendering of the Exception, with stack trace information and line numbers where possible.</returns>
		public static string ToStringEnhanced(this Exception ex)
		{
			return CleanupAndIndent(ex.ToStringEnhancedInternal());
		}


		/// <summary>
		/// translate exception object to string, with additional system info
		/// </summary>
		/// <param name="Ex"></param>
		/// <returns></returns>
		private static string ToStringEnhancedInternal(this Exception ex)
		{
			try
			{
				//-- sometimes the original exception is wrapped in a more relevant outer exception
				//-- the detail exception is the "inner" exception
				//-- see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnbda/html/exceptdotnet.asp
				string innertemplate = "";
				if ((ex.InnerException != null))
				{
					innertemplate = Cleanup($@"
					(Inner Exception)
					{ex.InnerException.ToStringEnhancedInternal()} 
					(Outer Exception)
					");
				}

				var msg = Cleanup($@"{ex.ToString()}").Replace("\r\n", $"\r\n{" ",24}");

				string template = Cleanup($@"
				{innertemplate}{SysInfoToString()}
				Exception Source:       {ex.Source} 
				Exception Type:         {ex.GetType().FullName}
				Exception Message:      {msg}
				Exception Target Site:  {ex.TargetSite?.Name}
				{ex.GetEnhancedStackTrace().ToString()}");

				return template;
			}
			catch (Exception localEx)
			{
				//there was a problem rendering the extended exception, fall back to generic ToString()
				return ex.ToString() + "\r\nProblem rendering Exception: " + localEx.ToString();
			}
		}
		#endregion


		#region Private Members

		/// <summary>
		/// gather some system information that is helpful to diagnosing
		/// exception
		/// </summary>
		/// <param name="ex"></param>
		/// <returns></returns>
		private static string SysInfoToString(Exception Ex = null)
		{
			string template = Cleanup($@"
			Date and Time:          {Date}
			Machine Name:           {MachineName}
			IP Address:             {CurrentIP}
			Current User:           {UserIdentity}
			Application Domain:     {CurrentDomain}
			Assembly Codebase:      {CodeBase}
			Assembly Full Name:     {AssemblyName}
			Assembly Version:       {AssemblyVersion}
			Assembly Build Date:    {AssemblyBuildDate}");

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
		/// retrieve user identity with fall back on error to safer method
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
					return Internals.ParentAssembly.CodeBase;
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
					return Internals.ParentAssembly.GetName().Version.ToString();
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
					return GetAssemblyBuildDate(Internals.ParentAssembly).ToString();
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
					return Internals.ParentAssembly.FullName;
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
		/// returns build date time of assembly
		/// assumes default assembly value in AssemblyInfo:
		/// {Assembly: AssemblyVersion("1.0.*")}
		///
		/// file system create time is used, if revision and build were overridden by user
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
		/// exception-safe file attribute retrieval; we don't care if this fails
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
		/// enhanced stack trace generator
		/// </summary>
		/// <param name="ex">An Exception object to generate a stack trace from.</param>
		/// <returns>An EnhancedStackTrace object based on the stack trace from the given exception.</returns>
		public static EnhancedStackTrace GetEnhancedStackTrace(this Exception ex)
		{
			if (ex == null)
			{
				//ignore any stack frames from THIS class
				return (new EnhancedStackTrace(true, MethodBase.GetCurrentMethod().DeclaringType.Name));
			}
			else
			{
				return (new EnhancedStackTrace(ex, true));
			}
		}


		/// <summary>
		/// Cleanup the templated string.
		/// A bit brute force, but this doesn't have to be performant
		/// </summary>
		/// <param name="buf"></param>
		/// <returns></returns>
		private static string Cleanup(string buf)
		{
			var i = buf.IndexOf("\r\n   at ");
			if (i > 0) buf = buf.Substring(0, i - 1);

			//trim leading and trailing crlfs
			buf = buf.Trim(new char[] { '\r', '\n' });
			//remove tabs (they could end up in the string via the IDE)
			buf = buf.Replace("\t", "");
			//remove leading line spacing (as a result of @"" strings in the IDE)
			buf = buf.Replace("\r\n    ", "\r\n ");
			buf = buf.Replace("\r\n   ", "\r\n ");
			buf = buf.Replace("\r\n  ", "\r\n ");
			return buf;
		}


		/// <summary>
		/// Cleanup and indent the templated string.
		/// A bit brute force, but this doesn't have to be performant
		/// </summary>
		/// <param name="buf"></param>
		/// <returns></returns>
		private static string CleanupAndIndent(string buf)
		{
			//cleanup and then put spaces back in for markers
			return Cleanup(buf).Replace($"{Internals.Constants.INDENT}", "   ");
		}

		#endregion
	}
	#endregion


	#region " Internal classes, and definitions for shared use by GenerateLineMap"
	/// <summary>
	/// Internal Utility classes and definitions used by DotNetLineNumbers.cs
	/// and the GenerateLineMap utility.
	/// </summary>
	/// <exclude />
	public static class Internals
	{
		/// <summary>
		/// Retrieve the root assembly of the executing assembly
		/// </summary>
		/// <returns></returns>
		public static Assembly ParentAssembly
		{
			get
			{
				if (Assembly.GetEntryAssembly() != null)
				{
					return Assembly.GetEntryAssembly();
				}
				else if (Assembly.GetCallingAssembly() != null)
				{
					return Assembly.GetCallingAssembly();
				}
				else
				{
					//TODO questionable
					return Assembly.GetExecutingAssembly();
				}
			}
		}


		/// <summary>
		/// Internal static class to contain definitions for the AssemblyLineMap objects
		/// and keep them out of the general DotNetLineNumbers name space.
		/// </summary>
		public static class Bookkeeping
		{
			/// <summary>
			/// Define a collection of assembly line maps (one for each assembly we
			/// need to generate a stack trace through)
			/// </summary>
			/// <remarks></remarks>
			/// <editHistory></editHistory>
			public class AssemblyLineMapCollection : Dictionary<string, AssemblyLineMap>
			{

				/// <summary>
				/// Load a line map file given the NAME of an assembly
				/// obviously, the Assembly must exist.
				/// </summary>
				/// <param name="FileName"></param>
				/// <remarks></remarks>
				public AssemblyLineMap Add(string FileName)
				{
					if (!File.Exists(FileName))
					{
						throw new FileNotFoundException("The file could not be found.", FileName);
					}

					if (this.ContainsKey(FileName))
					{
						// no need, already loaded (should it reload?)
						return this[FileName];
					}
					else
					{
						var alm = new AssemblyLineMap(FileName);
						this.Add(FileName, alm);
						return alm;
					}
				}


				public AssemblyLineMap Add(Assembly Assembly)
				{
					var FileName = Assembly.CodeBase;
					if (this.ContainsKey(FileName))
					{
						// no need, already loaded (should it reload?)
						return this[FileName];
					}
					else
					{
						var alm = new AssemblyLineMap(Assembly);
						this.Add(FileName, alm);
						return alm;
					}
				}
			}
			public static AssemblyLineMapCollection AssemblyLineMaps = new AssemblyLineMapCollection();
		}


		/// <summary>
		/// Tracks symbol cache entries for all assemblies that appear in a stack frame
		/// Public for serialization purposes
		/// </summary>
		/// <remarks></remarks>
		/// <editHistory></editHistory>
		[DataContract(Namespace = Constants.LINEMAPNAMESPACE)]
		public class AssemblyLineMap
		{
			/// <summary>
			/// Track the assembly this map is for
			/// no need to persist this information
			/// </summary>
			/// <remarks></remarks>
			public string FileName;


			/// <summary>
			/// Track each Address to Line mapping
			/// Public for serialization purposes
			/// </summary>
			[DataContract(Namespace = Constants.LINEMAPNAMESPACE)]
			public class AddressToLine
			{
				// these members must get serialized
				[DataMember()]
				public Int64 Address;
				[DataMember()]
				public Int32 Line;
				[DataMember()]
				public int SourceFileIndex;
				[DataMember()]

				public int ObjectNameIndex;
				// these members do not need to be serialized
				public string SourceFile;

				public string ObjectName;

				/// <summary>
				/// Parameterless constructor for serialization
				/// </summary>
				/// <remarks></remarks>
				public AddressToLine() { }


				public AddressToLine(Int32 Line, Int64 Address, string SourceFile, int SourceFileIndex, string ObjectName, int ObjectNameIndex)
				{
					this.Line = Line;
					this.Address = Address;
					this.SourceFile = SourceFile;
					this.SourceFileIndex = SourceFileIndex;
					this.ObjectName = ObjectName;
					this.ObjectNameIndex = ObjectNameIndex;
				}
			}


			/// <summary>
			/// Track the Line number list enumerated from the PDB
			/// Note, the list is already sorted
			/// </summary>
			/// <remarks></remarks>
			[DataMember()]
			public List<AddressToLine> AddressToLineMap = new List<AddressToLine>();


			/// <summary>
			/// Private class to track Symbols read from the PDB file
			/// </summary>
			/// <remarks></remarks>
			/// <editHistory></editHistory>
			[DataContract(Namespace = Constants.LINEMAPNAMESPACE)]
			public class SymbolInfo
			{
				// these need to be persisted
				[DataMember()]
				public long Address;
				[DataMember()]

				public long Token;
				// these aren't persisted

				public string Name;

				/// <summary>
				/// Parameterless constructor for serialization
				/// </summary>
				/// <remarks></remarks>
				public SymbolInfo() { }


				public SymbolInfo(string Name, long Address, long Token)
				{
					this.Name = Name;
					this.Address = Address;
					this.Token = Token;
				}
			}
			/// <summary>
			/// Track the Symbols enumerated from the PDB keyed by their token
			/// </summary>
			/// <remarks></remarks>
			[DataMember()]
			public Dictionary<long, SymbolInfo> Symbols = new Dictionary<long, SymbolInfo>();

			/// <summary>
			/// Track a list of string values
			/// </summary>
			/// <remarks></remarks>
			/// <editHistory></editHistory>
			public class NamesList : List<string>
			{
				/// <summary>
				/// When adding names, if the name already exists in the list
				/// don't bother to add it again
				/// </summary>
				/// <param name="Name"></param>
				/// <returns></returns>
				public new int Add(string Name)
				{
					//Don't lower case everything
					//Name = Name.ToLower();
					var i = this.IndexOf(Name);
					if (i >= 0)
						return i;

					// gotta add the name
					base.Add(Name);
					return this.Count - 1;
				}


				/// <summary>
				/// Override this prop so that requesting an index that doesn't exist just
				/// returns a blank string
				/// </summary>
				/// <param name="Index"></param>
				/// <value></value>
				/// <remarks></remarks>
				public new string this[int Index]
				{
					get
					{
						if (Index >= 0 & Index < this.Count)
						{
							return base[Index];
						}
						else
						{
							return string.Empty;
						}
					}
					set
					{
						base[Index] = value;
					}
				}
			}


			/// <summary>
			/// Tracks various string values in a flat list that is indexed into
			/// </summary>
			/// <remarks></remarks>
			[DataMember()]
			public NamesList Names = new NamesList();


			/// <summary>
			/// Create a new map based on an assembly filename
			/// </summary>
			/// <param name="FileName"></param>
			/// <remarks></remarks>
			public AssemblyLineMap(string FileName)
			{
				this.FileName = FileName;
				Load();
			}


			/// <summary>
			/// Parameterless constructor for serialization
			/// </summary>
			/// <remarks></remarks>

			public AssemblyLineMap() { }


			/// <summary>
			/// Clear out all internal information
			/// </summary>
			/// <remarks></remarks>
			public void Clear()
			{
				this.Symbols.Clear();
				this.AddressToLineMap.Clear();
				this.Names.Clear();
			}


			/// <summary>
			/// Load a LINEMAP file
			/// </summary>
			/// <remarks></remarks>
			private void Load()
			{
				Stream buf = FileToStream(this.FileName + ".lmp");
				var buf2 = DecryptStream(buf);
				Transfer(Depersist(DecompressStream(buf2)));
			}


			/// <summary>
			/// Create a new assembly line map based on an assembly
			/// </summary>
			/// <remarks></remarks>
			public AssemblyLineMap(Assembly Assembly)
			{
				this.FileName = Assembly.CodeBase.Replace("file:///", string.Empty);

				try
				{
					// Get the hInstance of the indicated exe/dll image
					var curmodule = Assembly.GetLoadedModules()[0];
					var hInst = System.Runtime.InteropServices.Marshal.GetHINSTANCE(curmodule);

					// retrieve a handle to the Line map resource
					// Since it's a standard Win32 resource, the nice .NET resource functions
					// can't be used
					//
					// Important Note: The FindResourceEx function appears to be case
					// sensitive in that you really HAVE to pass in UPPER CASE search
					// arguments
					var hres = WindowsAPI.FindResourceEx(hInst.ToInt32(), Constants.ResTypeName, Constants.ResName, Constants.ResLang);

					byte[] bytes = null;
					if (hres != IntPtr.Zero)
					{
						// Load the resource to get it into memory
						var hresdata = WindowsAPI.LoadResource(hInst, hres);

						IntPtr lpdata = WindowsAPI.LockResource(hresdata);
						var sz = WindowsAPI.SizeofResource(hInst, hres);

						if (lpdata != IntPtr.Zero & sz > 0)
						{
							// able to lock it,
							// so copy the data into a byte array
							bytes = new byte[sz];
							WindowsAPI.CopyMemory(ref bytes[0], lpdata, sz);
							WindowsAPI.FreeResource(hresdata);
						}
					}
					else
					{
						//Check for a side by side line map file
						string mapfile = this.FileName + ".lmp";
						if (File.Exists(mapfile))
						{
							//load it from there
							bytes = File.ReadAllBytes(mapfile);
						}
					}

					if (bytes != null)
					{
						// deserialize the symbol map and line number list
						using (System.IO.MemoryStream MemStream = new MemoryStream(bytes))
						{
							// release the byte array to free up the memory
							bytes = null;
							// and depersist the object
							Stream temp = MemStream;
							var temp2 = DecryptStream(temp);
							Transfer(Depersist(DecompressStream(temp2)));
						}
					}

				}
				catch (Exception ex)
				{
					Debug.WriteLine(string.Format("ERROR: {0}", ex.ToString()));
					// yes, it's bad form to catch all exceptions like this
					// but this is part of an exception handler
					// so it really can't be allowed to fail with an exception!
				}

				try
				{
					if (this.Symbols.Count == 0)
					{
						// weren't able to load resources, so try the LINEMAP (LMP) file
						Load();
					}

				}
				catch (Exception ex)
				{
					Debug.WriteLine(string.Format("ERROR: {0}", ex.ToString()));
				}
			}


			/// <summary>
			/// Transfer a given AssemblyLineMap's contents to this one
			/// </summary>
			/// <param name="alm"></param>
			/// <remarks></remarks>
			private void Transfer(AssemblyLineMap alm)
			{
				// transfer Internal variables over
				this.AddressToLineMap = alm.AddressToLineMap;
				this.Symbols = alm.Symbols;
				this.Names = alm.Names;
			}


			/// <summary>
			/// Read an entire file into a memory stream
			/// </summary>
			/// <param name="Filename"></param>
			/// <returns></returns>
			private MemoryStream FileToStream(string Filename)
			{
				if (File.Exists(Filename))
				{
					using (System.IO.FileStream FileStream = new System.IO.FileStream(Filename, FileMode.Open))
					{
						if (FileStream.Length > 0)
						{
							byte[] Buffer = null;
							Buffer = new byte[Convert.ToInt32(FileStream.Length - 1) + 1];
							FileStream.Read(Buffer, 0, Convert.ToInt32(FileStream.Length));
							return new MemoryStream(Buffer);
						}
					}
				}

				// just return an empty stream
				return new MemoryStream();
			}


			/// <summary>
			/// Decrypt a stream based on fixed internal keys
			/// </summary>
			/// <param name="EncryptedStream"></param>
			/// <returns></returns>
			private MemoryStream DecryptStream(Stream EncryptedStream)
			{

				try
				{
					System.Security.Cryptography.RijndaelManaged Enc = new System.Security.Cryptography.RijndaelManaged
					{
						KeySize = 256,
						// KEY is 32 byte array
						Key = Constants.ENCKEY,
						// IV is 16 byte array
						IV = Constants.ENCIV
					};

					var cryptoStream = new System.Security.Cryptography.CryptoStream(EncryptedStream, Enc.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Read);

					byte[] buf = null;
					buf = new byte[1024];
					MemoryStream DecryptedStream = new MemoryStream();
					while (EncryptedStream.Length > 0)
					{
						var l = cryptoStream.Read(buf, 0, 1024);
						if (l == 0)
							break; // TODO: might not be correct. Was : Exit Do
						if (l < 1024)
							Array.Resize(ref buf, l);
						DecryptedStream.Write(buf, 0, buf.GetUpperBound(0) + 1);
						if (l < 1024)
							break; // TODO: might not be correct. Was : Exit Do
					}
					DecryptedStream.Position = 0;
					return DecryptedStream;

				}
				catch (Exception ex)
				{
					Debug.WriteLine(string.Format("ERROR: {0}", ex.ToString()));
					// any problems, nothing much to do, so return an empty stream
					return new MemoryStream();
				}
			}


			/// <summary>
			/// Uncompress a memory stream
			/// </summary>
			/// <param name="CompressedStream"></param>
			/// <returns></returns>
			private MemoryStream DecompressStream(MemoryStream CompressedStream)
			{

				System.IO.Compression.GZipStream GZip = new System.IO.Compression.GZipStream(CompressedStream, System.IO.Compression.CompressionMode.Decompress);

				byte[] buf = null;
				buf = new byte[1024];
				MemoryStream UncompressedStream = new MemoryStream();
				while (CompressedStream.Length > 0)
				{
					var l = GZip.Read(buf, 0, 1024);
					if (l == 0)
						break; // TODO: might not be correct. Was : Exit Do
					if (l < 1024)
						Array.Resize(ref buf, l);
					UncompressedStream.Write(buf, 0, buf.GetUpperBound(0) + 1);
					if (l < 1024)
						break; // TODO: might not be correct. Was : Exit Do
				}
				UncompressedStream.Position = 0;
				return UncompressedStream;
			}


			private AssemblyLineMap Depersist(MemoryStream MemoryStream)
			{
				MemoryStream.Position = 0;

				if (MemoryStream.Length != 0)
				{
					var binaryDictionaryreader = XmlDictionaryReader.CreateBinaryReader(MemoryStream, new XmlDictionaryReaderQuotas());
					var serializer = new DataContractSerializer(typeof(AssemblyLineMap));
					return (AssemblyLineMap)serializer.ReadObject(binaryDictionaryreader);
				}
				else
				{
					// just return an empty object
					return new AssemblyLineMap();
				}
			}


			/// <summary>
			/// Stream this object to a memory stream and return it
			/// All other persistence is handled externally because none of that 
			/// is necessary under normal usage, only when generating a line map
			/// </summary>
			/// <returns></returns>
			public MemoryStream ToStream()
			{
				System.IO.MemoryStream MemStream = new System.IO.MemoryStream();

				var binaryDictionaryWriter = XmlDictionaryWriter.CreateBinaryWriter(MemStream);
				var serializer = new DataContractSerializer(typeof(AssemblyLineMap));
				serializer.WriteObject(binaryDictionaryWriter, this);
				binaryDictionaryWriter.Flush();

				return MemStream;
			}


			/// <summary>
			/// Helper function to add an address to line map entry
			/// </summary>
			/// <param name="Line"></param>
			/// <param name="Address"></param>
			/// <param name="SourceFile"></param>
			/// <param name="ObjectName"></param>
			/// <remarks></remarks>
			public void AddAddressToLine(Int32 Line, Int64 Address, string SourceFile, string ObjectName)
			{
				var SourceFileIndex = this.Names.Add(SourceFile);
				var ObjectNameIndex = this.Names.Add(ObjectName);
				var atl = new AssemblyLineMap.AddressToLine(Line, Address, SourceFile, SourceFileIndex, ObjectName, ObjectNameIndex);
				this.AddressToLineMap.Add(atl);
			}
		}


		#region " Constants"
		/// <summary>
		/// Constant values used for resolving line number information from a line map resource
		/// </summary>
		/// <remarks></remarks>
		/// <editHistory></editHistory>
		public static class Constants
		{
			/// <summary>
			/// Tag used for indenting error and stack information to be more readable in string format
			/// </summary>
			public const string INDENT = "[[3]]";

			/// <summary>
			/// Name space used for LineMap XML resource
			/// </summary>
			public const string LINEMAPNAMESPACE = "http://schemas.linemap.net";

			/// <summary>
			/// Encryption key used to obfuscate line map information in the compiled executable
			/// </summary>
			public static byte[] ENCKEY = new byte[32] {
			1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
		};
			/// <summary>
			/// Encryption key used to obfuscate line map information in the compiled executable
			/// </summary>
			public static byte[] ENCIV = new byte[16] {
			65, 2, 68, 26, 7, 178, 200, 3, 65, 110, 68, 13, 69, 16, 200, 219
		};

			/// <summary>
			/// Extension of a standalone line map resource
			/// </summary>
			public static string ResTypeName = "LMP";

			/// <summary>
			/// Name to give to the embedded line map resource
			/// </summary>
			public static string ResName = "LMPDATA";

			/// <summary>
			/// Resource Language to use for embedded line map resource
			/// </summary>
			public static short ResLang = 0;
		}
		#endregion


		#region " API Declarations for working with Resources"
		/// <summary>
		/// Static class to define WindowsAPI calls used during creation, embedding, and reading of line map resources
		/// </summary>
		public static class WindowsAPI
		{
			[DllImport("kernel32", EntryPoint = "FindResourceExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			public static extern IntPtr FindResourceEx(Int32 hModule, [MarshalAs(UnmanagedType.LPStr)]
		string lpType, [MarshalAs(UnmanagedType.LPStr)]
		string lpName, Int16 wLanguage);
			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			public static extern IntPtr LoadResource(IntPtr hInstance, IntPtr hResInfo);
			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			public static extern IntPtr LockResource(IntPtr hResData);
			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			public static extern IntPtr FreeResource(IntPtr hResData);
			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			public static extern Int32 SizeofResource(IntPtr hInstance, IntPtr hResInfo);
			[DllImport("kernel32.dll", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			public static extern void CopyMemory(ref byte pvDest, IntPtr pvSrc, Int32 cbCopy);
		}
		#endregion
	}
	#endregion
}
