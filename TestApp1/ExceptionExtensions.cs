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
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Xml.Serialization;

using ExceptionExtensions.Internal;


namespace ExceptionExtensions
{
	/// <summary>
	/// Exception Extension methods.
	/// Portions are from a post on StackOverflow here:
	/// https://stackoverflow.com/questions/2176707/exception-message-vs-exception-tostring
	/// </summary>
	public static partial class ExceptionExtensions
	{
		/// <summary>
		/// Force the use of the PDB if available
		/// This is primarily used when debugging/Unit testing
		/// </summary>
		public static bool AllowUseOfPDB { get; set; }


		/// <summary>
		/// translate exception object to string, with additional system info
		/// The serializable exception object is much easier to work with, serialize, convert to string
		/// or retrieve specific information from.
		/// </summary>
		/// <param name="ex"></param>
		/// <returns></returns>
		public static string ToExpandedString(this Exception ex)
		{
			return ToExpandedString(ex, new ExceptionOptions());
		}


		/// <summary>
		/// translate exception object to string, with additional system info
		/// The serializable exception object is much easier to work with, serialize, convert to string
		/// or retrieve specific information from.
		/// </summary>
		/// <param name="ex"></param>
		/// <returns></returns>
		public static string ToExpandedString(this Exception ex, ExceptionOptions options)
		{
			return _toExpandedString(ex, options);
		}


		/// <summary>
		/// This method provides the default ToString rendering
		/// </summary>
		/// <returns></returns>
		private static string _toExpandedString(Exception ex, ExceptionOptions options)
		{
			if (ex == null) return string.Empty;

			StringBuilder sb = new StringBuilder();

			try
			{
				// grab some extended information for the exception
				var extendedProps = new List<KeyValuePair<string, object>>();

				// only need the extended properties once
				if (options.CurrentIndentLevel == 0)
				{
					sb.AppendLine(string.Format("{0}Exception: {1}", options.Indent, ex.GetType().FullName));

					extendedProps.Add(new KeyValuePair<string, object>("Type", ex.GetType().FullName));
					extendedProps.Add(new KeyValuePair<string, object>("Date and Time", DateTime.Now.ToString()));
					extendedProps.Add(new KeyValuePair<string, object>("Machine Name", Environment.MachineName));
					extendedProps.Add(new KeyValuePair<string, object>("Current IP", Utilities.GetCurrentIP()));
					extendedProps.Add(new KeyValuePair<string, object>("Current User", Utilities.GetUserIdentity()));
					extendedProps.Add(new KeyValuePair<string, object>("Application Domain", System.AppDomain.CurrentDomain.FriendlyName));
					extendedProps.Add(new KeyValuePair<string, object>("Assembly Codebase", Utilities.ParentAssembly.CodeBase));
					extendedProps.Add(new KeyValuePair<string, object>("Assembly Fullname", Utilities.ParentAssembly.FullName));
					extendedProps.Add(new KeyValuePair<string, object>("Assembly Version", Utilities.ParentAssembly.GetName().Version.ToString()));
					extendedProps.Add(new KeyValuePair<string, object>("Assembly Build Date", Utilities.GetAssemblyBuildDate(Utilities.ParentAssembly).ToString()));
				}

				// gather up all the properties of the Exception, plus the extended info above
				// sort it, and render to a stringbuilder
				foreach (KeyValuePair<string, object> item in ex
					.GetType()
					.GetProperties()
					.Select(x => new KeyValuePair<string, object>(x.Name, x))
					.Concat(extendedProps)
					.OrderByDescending(x => string.Equals(x.Key, "Type", StringComparison.Ordinal))
					.ThenByDescending(x => string.Equals(x.Key, nameof(ex.Message), StringComparison.Ordinal))
					.ThenByDescending(x => string.Equals(x.Key, nameof(ex.Source), StringComparison.Ordinal))
					.ThenBy(x => string.Equals(x.Key, nameof(ex.InnerException), StringComparison.Ordinal))
					.ThenBy(x => string.Equals(x.Key, nameof(AggregateException.InnerExceptions), StringComparison.Ordinal))
					.ThenBy(x => x.Key))
				{
					object value = item.Value;
					if (item.Key == "StackTrace")
					{
						//handle the stacktrace special
						var buf = new StackTrace(ex).ToString(true).TrimEnd('\r', '\n').Replace("\r\n", string.Format("\r\n{0, -23}", "")).TrimEnd();
						if (string.IsNullOrEmpty(buf) && options.OmitNullProperties) continue;
						value = buf;
					}
					else if (item.Value is PropertyInfo)
					{
						value = (item.Value as PropertyInfo).GetValue(ex, null);
						if (value == null || (value is string && string.IsNullOrEmpty((string)value)))
						{
							if (options.OmitNullProperties)
							{
								continue;
							}
							else
							{
								value = string.Empty;
							}
						}
					}

					sb.AppendValue(item.Key, value, options);
				}

				return sb.ToString().TrimEnd('\r', '\n');
			}
			catch (Exception ex2)
			{
				sb.Clear();
				sb.AppendLine(string.Format("{0} Error '{1}' while generating exception description", ex2.GetType().Name, ex2.Message));
			}

			return sb.ToString();
		}
	}
}


/// <summary>
/// Seperate namespace to allow extension of the StringBuilder and Stacktrace without polluting it elsewhere in the host app
/// </summary>
namespace ExceptionExtensions.Internal
{
	public static class StackTraceExtensions
	{
		public static string ToString(this StackTrace stackTrace, bool skipLocalFrames = false)
		{
			StringBuilder sb = new StringBuilder();

			for (int intFrame = 0; intFrame <= stackTrace.FrameCount - 1; intFrame++)
			{
				StackFrame sf = stackTrace.GetFrame(intFrame);

				if (skipLocalFrames && sf.GetMethod().DeclaringType.Name.IndexOf(Utilities.CLASSNAME) > -1)
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

			return sb.ToString();
		}


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
			sb.Append("   ");
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
				if (Utilities.ParentAssembly != null)
				{
					Filename = System.IO.Path.GetFileName(Utilities.ParentAssembly.CodeBase);
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

		
		/// <summary>
		/// Map an address offset from a stack frame entry to a linenumber
		/// using the Method name, the base address of the method and the
		/// IL offset from the base address
		/// </summary>
		/// <param name="sf"></param>
		/// <param name="Line"></param>
		/// <param name="SourceFile"></param>
		/// <remarks></remarks>
		public static SourceLine MapStackFrameToSourceLine(this StackFrame sf)
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
		/// Used to pass linemap information to the stackframe renderer
		/// </summary>
		public class SourceLine
		{
			public string SourceFile = string.Empty;
			public int Line = 0;
		}
	}


	public static class Utilities
	{
		/// <summary>
		/// Used to identify stack frames that relate to "this" class so they can be skipped
		/// </summary>
		public static string CLASSNAME = "GenerateLineMapExceptionExtensions";


		private static string IndentString(string value, ExceptionOptions options)
		{
			return value.Replace(Environment.NewLine, Environment.NewLine + options.Indent);
		}


		private static Assembly _parentAssembly = null;
		/// <summary>
		/// Retrieve the root assembly of the executing assembly
		/// </summary>
		/// <returns></returns>
		public static Assembly ParentAssembly
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
		/// exception-safe file attrib retrieval; we don't care if this fails
		/// </summary>
		/// <param name="objAssembly"></param>
		/// <returns></returns>
		public static DateTime GetAssemblyFileTime(System.Reflection.Assembly objAssembly)
		{
			try
			{
				return System.IO.File.GetLastWriteTime(objAssembly.Location);
			}
			catch
			{
				// just provide a default value
				return DateTime.MinValue;
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
		public static DateTime GetAssemblyBuildDate(System.Reflection.Assembly objAssembly, bool bForceFileDate = false)
		{
			var objVersion = objAssembly.GetName().Version;
			DateTime dtBuild = default(DateTime);

			if (bForceFileDate)
			{
				dtBuild = Utilities.GetAssemblyFileTime(objAssembly);
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
					dtBuild = Utilities.GetAssemblyFileTime(objAssembly);
				}
			}

			return dtBuild;
		}


		/// <summary>
		/// get IP address of this machine
		/// not an ideal method for a number of reasons (guess why!)
		/// but the alternatives are very ugly
		/// </summary>
		/// <returns></returns>
		public static string GetCurrentIP()
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
		public static string GetUserIdentity()
		{
			string strTemp = GetCurrentWindowsIdentity();
			if (string.IsNullOrEmpty(strTemp))
			{
				strTemp = GetCurrentEnvironmentIdentity();
			}
			return strTemp;
		}


		/// <summary>
		/// exception-safe WindowsIdentity.GetCurrent retrieval returns "domain\username"
		/// per MS, this sometimes randomly fails with "Access Denied" particularly on NT4
		/// </summary>
		/// <returns></returns>
		private static string GetCurrentWindowsIdentity()
		{
			try
			{
				return System.Security.Principal.WindowsIdentity.GetCurrent().Name;
			}
			catch
			{
				// just provide a default value
				return string.Empty;
			}
		}


		/// <summary>
		/// exception-safe "domain\username" retrieval from Environment
		/// </summary>
		/// <returns></returns>
		private static string GetCurrentEnvironmentIdentity()
		{
			try
			{
				return System.Environment.UserDomainName + "\\" + System.Environment.UserName;
			}
			catch
			{
				// just provide a default value
				return string.Empty;
			}
		}



	}
}

