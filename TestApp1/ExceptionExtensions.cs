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
		public static SerializableException ToSerializableException(this Exception ex)
		{
			return new SerializableException(ex);
		}
	}


	[Serializable]
	[XmlRoot(ElementName = "Exception")]
	[XmlInclude(typeof(SerializableStackTrace))]
	[KnownType(typeof(SerializableStackTrace))]
	/// <summary>
	/// This serves as a standin for an exception that can contain any arbitary
	/// information for that exception and can be either directly serialized
	/// or have custom serialization applied
	/// </summary>
	public class SerializableException : List<SerializableException.Property>
	{
		#region Constructors
		public SerializableException() { }

		public SerializableException(Exception ex) : this(ex, 0)
		{ }


		protected SerializableException(Exception ex, int level = 0)
		{
			var stackTrace = (level == 0) ? new StackTrace(ex) : null;
			try
			{
				// grab some extended information for the exception
				this["Type"] = ex.GetType().FullName;
				this["Date and Time"] = DateTime.Now.ToString();
				this["Machine Name"] = Environment.MachineName;
				this["Current IP"] = GetCurrentIP();
				this["Current User"] = GetUserIdentity();
				this["Application Domain"] = System.AppDomain.CurrentDomain.FriendlyName;
				this["Assembly Codebase"] = Utilities.ParentAssembly.CodeBase;
				this["Assembly Fullname"] = Utilities.ParentAssembly.FullName;
				this["Assembly Version"] = Utilities.ParentAssembly.GetName().Version.ToString();
				this["Assembly Build Date"] = GetAssemblyBuildDate(Utilities.ParentAssembly).ToString();
				this.GetExceptionProperties(ex, 0, stackTrace);
			}
			catch (Exception ex2)
			{
				this.Clear();
				this["Message"] = string.Format("{0} Error '{1}' while generating exception description", ex2.GetType().Name, ex2.Message);
			}
		}
		#endregion


		[Serializable]
		[XmlRoot(ElementName = "Properties")]
		public class PropertyList : List<Property>
		{
			public object this[string key]
			{
				get
				{
					var r = this.Where(i => i.Key == key).FirstOrDefault();
					return r.Value;
				}
				set
				{
					if (!string.IsNullOrEmpty(key) && value != null)
						this.Add(new Property(key, value));
				}
			}
		}


		public object this[string key]
		{
			get
			{
				var r = this.Where(i => i.Key == key).FirstOrDefault();
				if (r != null) return r.Value;
				return null;
			}
			set
			{
				if (!string.IsNullOrEmpty(key) && value != null)
					this.Add(new Property(key, value));
			}
		}


		#region Internal Info retrieval functions

		/// <summary>
		/// Retrieve all the relavent properties of an exception 
		/// </summary>
		/// <param name="ex"></param>
		/// <param name="level"></param>
		private void GetExceptionProperties(Exception ex, int level = 0, StackTrace stackTrace = null)
		{
			if (stackTrace != null)
			{
				//requesting the value for stack trace just renders to a string using the internal
				//.net functionality. We don't want that.
				//trace MUST be created in the constructor above or it won't be correct
				//so it's created there, and passed through to here as an arg
				this["StackTrace"] = new SerializableStackTrace(stackTrace);
			}

			// gather up all the properties of the Exception, plus the extended info above
			// sort it, and render to a stringbuilder
			foreach (var item in ex
				.GetType()
				.GetProperties())
			{
				var name = item.Name;
				if (name != "StackTrace")
				{
					var value = item.GetValue(ex, null);
					if (value is Exception)
					{
						this["InnerException"] = new SerializableException((Exception)value, ++level);
					}
					else
					{
						this[name] = value;
					}
				}
			}
		}


		/// <summary>
		/// get IP address of this machine
		/// not an ideal method for a number of reasons (guess why!)
		/// but the alternatives are very ugly
		/// </summary>
		/// <returns></returns>
		private string GetCurrentIP()
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
		private string GetUserIdentity()
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
		private string GetCurrentWindowsIdentity()
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
		private string GetCurrentEnvironmentIdentity()
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
		private DateTime GetAssemblyBuildDate(System.Reflection.Assembly objAssembly, bool bForceFileDate = false)
		{
			var objVersion = objAssembly.GetName().Version;
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
		private DateTime GetAssemblyFileTime(System.Reflection.Assembly objAssembly)
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
		#endregion


		[Serializable]
		[XmlRoot(ElementName = "Property")]
		public class Property
		{
			public string Key { get; set; }
			public object Value { get; set; }

			public Property() { }

			public Property(string key, object value)
			{
				Key = key;
				Value = value;
			}
		}


		[Serializable]
		[XmlRoot(ElementName = "StackTrace")]
		public class SerializableStackTrace
		{
			public SerializableStackTrace()
			{
			}


			public SerializableStackTrace(StackTrace stackTrace)
			{
				foreach (var stackFrame in stackTrace.GetFrames())
				{
					this.StackFrames.Add(new SerializableStackFrame(stackFrame));
				}
			}

			[XmlArray]
			public SerializableStackFrameList StackFrames = new SerializableStackFrameList();
		}


		[Serializable]
		[XmlRoot(ElementName = "StackFrames")]
		public class SerializableStackFrameList : List<SerializableStackFrame>
		{
		}


		[Serializable]
		[XmlRoot(ElementName = "StackFrame")]
		public class SerializableStackFrame
		{
			public SerializableStackFrame()
			{
			}


			public SerializableStackFrame(StackFrame stackFrame)
			{
				this.ILOffset = stackFrame.GetILOffset();
				this.NativeOffset = stackFrame.GetNativeOffset();
				this.MethodBase = new SerializableMethodBase(stackFrame.GetMethod());

				if (stackFrame.GetFileName() != null && stackFrame.GetFileName().Length != 0 && ExceptionExtensions.AllowUseOfPDB)
				{
					// the PDB appears to be available, since the above elements are 
					// not blank, so just use it's information
					this.FileName = System.IO.Path.GetFileName(stackFrame.GetFileName());
					this.FileLineNumber = stackFrame.GetFileLineNumber();
					this.FileColumnNumber = stackFrame.GetFileColumnNumber();
				}
				else
				{
					// the PDB is not available, so attempt to retrieve 
					// any embedded linemap information
					if (Utilities.ParentAssembly != null)
					{
						this.FileName = System.IO.Path.GetFileName(Utilities.ParentAssembly.CodeBase);
					}
					else
					{
						this.FileName = "Unable to determine Assembly Filename";
					}

					// Get the native code offset and convert to a line number
					// first, make sure our linemap is loaded
					try
					{
						var mi = (MethodInfo)stackFrame.GetMethod();
						if (mi != null)
						{
							LineMap.AssemblyLineMaps.Add(mi.DeclaringType.Assembly);

							var sl = stackFrame.MapStackFrameToSourceLine();
							if (sl.Line != 0)
							{
								this.FileName = sl.SourceFile;
								this.FileLineNumber = sl.Line;
							}
						}
					}
					catch (Exception ex)
					{
						// any problems in loading the Linemap, just write to debugger and call it a day
						Debug.WriteLine(string.Format("Unable to load line map information. Error: {0}", ex.ToString()));
					}
				}
			}
			public int FileColumnNumber { get; set; }
			public int FileLineNumber { get; set; }
			public string FileName { get; set; }
			public int ILOffset { get; set; }
			public SerializableMethodBase MethodBase { get; set; }
			public int NativeOffset { get; set; }
		}


		[Serializable]
		[XmlRoot(ElementName = "Method")]
		public class SerializableMethodBase
		{
			public SerializableMethodBase() { }


			public SerializableMethodBase(MethodBase methodBase)
			{
				this.DeclaringTypeNameSpace = methodBase.DeclaringType.Namespace;
				this.DeclaringTypeName = methodBase.DeclaringType.Name;
				this.Name = methodBase.Name;
				this.Parameters = new SerializableParameterInfoList(methodBase.GetParameters());
			}
			public string DeclaringTypeNameSpace { get; set; }
			public string DeclaringTypeName { get; set; }
			public string Name { get; set; }
			public SerializableParameterInfoList Parameters { get; set; }
		}


		[Serializable]
		[XmlRoot(ElementName = "Parameters")]
		public class SerializableParameterInfoList : List<SerializableParameterInfo>
		{
			public SerializableParameterInfoList() { }
			public SerializableParameterInfoList(ParameterInfo[] parameterInfo)
			{
				foreach (var pi in parameterInfo)
				{
					this.Add(new SerializableParameterInfo(pi));
				}
			}
		}


		[Serializable]
		[XmlRoot(ElementName = "Parameter")]
		public class SerializableParameterInfo
		{
			public SerializableParameterInfo() { }
			public SerializableParameterInfo(ParameterInfo parameterInfo)
			{
				this.Name = parameterInfo.Name;
				this.Type = parameterInfo.ParameterType.Name;
			}
			public string Name { get; set; }
			public string Type { get; set; }
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
	}
}
