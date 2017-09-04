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
	[DataContract]
	[XmlRoot(ElementName = "Exception")]
	[KnownType(typeof(SerializableStackTrace))]
	/// <summary>
	/// This serves as a standin for an exception that can contain any arbitary
	/// information for that exception and can be either directly serialized
	/// or have custom serialization applied
	/// </summary>
	public class SerializableException
	{
		#region Constructors
		public SerializableException() : this(null, 0)
		{ }

		public SerializableException(Exception ex) : this(ex, 0)
		{ }


		protected SerializableException(Exception ex, int level = 0)
		{
			_properties = new PropertyList();

			if (ex != null)
			{
				var stackTrace = (level == 0) ? new StackTrace(ex) : null;
				try
				{
					// grab some extended information for the exception
					this["Type"] = ex.GetType().FullName;
					this["Date and Time"] = DateTime.Now.ToString();
					this["Machine Name"] = Environment.MachineName;
					this["Current IP"] = Utilities.GetCurrentIP();
					this["Current User"] = Utilities.GetUserIdentity();
					this["Application Domain"] = System.AppDomain.CurrentDomain.FriendlyName;
					this["Assembly Codebase"] = Utilities.ParentAssembly.CodeBase;
					this["Assembly Fullname"] = Utilities.ParentAssembly.FullName;
					this["Assembly Version"] = Utilities.ParentAssembly.GetName().Version.ToString();
					this["Assembly Build Date"] = Utilities.GetAssemblyBuildDate(Utilities.ParentAssembly).ToString();
					this.GetExceptionProperties(ex, 0, stackTrace);
				}
				catch (Exception ex2)
				{
					this.Properties.Clear();
					this["Message"] = string.Format("{0} Error '{1}' while generating exception description", ex2.GetType().Name, ex2.Message);
				}
			}
		}
		#endregion


		private PropertyList _properties;
		[DataMember]
		public PropertyList Properties
		{
			get
			{
				return _properties;
			}
		}

		public object this[string name]
		{
			get
			{
				if (this.Properties.ContainsKey(name))
				{
					return this.Properties[name];
				}
				return null;
			}
			set
			{
				if (!string.IsNullOrEmpty(name) && value != null)
					this.Properties[name] = value;
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
						if (value == null || (value != null && value.GetType().IsPrimitive))
						{
							//primitive types should always be serializable
							this[name] = value;
						}
						else
						{
							//non primitive types need checking
							//TBD for now, just ToString it
							this[name] = value.ToString();
						}
					}
				}
			}
		}


		#endregion


		[Serializable]
		[CollectionDataContract]
		[XmlRoot(ElementName = "Properties")]
		public class PropertyList : Dictionary<string, object>
		{ }


		[Serializable]
		[DataContract]
		[XmlRoot(ElementName = "Property")]
		public class Property
		{
			[DataMember]
			public string Name { get; set; }
			[DataMember]
			public object Value { get; set; }

			public Property() { }

			public Property(string name, object value)
			{
				Name = name;
				Value = value;
			}
		}


		[Serializable]
		[DataContract]
		[XmlRoot(ElementName = "StackTrace")]
		//[KnownType(typeof(SerializableStackFrameList))]
		public class SerializableStackTrace
		{
			public SerializableStackTrace()
			{
				this.StackFrames = new SerializableStackFrameList();
			}


			public SerializableStackTrace(StackTrace stackTrace) : this()
			{
				foreach (var stackFrame in stackTrace.GetFrames())
				{
					this.StackFrames.Add(new SerializableStackFrame(stackFrame));
				}
			}


			[DataMember]
			[XmlArray]
			public SerializableStackFrameList StackFrames { get; private set; }
		}


		[Serializable]
		[CollectionDataContract]
		[XmlRoot(ElementName = "StackFrames")]
		//[KnownType(typeof(SerializableStackFrame))]
		public class SerializableStackFrameList : List<SerializableStackFrame>
		{
		}


		[Serializable]
		[DataContract]
		[XmlRoot(ElementName = "StackFrame")]
		//[KnownType(typeof(SerializableMethodBase))]
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

				if (stackFrame.GetFileName() != null && stackFrame.GetFileName().Length != 0 && Utilities.AllowUseOfPDB)
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
			[DataMember]
			public int FileColumnNumber { get; set; }
			[DataMember]
			public int FileLineNumber { get; set; }
			[DataMember]
			public string FileName { get; set; }
			[DataMember]
			public int ILOffset { get; set; }
			[DataMember]
			public SerializableMethodBase MethodBase { get; set; }
			[DataMember]
			public int NativeOffset { get; set; }
		}


		[Serializable]
		[DataContract]
		[XmlRoot(ElementName = "Method")]
		//[KnownType(typeof(SerializableParameterInfoList))]
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
			[DataMember]
			public string DeclaringTypeNameSpace { get; set; }
			[DataMember]
			public string DeclaringTypeName { get; set; }
			[DataMember]
			public string Name { get; set; }
			[DataMember]
			public SerializableParameterInfoList Parameters { get; set; }
		}


		[Serializable]
		[CollectionDataContract]
		[XmlRoot(ElementName = "Parameters")]
		//[KnownType(typeof(SerializableParameterInfo))]
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
		[DataContract]
		[XmlRoot(ElementName = "Parameter")]
		public class SerializableParameterInfo
		{
			public SerializableParameterInfo() { }
			public SerializableParameterInfo(ParameterInfo parameterInfo)
			{
				this.Name = parameterInfo.Name;
				this.Type = parameterInfo.ParameterType.Name;
				this.Value = "TBD";
			}
			[DataMember]
			public string Name { get; set; }
			[DataMember]
			public string Type { get; set; }
			[DataMember]
			public string Value { get; set; }
		}
	}
}
