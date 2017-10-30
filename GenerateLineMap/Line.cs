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
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;


/// <summary>
/// This file will need to be included in your project to allow you to retrieve 
/// Line number and source file information from the embedded GenerateLineMap resource
/// 
/// REFERENCES REQUIRED:
///		System.Runtime.Serialization
/// </summary>


namespace ExceptionExtensions
{
	public static class ExceptionExtensions
	{
		private static bool _ignorePDB = false;

		/// <summary>
		/// really only for debugging
		/// </summary>
		public static void IgnorePDB(this Exception ex)
		{
			_ignorePDB = true;
		}


		public static int GetLine(this Exception ex)
		{
			return LineMap.InfoFrom(ex).Line;
		}


		public static string GetSourceFile(this Exception ex)
		{
			return LineMap.InfoFrom(ex).SourceFile;
		}


		/// <summary>
		/// This file contains the necessary logic to deserialize a line map resource from an
		/// assembly and resolve a line number from a stack trace, if possible.
		/// 
		/// This File (and ONLY this file) is included in the NUGET as a content item
		/// that gets inserted into the target project.
		/// 
		/// </summary>
		/// <remarks></remarks>
		/// <editHistory></editHistory>
		private static class LineMap
		{
			private const string LINEMAPNAMESPACE = "http://schemas.linemap.net";


			/// <summary>
			/// Resolve a Line number from the stack trace of an exception
			/// </summary>
			/// <param name="sf"></param>
			/// <param name="Line"></param>
			/// <param name="SourceFile"></param>
			/// <remarks></remarks>
			public static SourceInfo InfoFrom(Exception ex)
			{
				var sl = new SourceInfo();

				//get stack frame from the Exception
				var stackTrace = new StackTrace(ex);
				var stackFrame = stackTrace.GetFrame(0);

				if (!_ignorePDB)
				{
					sl.Line = stackFrame.GetFileLineNumber();
					if (sl.Line != 0) return sl;
				}


				// first, get the base addr of the method
				// if possible
				// you have to have symbols to do this
				if (AssemblyLineMaps.Count == 0)
					return sl;

				// first, check if for symbols for the assembly for this stack frame
				if (!AssemblyLineMaps.Keys.Contains(stackFrame.GetMethod().DeclaringType.Assembly.CodeBase))
				{
					AssemblyLineMaps.Add(stackFrame.GetMethod().DeclaringType.Assembly);
				}

				// if it's still not available, not much else we can do
				if (!AssemblyLineMaps.Keys.Contains(stackFrame.GetMethod().DeclaringType.Assembly.CodeBase))
				{
					return sl;
				}

				// retrieve the cache
				var alm = AssemblyLineMaps[stackFrame.GetMethod().DeclaringType.Assembly.CodeBase];

				// does the symbols list contain the metadata token for this method?
				MemberInfo mi = stackFrame.GetMethod();
				// Don't call this mdtoken or PostSharp will barf on it! Jeez
				long mdtokn = mi.MetadataToken;
				if (!alm.Symbols.ContainsKey(mdtokn))
					return sl;

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
					sl.Line = alm.AddressToLineMap[i].Line;
					sl.SourceFile = alm.Names[alm.AddressToLineMap[i].SourceFileIndex];
				}

				return sl;
			}


			/// <summary>
			/// Used to pass linemap information to the stackframe renderer
			/// </summary>
			public class SourceInfo
			{
				public string SourceFile = string.Empty;
				public int Line = 0;
			}


			#region " API Declarations for working with Resources"
			[DllImport("kernel32", EntryPoint = "FindResourceExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			private static extern IntPtr FindResourceEx(Int32 hModule, [MarshalAs(UnmanagedType.LPStr)]
				string lpType, [MarshalAs(UnmanagedType.LPStr)]
				string lpName, Int16 wLanguage);

			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			private static extern IntPtr LoadResource(IntPtr hInstance, IntPtr hResInfo);

			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			private static extern IntPtr LockResource(IntPtr hResData);

			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			private static extern IntPtr FreeResource(IntPtr hResData);

			[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			private static extern Int32 SizeofResource(IntPtr hInstance, IntPtr hResInfo);

			[DllImport("kernel32.dll", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
			private static extern void CopyMemory(ref byte pvDest, IntPtr pvSrc, Int32 cbCopy);
			#endregion


			/// <summary>
			/// The class simply defines several values we need to share between 
			/// the GenerateLineMap utility and whatever project needs to read line numbers
			/// </summary>
			/// <remarks></remarks>
			/// <editHistory></editHistory>
			internal class LineMapKeys
			{
				public static byte[] ENCKEY = new byte[32] {
					1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32
				};
				public static byte[] ENCIV = new byte[16] {
					65, 2, 68, 26, 7, 178, 200, 3, 65, 110, 68, 13, 69, 16, 200, 219
				};
				public static string ResTypeName = "LINEMAP";
				public static string ResName = "LINEMAPDATA";
				public static short ResLang = 0;
			}


			/// <summary>
			/// Tracks symbol cache entries for all assemblies that appear in a stack frame
			/// </summary>
			/// <remarks></remarks>
			/// <editHistory></editHistory>
			[DataContract(Namespace = LINEMAPNAMESPACE)]
			public class AssemblyLineMap
			{
				/// <summary>
				/// Track the assembly this map is for
				/// no need to persist this information
				/// </summary>
				/// <remarks></remarks>

				public string FileName;
				[DataContract(Namespace = LINEMAPNAMESPACE)]
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
					public AddressToLine()
					{
					}


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
				[DataContract(Namespace = LINEMAPNAMESPACE)]
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
					public SymbolInfo()
					{
					}


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
					/// Override this prop so that requested an index that doesn't exist just
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

				public AssemblyLineMap()
				{
				}


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
					Stream buf = FileToStream(this.FileName + ".linemap");
					var buf2 = DecryptStream(buf);
					Transfer(Depersist(DecompressStream(buf2)));
				}


				/// <summary>
				/// Create a new assemblylinemap based on an assembly
				/// </summary>
				/// <remarks></remarks>
				public AssemblyLineMap(Assembly Assembly)
				{
					//LoadLineMapResource(Assembly.GetExecutingAssembly(), Symbols, Lines)

					this.FileName = Assembly.CodeBase.Replace("file:///", "");

					try
					{
						// Get the hInstance of the indicated exe/dll image
						var curmodule = Assembly.GetLoadedModules()[0];
						var hInst = System.Runtime.InteropServices.Marshal.GetHINSTANCE(curmodule);

						// retrieve a handle to the Linemap resource
						// Since it's a standard Win32 resource, the nice .NET resource functions
						// can't be used
						//
						// Important Note: The FindResourceEx function appears to be case
						// sensitive in that you really HAVE to pass in UPPER CASE search
						// arguments
						var hres = FindResourceEx(hInst.ToInt32(), LineMapKeys.ResTypeName, LineMapKeys.ResName, LineMapKeys.ResLang);

						byte[] bytes = null;
						if (hres != IntPtr.Zero)
						{
							// Load the resource to get it into memory
							var hresdata = LoadResource(hInst, hres);

							IntPtr lpdata = LockResource(hresdata);
							var sz = SizeofResource(hInst, hres);

							if (lpdata != IntPtr.Zero & sz > 0)
							{
								// able to lock it,
								// so copy the data into a byte array
								bytes = new byte[sz];
								CopyMemory(ref bytes[0], lpdata, sz);
								FreeResource(hresdata);
							}
						}
						else
						{
							//Check for a side by side linemap file
							string mapfile = this.FileName + ".linemap";
							if (File.Exists(mapfile))
							{
								//load it from there
								bytes = File.ReadAllBytes(mapfile);
							}
						}

						// deserialize the symbol map and line num list
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
					catch (Exception ex)
					{
						System.Diagnostics.Debug.WriteLine(string.Format("ERROR: {0}", ex.ToString()));
						// yes, it's bad form to catch all exceptions like this
						// but this is part of an exception handler
						// so it really can't be allowed to fail with an exception!
					}

					try
					{
						if (this.Symbols.Count == 0)
						{
							// weren't able to load resources, so try the LINEMAP file
							Load();
						}

					}
					catch (Exception ex)
					{
						System.Diagnostics.Debug.WriteLine(string.Format("ERROR: {0}", ex.ToString()));
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
						System.Security.Cryptography.RijndaelManaged Enc = new System.Security.Cryptography.RijndaelManaged();
						Enc.KeySize = 256;
						// KEY is 32 byte array
						Enc.Key = LineMapKeys.ENCKEY;
						// IV is 16 byte array
						Enc.IV = LineMapKeys.ENCIV;

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
						System.Diagnostics.Debug.WriteLine(string.Format("ERROR: {0}", ex.ToString()));
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
						// jsut return an empty object
						return new AssemblyLineMap();
					}
				}


				/// <summary>
				/// Stream this object to a memory stream and return it
				/// All other persistence is handled externally because none of that 
				/// is necessary under normal usage, only when generating a linemap
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


			/// <summary>
			/// Define a collection of assemblylinemaps (one for each assembly we
			/// need to generate a stack trace through)
			/// </summary>
			/// <remarks></remarks>
			/// <editHistory></editHistory>
			public class AssemblyLineMapCollection : Dictionary<string, AssemblyLineMap>
			{

				/// <summary>
				/// Load a linemap file given the NAME of an assembly
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
	}
}