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
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using DotNetLineNumbers;


namespace GenerateLineMap
{
	/// <summary>
	/// Utility Class to generate a LineMap based on the
	/// contents of the PDB for a specific EXE or DLL
	/// </summary>
	/// <remarks>
	/// If some of the functions seem a little odd here, that's mostly because
	/// I was experimenting with "Fluent" style apis during most of the 
	/// development of this project.
	/// </remarks>
	public class LineMapBuilder
	{

		#region " Structures"
		/// <summary>
		/// Structure used when Enumerating symbols from a PDB file
		/// Note that this structure must be marshalled very specifically
		/// to be compatible with the SYM* dbghelp dll calls
		/// </summary>
		/// <remarks></remarks>
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
		private struct SYMBOL_INFO
		{
			public int SizeOfStruct;
			public int TypeIndex;
			public Int64 Reserved0;
			public Int64 Reserved1;
			public int Index;
			public int size;
			public Int64 ModBase;
			public CV_SymbolInfoFlags Flags;
			public Int64 Value;
			public Int64 Address;
			public int Register;
			public int Scope;
			public int Tag;
			public int NameLen;
			public int MaxNameLen;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 300)]
			public string Name;
		}

		private const int MAX_PATH = 260;


		/// <summary>
		/// Structure used when Enumerating line numbers from a PDB file
		/// Note that this structure must be marshalled very specifically
		/// to be compatible with the SYM* dbghelp dll calls
		/// </summary>
		/// <remarks></remarks>
		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
		private struct SRCCODEINFO
		{
			public int SizeOfStruct;
			public IntPtr Key;
			public Int64 ModBase;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH + 1)]
			public string obj;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH + 1)]
			public string FileName;
			public int LineNumber;
			public Int64 Address;
		}


		private enum SYM_TYPE
		{
			SymNone = 0,
  			SymCoff = 1,
  			SymCv = 2,
  			SymPdb = 3,
  			SymExport = 4,
  			SymDeferred = 5,
  			SymSym = 6,
  			SymDia = 7,
  			SymVirtual = 8,
		}


		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
		private struct IMAGEHLP_MODULE
		{
			public int SizeOfStruct;
			public int BaseOfImage;
			public int ImageSize;
			public int TimeDateStamp;
			public int CheckSum;
			public int NumSyms;
			public SYM_TYPE SymType;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
			public string ModuleName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
			public string ImageName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
			public string LoadedImageName;

			public static IMAGEHLP_MODULE Create()
			{
				var r = new IMAGEHLP_MODULE();
				r.SizeOfStruct = System.Runtime.InteropServices.Marshal.SizeOf(r);
				return r;
			}
		}


		[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
		private struct IMAGEHLP_MODULE64
		{
			public int SizeOfStruct;
			public Int64 BaseOfImage;
			public int ImageSize;
			public int TimeDateStamp;
			public int CheckSum;
			public int NumSyms;
			public SYM_TYPE SymType;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
			public string ModuleName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
			public string ImageName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
			public string LoadedImageName;
			[MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
			public string LoadedPdbName;
			public int CVSig;
			[MarshalAs(UnmanagedType.LPStr, SizeConst = MAX_PATH + 3)]
			public string CVData;
			public int PdbSig;
			Guid PdbSig70;
			public int PdbAge;
			public bool PdbUnmatched;
			public bool DbgUnmatched;
			public bool LineNumbers;
			public bool GlobalSymbols;
			public bool TypeInfo;
			public bool SourceIndexed;
			public bool Publics;
			public int MachineType;
			public int Reserved;
			[MarshalAs(UnmanagedType.LPStr, SizeConst = MAX_PATH)]
			public string Padding;

			public static IMAGEHLP_MODULE64 Create()
			{
				var r = new IMAGEHLP_MODULE64();
				r.SizeOfStruct = System.Runtime.InteropServices.Marshal.SizeOf(r);
				return r;
			}
		}


		private enum CV_SymbolInfoFlags
		{
			IMAGEHLP_SYMBOL_INFO_VALUEPRESENT = 0x1,
			IMAGEHLP_SYMBOL_INFO_REGISTER = 0x8,
			IMAGEHLP_SYMBOL_INFO_REGRELATIVE = 0x10,
			IMAGEHLP_SYMBOL_INFO_FRAMERELATIVE = 0x20,
			IMAGEHLP_SYMBOL_INFO_PARAMETER = 0x40,
			IMAGEHLP_SYMBOL_INFO_LOCAL = 0x80,
			IMAGEHLP_SYMBOL_INFO_CONSTANT = 0x100,
			IMAGEHLP_SYMBOL_FUNCTION = 0x800,
			IMAGEHLP_SYMBOL_VIRTUAL = 0x1000,
			IMAGEHLP_SYMBOL_THUNK = 0x2000,
			IMAGEHLP_SYMBOL_TLSREL = 0x4000,
			IMAGEHLP_SYMBOL_SLOT = 0x8000,
			IMAGEHLP_SYMBOL_ILREL = 0x10000,
			IMAGEHLP_SYMBOL_METADATA = 0x20000,
			IMAGEHLP_SYMBOL_CLR_TOKEN = 0x40000,
		}
		#endregion


		#region Callbacks
		private delegate bool PSYM_ENUMLINES_CALLBACK(ref SRCCODEINFO srcinfo, IntPtr userContext);


		/// <summary>
		/// Delegate to handle Symbol Enumeration Callback
		/// </summary>
		/// <param name="syminfo"></param>
		/// <param name="Size"></param>
		/// <param name="UserContext"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private delegate int PSYM_ENUMSYMBOLS_CALLBACK(ref SYMBOL_INFO syminfo, int Size, int UserContext);

		#endregion


		#region " API Definitions"
		private int SymEnumSymbolsCallback_proc(ref SYMBOL_INFO syminfo, int Size, int UserContext)
		{
			if ((syminfo.Flags &
				(CV_SymbolInfoFlags.IMAGEHLP_SYMBOL_CLR_TOKEN |
				CV_SymbolInfoFlags.IMAGEHLP_SYMBOL_METADATA |
				CV_SymbolInfoFlags.IMAGEHLP_SYMBOL_FUNCTION))
				!= 0)
			{
				// we only really care about CLR metadata tokens
				// anything else is basically a variable or internal
				// info we wouldn't be worried about anyway.
				// This might change and I get to know more about debugging
				//.net!

				var Tokn = syminfo.Value;
				var si = new AssemblyLineMap.SymbolInfo(syminfo.Name.Substring(0, syminfo.NameLen), syminfo.Address, Tokn);

				_alm.Symbols.Add(Tokn, si);
			}

			// return this to the call to let it keep enumerating
			return -1;
		}


		/// <summary>
		/// Handle the callback to enumerate all the Line numbers in the given DLL/EXE
		/// based on info from the PDB
		/// </summary>
		/// <param name="srcinfo"></param>
		/// <param name="UserContext"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private bool SymEnumLinesCallback_proc(ref SRCCODEINFO srcinfo, IntPtr UserContext)
		{
			//if (srcinfo.LineNumber == 335) System.Diagnostics.Debugger.Break();
			//if (srcinfo.LineNumber == 343) System.Diagnostics.Debugger.Break();

			if (srcinfo.LineNumber == 0xFEEFEE)
			{
				// skip these entries
				// I believe they mark the end of a set of linenumbers associated
				// with a particular method, but they don't appear to contain
				// valid line number info in any case.
			}
			else
			{
				try
				{
					// add the new line number and it's address
					// NOTE, this address is an IL Offset, not a native code offset

					var FileName = srcinfo.FileName.Split('\\');

					string Name;

					var i = FileName.GetUpperBound(0);

					if (i > 2)
					{
						Name = "...\\" + FileName[i - 2] + "\\" + FileName[i - 1] + "\\" + FileName[i];
					}
					else if (i > 1)
					{
						Name = "...\\" + FileName[i - 1] + "\\" + FileName[i];
					}
					else
					{
						Name = Path.GetFileName(srcinfo.FileName);
					}

					_alm.AddAddressToLine(srcinfo.LineNumber, srcinfo.Address, Name, srcinfo.obj);
				}
				catch (Exception ex)
				{
					Log.LogError(ex, "Unable to enum lines");
					// catch everything because we DO NOT
					// want to throw an exception from here!
				}
			}

			// Tell the caller we succeeded
			return true;
		}


		[DllImport("dbghelp.dll")]
		private static extern int SymInitialize(
		   int hProcess,
		   string UserSearchPath,
		   bool fInvadeProcess);



		[DllImport("dbghelp.dll")]
		private static extern Int64 SymLoadModuleEx(
			int hProcess, 
			int hFile, 
			string ImageName, 
			int ModuleName, 
			Int64 BaseOfDll, 
			int SizeOfDll, 
			int pModData, 
			int flags
		);

		[DllImport("dbghelp.dll")]
		private static extern bool SymGetModuleInfo(
			int hProcess,
			int dwAddr,
			ref IMAGEHLP_MODULE64 ModuleInfo
		);


		// I believe this is deprecatedin dbghelp 6.0 or later
		//Private Declare Function SymLoadModule Lib "dbghelp.dll" ( _
		//  ByVal hProcess As Integer, _
		//  ByVal hFile As Integer, _
		//  ByVal ImageName As String, _
		//  ByVal ModuleName As Integer, _
		//  ByVal BaseOfDll As Integer, _
		//  ByVal SizeOfDll As Integer) As Integer

		[DllImport("dbghelp.dll")]
		private static extern int SymCleanup(int hProcess);



		[DllImport("dbghelp.dll")]
		private static extern int UnDecorateSymbolName(
		   string DecoratedName ,
		   string UnDecoratedName ,
		   int UndecoratedLength ,
		   int Flags );


		[DllImport("dbghelp.dll")]
		private static extern int SymEnumSymbols(
		   int hProcess,
		   Int64 BaseOfDll,
		   int Mask,
		   PSYM_ENUMSYMBOLS_CALLBACK lpCallback,
		   int UserContext);


		[DllImport("dbghelp.dll")]
		private static extern int SymEnumLines(
		   int hProcess,
		   Int64 BaseOfDll,
		   int obj,
		   int File,
		   PSYM_ENUMLINES_CALLBACK lpCallback,
		   int UserContext);


		[DllImport("dbghelp.dll")]
		private static extern int SymEnumSourceLines(
		   int hProcess,
		   Int64 BaseOfDll,
		   int obj,
		   int File,
		   int Line,
		   int Flags,
		   PSYM_ENUMLINES_CALLBACK lpCallback,
		   int UserContext);


		// I believe this is deprecated in dbghelp 6.0+
		// Private Declare Function SymUnloadModule Lib "dbghelp.dll" ( _
		// ByVal hProcess As Integer, _
		// ByVal BaseOfDll As Integer) As Integer
																							

		[DllImport("dbghelp.dll")]
		private static extern bool SymUnloadModule64(int hProcess, Int64 BaseOfDll);

		#endregion


		#region " Exceptions"
		public class UnableToEnumerateSymbolsException : ApplicationException
		{ }


		public class UnableToEnumLinesException : ApplicationException
		{ }

		#endregion


		#region Constructors
		/// <summary>
		/// Constructor to setup this class to read a specific PDB
		/// </summary>
		/// <param name="FileName"></param>
		/// <remarks></remarks>
		public LineMapBuilder(string FileName)
		{
			this.Filename = FileName;
		}


		public LineMapBuilder(string FileName, string OutFileName)
		{
			this.Filename = FileName;
			this.OutFilename = OutFileName;
		}
		#endregion


		#region Public Members
		private string _Filename = "";
		/// <summary>
		/// Name of EXE/DLL file to process
		/// </summary>
		/// <value></value>
		/// <remarks></remarks>

		public string Filename
		{
			get
			{
				return _Filename;
			}

			set
			{
				if (!System.IO.File.Exists(value))
				{
					throw new FileNotFoundException("The file could not be found.", value);
				}
				_Filename = value;
			}
		}


		private string _OutFilename;
		/// <summary>
		/// Name of the output file
		/// </summary>
		/// <value></value>
		/// <remarks></remarks>
		public string OutFilename
		{
			get
			{
				if (string.IsNullOrEmpty(_OutFilename))
					return _Filename;
				else
					return _OutFilename;
			}
			set
			{
				_OutFilename = value;
			}
		}


		private bool _CreateMapReport = false;
		/// <summary>
		/// true to generate a Line map report
		/// this is mainly for Debugging purposes
		/// </summary>
		/// <value></value>
		/// <remarks></remarks>
		public bool CreateMapReport
		{
			get
			{
				return _CreateMapReport;
			}
			set
			{
				_CreateMapReport = value;
			}
		}


		/// <summary>
		/// Create a linemap file from the PDB for the given executable file
		/// Theoretically, you could read the LineMap info from the linemap file
		/// OR from a resource in the EXE/DLL, but I doubt you'd ever really want
		/// to. This is mainly for testing all the streaming functions
		/// </summary>
		/// <remarks></remarks>
		public void CreateLineMapFile()
		{
			// compress and encrypt the stream
			// technically, I should probably CLOSE and DISPOSE the streams
			// but this is just a quick and dirty tool

			var alm = GetAssemblyLineMap(this.Filename);

			var CompressedStream = CompressStream(alm.ToStream());
			var EncryptedStream = EncryptStream(CompressedStream);


			// swap out the below two lines to generate a linemap (lmp) file that is not compressed or encrypted
			StreamToFile(this.OutFilename + ".lmp", EncryptedStream);
			//pStreamToFile(Me.Filename & ".lmp", alm.ToStream);

			// write the report
			WriteReport(alm);
		}


		/// <summary>
		/// Inject a linemap resource into the given EXE/DLL file
		/// from the PDB for that file
		/// </summary>
		/// <remarks></remarks>
		public void CreateLineMapAPIResource()
		{
			// retrieve all the symbols
			var alm = GetAssemblyLineMap(this.Filename);

			// done one step at a time for debugging purposes
			var CompressedStream = CompressStream(alm.ToStream());

			var EncryptedStream = EncryptStream(CompressedStream);

			StreamToAPIResource(this.OutFilename,
							  Constants.ResTypeName,
							  Constants.ResName,
							  Constants.ResLang,
							  EncryptedStream);

			// write the report
			WriteReport(alm);
		}


		/// <summary>
		/// Inject a linemap .net resource into the given EXE/DLL file
		/// from the PDB for that file
		/// </summary>
		/// <remarks></remarks>
		public void CreateLineMapResource()
		{
			// retrieve all the symbols
			var alm = GetAssemblyLineMap(this.Filename);

			// done one step at a time for debugging purposes

			var CompressedStream = CompressStream(alm.ToStream());

			var EncryptedStream = EncryptStream(CompressedStream);

			StreamToResource(this.OutFilename,
							  Constants.ResName,
							  EncryptedStream);

			// write the report
			WriteReport(alm);
		}
		#endregion


		#region Private members
		/// <summary>
		/// Internal function to write out a line map report if asked to
		/// </summary>
		/// <remarks></remarks>
		private void WriteReport(AssemblyLineMap AssemblyLineMap)
		{
			//only write it if requested
			if (this.CreateMapReport)
			{
				Log.LogMessage("Creating symbol buffer report");

				using (var tw = new StreamWriter(this.Filename + ".linemapreport", false))
				{
					tw.Write(CreatePDBReport(AssemblyLineMap).ToString());
					tw.Flush();
				}
			}
		}


		/// <summary>
		/// Create a linemap report buffer from the PDB for the given executable file
		/// </summary>
		/// <remarks></remarks>
		private StringBuilder CreatePDBReport(AssemblyLineMap AssemblyLineMap)
		{

			var sb = new StringBuilder();

			sb.AppendLine("========");
			sb.AppendLine("SYMBOLS:");
			sb.AppendLine("========");
			sb.AppendLine(string.Format("   {0,-10}  {1,-10}  {2}  ", "Token", "Address", "Symbol"));
			sb.AppendLine(string.Format("   {0,-10}  {1,-10}  {2}  ", "-----", "-------", "------"));

			var symbols = AssemblyLineMap.Symbols.Values.Select(s =>
			{
				var lineex = AssemblyLineMap.AddressToLineMap.Where(l => l.Address == s.Address).FirstOrDefault();
				var name = lineex?.ObjectName;
				name = (!string.IsNullOrEmpty(name) ? name + "." : "") + s.Name;
				return new AssemblyLineMap.SymbolInfo(name, s.Address, s.Token);
			}).ToList();
			symbols.Sort((x, y) => x.Name.CompareTo(y.Name));

			foreach (var symbolEx in symbols)
			{
				sb.AppendLine(string.Format("   {0,-10:X}  {1,-10}  {2}", symbolEx.Token, symbolEx.Address, symbolEx.Name));
			}
			sb.AppendLine("========");
			sb.AppendLine("LINE NUMBERS:");
			sb.AppendLine("========");
			sb.AppendLine(string.Format("   {0,-10}  {1,-11}  {2,-10}  {3}", "Address", "Line number", "Token", "Symbol/FileName"));
			sb.AppendLine(string.Format("   {0,-10}  {1,-11}  {2,-10}  {3}", "-------", "-----------", "-----", "---------------"));

			//Order by line and then by address for reporting
			AssemblyLineMap.AddressToLineMap.Sort((x, y) =>
			{
				if (x.SourceFile.CompareTo(y.SourceFile) < 0)
					return -1;
				else if (x.SourceFile.CompareTo(y.SourceFile) > 0)
					return 1;
				else if (x.Line < y.Line)
					return -1;
				else if (x.Line > y.Line)
					return 1;
				else if (x.Address < y.Address)
					return -1;
				else if (x.Address > y.Address)
					return 1;
				else return 0;
			});

			// let the symbol run till the next transition is detected
			AssemblyLineMap.SymbolInfo sym = null;
			foreach (var lineex in AssemblyLineMap.AddressToLineMap)
			{
				// find the symbol for this line number
				foreach (var symbolex in AssemblyLineMap.Symbols.Values)
				{
					if (symbolex.Address == lineex.Address)
					{
						// found the symbol for this line
						sym = symbolex;
						break;
					}
				}
				//if (lineex.Line == 138) System.Diagnostics.Debugger.Break();
				var name = lineex.SourceFile + ":" + lineex.ObjectName + (sym != null ? "." + sym.Name : "");
				var token = sym != null ? sym.Token : 0;

				sb.AppendLine(string.Format("   {0,-10}  {1,-11}  {2,-10:X}  {3}", lineex.Address, lineex.Line, token, name));
			}

			sb.AppendLine("========");
			sb.AppendLine("NAMES:");
			sb.AppendLine("========");
			sb.AppendLine(string.Format("   {0,-10}  {1}", "Index", "Name"));
			sb.AppendLine(string.Format("   {0,-10}  {1}", "-----", "----"));
			for (int i = 0; i < AssemblyLineMap.Names.Count; i++)
			{
				sb.AppendLine(string.Format("   {0,-10}  {1}", i, AssemblyLineMap.Names[i]));
			}
			return sb;
		}


		/// <summary>
		/// Retrieve symbols and linenums and write them to a memory stream
		/// </summary>
		/// <param name="FileName"></param>
		/// <returns></returns>
		private AssemblyLineMap GetAssemblyLineMap(string FileName)
		{
			// create a new map to capture symbols and line info with
			_alm = Utilities.Bookkeeping.AssemblyLineMaps.Add(FileName);

			if (!System.IO.File.Exists(FileName))
			{
				throw new FileNotFoundException("The file could not be found.", FileName);
			}

			var hProcess = System.Diagnostics.Process.GetCurrentProcess().Id;
			Int64 dwModuleBase = 0;

			// clear the map
			_alm.Clear();

			try
			{

				if (SymInitialize(hProcess, "", false) != 0)
				{
					dwModuleBase = SymLoadModuleEx(hProcess, 0, FileName, 0, 0, 0, 0, 0);

					if (dwModuleBase != 0)
					{
						// this appearently is required in some cases where the moduleinfo load may be deferred
						IMAGEHLP_MODULE64 moduleInfo = IMAGEHLP_MODULE64.Create();
						var r = SymGetModuleInfo(hProcess, 0, ref moduleInfo);

						// Enumerate all the symbol names 
						var rEnumSymbolsDelegate = new PSYM_ENUMSYMBOLS_CALLBACK(SymEnumSymbolsCallback_proc);
						if (SymEnumSymbols(hProcess, dwModuleBase, 0, rEnumSymbolsDelegate, 0) == 0)
						{
							// unable to retrieve the symbol list
							throw new UnableToEnumerateSymbolsException();
						}

						// now enum all the source lines and their respective addresses
						var pSymEnumLinesCallback_proc = new PSYM_ENUMLINES_CALLBACK(SymEnumLinesCallback_proc);

						if (SymEnumSourceLines(hProcess, dwModuleBase, 0, 0, 0, 0, pSymEnumLinesCallback_proc, 0) == 0)
						{
							// unable to retrieve the line number list
							throw new UnableToEnumLinesException();
						}
					}
				}
			}
			catch (Exception ex)
			{
				// return return vars
				Log.LogError(ex, "Unable to retrieve symbols.");
				_alm.Clear();

				// and rethrow
				throw;
			}
			finally
			{
				Log.LogMessage("Retrieved {0} symbols", _alm.Symbols.Count);

				Log.LogMessage("Retrieved {0} lines", _alm.AddressToLineMap.Count) ;

				Log.LogMessage("Retrieved {0} strings", _alm.Names.Count);
	
				// release the module
				if (dwModuleBase != 0) SymUnloadModule64(hProcess, dwModuleBase);

				// can clean up the dbghelp system
				SymCleanup(hProcess);
			}
			return _alm;
		}
		private AssemblyLineMap _alm;


		/// <summary>
		/// Given a Filename and memorystream, write the stream to the file
		/// </summary>
		/// <param name="Filename"></param>
		/// <param name="MemoryStream"></param>
		/// <remarks></remarks>
		private void StreamToFile(String Filename, MemoryStream MemoryStream)
		{
			Log.LogMessage("Writing symbol buffer to file");

			using (var fileStream = new System.IO.FileStream(Filename, FileMode.Create))
			{ 
				MemoryStream.WriteTo(fileStream);
				fileStream.Flush();
			}
		}


		/// <summary>
		/// Write out a memory stream to a named resource in the given 
		/// WIN32 format executable file (iether EXE or DLL)
		/// </summary>
		/// <param name="Filename"></param>
		/// <param name="ResourceName"></param>
		/// <param name="MemoryStream"></param>
		/// <remarks></remarks>
		private void StreamToAPIResource(string filename, object ResourceType, object ResourceName, Int16 ResourceLanguage, MemoryStream memoryStream)
		{
			var ResWriter = new ResourceWriter();

			// the target file has to exist
			if (this.Filename != filename)
			{
				File.Copy(this.Filename, filename, true);
				File.SetAttributes(filename, FileAttributes.Normal);
			}

			// convert memorystream to byte array and write out to 
			// the linemap resource
			Log.LogMessage("Writing symbol buffer to resource");

			ResWriter.FileName = filename;
			var buf = memoryStream.ToArray();
			ResWriter.Update(ResourceType, ResourceName, ResourceLanguage, ref buf);
		}


		/// <summary>
		/// Write out a memory stream to a named resource in the given 
		/// WIN32 format executable file (iether EXE or DLL)
		/// </summary>
		/// <param name="OutFilename"></param>
		/// <param name="ResourceName"></param>
		/// <param name="MemoryStream"></param>
		/// <remarks></remarks>
		private void StreamToResource(string OutFilename, string ResourceName, MemoryStream memoryStream)
		{
			var ResWriter = new ResourceWriterCecil();

			// the target file has to 
			if (this.Filename != OutFilename)
			{
				System.IO.File.Copy(this.Filename, OutFilename);
			}

			// convert memorystream to byte array and write out to 
			// the linemap resource

			Log.LogMessage("Writing symbol buffer to resource");

			ResWriter.FileName = this.Filename;
			ResWriter.Add(ResourceName, memoryStream.ToArray());

			ResWriter.Save(OutFilename);
		}


		/// <summary>
		/// Given an input stream, compress it into an output stream
		/// </summary>
		/// <param name="UncompressedStream"></param>
		/// <returns></returns>
		private MemoryStream CompressStream(MemoryStream UncompressedStream)
		{
			var CompressedStream = new System.IO.MemoryStream();

			// note, the LeaveOpen parm MUST BE true into order to read the memory stream afterwards!
			Log.LogMessage("Compressing symbol buffer");

			using (var GZip = new System.IO.Compression.GZipStream(CompressedStream, System.IO.Compression.CompressionMode.Compress, true))
			{
				GZip.Write(UncompressedStream.ToArray(), 0, (int)UncompressedStream.Length);
			}
			return CompressedStream;
		}


		/// <summary>
		/// Given a stream, encrypt it into an output stream
		/// </summary>
		/// <param name="CompressedStream"></param>
		/// <returns></returns>
		private MemoryStream EncryptStream(MemoryStream CompressedStream)
		{
			var Enc = new System.Security.Cryptography.RijndaelManaged();

			Log.LogMessage("Encrypting symbol buffer");
			// setup our encryption key
			Enc.KeySize = 256;
			// KEY is 32 byte array
			Enc.Key = Constants.ENCKEY;
			// IV is 16 byte array
			Enc.IV = Constants.ENCIV;

			var EncryptedStream = new System.IO.MemoryStream();

			var cryptoStream = new System.Security.Cryptography.CryptoStream(EncryptedStream, Enc.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write);
			cryptoStream.Write(CompressedStream.ToArray(), 0, (int)CompressedStream.Length);
			cryptoStream.FlushFinalBlock();

			return EncryptedStream;
		}
		#endregion
	}
}
