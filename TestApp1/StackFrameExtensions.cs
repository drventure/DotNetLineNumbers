using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TestApp1
{
	class StackFrameExtensions
	{
		/// <summary>
		/// turns a single stack frame object into an informative string
		/// </summary>
		/// <param name="FrameNum"></param>
		/// <param name="sf"></param>
		/// <returns></returns>
		private string StackFrameToString(int FrameNum, StackFrame sf)
		{
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			int intParam = 0;
			MemberInfo mi = sf.GetMethod();

			var _with1 = sb;
			//-- build method name
			_with1.Append("   ");
			string MethodName = mi.DeclaringType.Namespace + "." + mi.DeclaringType.Name + "." + mi.Name;
			_with1.Append(MethodName);
			//If FrameNum = 1 Then rCachedErr.Method = MethodName

			//-- build method params
			ParameterInfo[] objParameters = sf.GetMethod().GetParameters();
			_with1.Append("(");
			intParam = 0;
			foreach (ParameterInfo objParameter in objParameters)
			{
				intParam += 1;
				if (intParam > 1)
					_with1.Append(", ");
				_with1.Append(objParameter.Name);
				_with1.Append(" As ");
				_with1.Append(objParameter.ParameterType.Name);
			}
			_with1.Append(")");
			_with1.Append(Environment.NewLine);

			//-- if source code is available, append location info
			_with1.Append("       ");
			//if (strings.InStr(1, Command, "/uselinemap", CompareMethod.Text) == 0 && string.IsNullOrEmpty(Interaction.Environ("uselinemap")) 
			if (sf.GetFileName() != null && sf.GetFileName().Length != 0)
			{
				//---- the PDB appears to be available, since the above elements are 
				//     not blank, so just use it's information

				_with1.Append(System.IO.Path.GetFileName(sf.GetFileName()));
				dynamic Line = sf.GetFileLineNumber();
				if (Line != 0)
				{
					_with1.Append(": line ");
					_with1.Append(string.Format("{0:#0000}", Line));
				}
				dynamic col = sf.GetFileColumnNumber();
				if (col != 0)
				{
					_with1.Append(", col ");
					_with1.Append(string.Format("{0:#00}", sf.GetFileColumnNumber()));
				}
				//-- if IL is available, append IL location info
				if (sf.GetILOffset() != StackFrame.OFFSET_UNKNOWN)
				{
					_with1.Append(", IL ");
					_with1.Append(string.Format("{0:#0000}", sf.GetILOffset()));
				}
			}
			else
			{
				//---- the PDB is not available, so attempt to retrieve 
				//     any embedded linemap information
				string FileName = System.IO.Path.GetFileName(ParentAssembly.CodeBase);
				//If FrameNum = 1 Then rCachedErr.FileName = FileName
				_with1.Append(FileName);
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
						_with1.Append(": Source File - ");
						_with1.Append(SourceFile);
						_with1.Append(": Line ");
						_with1.Append(string.Format("{0:#0000}", Line));
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
					dynamic IL = sf.GetILOffset();
					if (IL != StackFrame.OFFSET_UNKNOWN)
					{
						_with1.Append(": IL ");
						_with1.Append(string.Format("{0:#00000}", IL));
					}
				}
			}
			_with1.Append(Environment.NewLine);
			return sb.ToString();
		}


		private Assembly _parentAssembly = null;
		/// <summary>
		/// Retrieve the root assembly of the executing assembly
		/// </summary>
		/// <returns></returns>
		private Assembly ParentAssembly
		{
			get
			{
				if (_parentAssembly == null)
				{ 
					if (Assembly.GetEntryAssembly() != null)
					{
						_parentAssembly = System.Reflection.Assembly.GetCallingAssembly();
					}
					else
					{
						_parentAssembly = System.Reflection.Assembly.GetEntryAssembly();
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
		private void MapStackFrameToSourceLine(StackFrame sf, ref int Line, ref string SourceFile)
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
			dynamic alm = LineMap.AssemblyLineMaps[sf.GetMethod().DeclaringType.Assembly.CodeBase];

			//---- does the symbols list contain the metadata token for this method?
			MemberInfo mi = sf.GetMethod();
			//---- Don't call this mdtoken or PostSharp will barf on it! Jeez
			long mdtokn = mi.MetadataToken;
			if (!alm.Symbols.ContainsKey(mdtokn))
				return;

			//---- all is good so get the line offset (as close as possible, considering any optimizations that
			//     might be in effect)
			dynamic ILOffset = sf.GetILOffset();
			if (ILOffset != StackFrame.OFFSET_UNKNOWN)
			{
				Int64 Addr = alm.Symbols(mdtokn).Address + ILOffset;

				//---- now start hunting down the line number entry
				//     use a simple search. LINQ might make this easier
				//     but I'm not sure how. Also, a binary search would be faster
				//     but this isn't something that's really performance dependent
				int i = 1;
				for (i = alm.AddressToLineMap.Count - 1; i >= 0; i += -1)
				{
					if (alm.AddressToLineMap(i).Address <= Addr)
					{
						break;
					}
				}
				//---- since the address may end up between line numbers,
				//     always return the line num found
				//     even if it's not an exact match
				Line = alm.AddressToLineMap(i).Line;
				SourceFile = alm.Names(alm.AddressToLineMap(i).SourceFileIndex);
			}
			else
			{
				return;
			}
		}
	}
}
