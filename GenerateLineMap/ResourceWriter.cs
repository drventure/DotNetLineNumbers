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
using System.Linq;
using System.Runtime.InteropServices;


namespace GenerateLineMap
{
	/// <summary>
	/// Class to handle updating resources in a Windows DLL or EXE
	/// (c) 2008-2010 Darin Higgins All Rights Reserved
	/// </summary>
	/// <remarks></remarks>
	/// <editHistory></editHistory>
	public class ResourceWriter : IDisposable
	{

		#region " Enumerations"
		public enum ResTypes
		{
			RT_ACCELERATOR = 9,
			RT_ANICURSOR = (21),

			RT_ANIICON = (22),

			RT_BITMAP = 2,
			RT_CURSOR = 1,
			RT_DIALOG = 5,
			RT_DLGINCLUDE = 17,

			RT_FONT = 8,
			RT_FONTDIR = 7,
			RT_HTML = 23,

			RT_ICON = 3,
			RT_MENU = 4,
			RT_MESSAGETABLE = 11,

			RT_PLUGPLAY = 19,

			RT_RCDATA = 10,
			RT_STRING = 6,
			RT_VERSION = 16,

			RT_VXD = 20,

			DIFFERENCE = 11,

			RT_GROUP_CURSOR = (RT_CURSOR + DIFFERENCE),

			RT_GROUP_ICON = (RT_ICON + DIFFERENCE),
		}
		#endregion


		#region " Exceptions"
		public class ResWriteCantOpenException : ApplicationException
		{
			public ResWriteCantOpenException(string Filename, int ErrCode) :
				base("Unable to file " + Filename + "   Code: " + ErrCode.ToString())
			{ }
		}
		public class ResWriteCantUpdateException : ApplicationException
		{
			public ResWriteCantUpdateException(int ErrCode) :
				base("Problem updating the resource. Code: " + ErrCode.ToString())
			{ }
		}
		public class ResWriteCantDeleteException : ApplicationException
		{
			public ResWriteCantDeleteException(int ErrCode) :
				base("Problem updating the resource. Code: " + ErrCode.ToString())
			{ }
		}
		public class ResWriteCantEndException : ApplicationException
		{
			public ResWriteCantEndException(int ErrCode) :
				base("Problem finalizing resource modifications. Code: " + ErrCode.ToString())
			{ }
		}

		#endregion



		public string FileName
		{
			get
			{
				return rFilename;
			}
			set
			{
				rFilename = value;
			}
		}
		private string rFilename;




		// '' <summary>
		// '' Add and Update are really the same thing, they both
		// '' add and update resources, as appropriate
		// '' </summary>
		// '' <param name="ResType"></param>
		// '' <param name="ResName"></param>
		// '' <param name="wLanguage"></param>
		// '' <param name="NewValue"></param>
		// '' <remarks></remarks>
		public void Add(object ResType, object ResName, short wLanguage, ref List<byte> NewValue)
		{
			string buf = "";
			NewValue.ForEach(c => { buf += c.ToString(); });
			this.Update(ResType, ResName, wLanguage, ref buf);
		}


		// '' <summary>
		// '' Update or insert a resource into a WinPE image.
		// '' This version accepts a string buffer
		// '' and converts it to a byte buffer
		// '' </summary>
		// '' <param name="ResType"></param>
		// '' <param name="ResName"></param>
		// '' <param name="NewValue"></param>
		// '' <param name="wLanguage"></param>
		// '' <remarks></remarks>
		public void Update(object ResType, object ResName, short wLanguage, ref string NewValue)
		{
			// make sure there's something to do
			if ((NewValue.Length == 0))
			{
				return;
			}

			// convert raw string to byte buffer
			List<byte> buf = new List<byte>();
			foreach (char c in NewValue.ToCharArray())
			{
				buf.Add((byte)c);
			}

			// and update the resource
			this.Update(ResType, ResName, wLanguage, ref buf);
		}


		public void Update(object ResType, object ResName, short wLanguage, ref byte[] NewValue)
		{
			// make sure there's something to do
			if ((NewValue.Length == 0))
			{
				return;
			}

			var buf = NewValue.ToList();

			// and update the resource
			this.Update(ResType, ResName, wLanguage, ref buf);
		}


		// '' <summary>
		// '' Update a resource using a raw byte buffer
		// '' </summary>
		// '' <param name="ResType"></param>
		// '' <param name="ResName"></param>
		// '' <param name="wLanguage"></param>
		// '' <param name="NewValue"></param>
		// '' <remarks></remarks>
		public void Update(object ResType, object ResName, short wLanguage, ref List<byte> NewValue)
		{
			int Result;
			IntPtr hUpdate = ResourceAPIs.BeginUpdateResource(this.FileName, 0);
			if (hUpdate == IntPtr.Zero)
			{
				throw new ResWriteCantOpenException(this.FileName, Marshal.GetLastWin32Error());
			}

			// guarantee data is aligned to 4 byte boundary
			//  not sure if this is strictly necessary, but it's documented
			int l = NewValue.Count;
			int Mod = (l % 4);
			if ((Mod > 0))
			{
				l += 4 - Mod;
			}

			byte[] buf = new byte[l];
			NewValue.ToArray().CopyTo(buf, 0);

			// important note, The Typename and Resourcename are
			// written as UPPER CASE. I'm not sure why but this appears to 
			// be required for the FindResource and FindResourceEx API calls to 
			// be able to work properly. Otherwise, they'll never find the resource
			// entries...
			if (ResType is string)
			{
				// Restype is a resource type name
				if (ResName is string)
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, ResType.ToString().ToUpper(), ResName.ToString().ToUpper(), wLanguage, buf, l);
				}
				else
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, ResType.ToString().ToUpper(), int.Parse(ResName.ToString()), wLanguage, buf, l);
				}

			}
			else
			{
				// Restype is a numeric resource ID
				if (ResName is string)
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, int.Parse(ResType.ToString()), ResName.ToString().ToUpper(), wLanguage, buf, l);
				}
				else
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, int.Parse(ResType.ToString()), int.Parse(ResName.ToString()), wLanguage, buf, l);
				}

			}

			if ((Result == 0))
			{
				throw new ResWriteCantUpdateException(Marshal.GetLastWin32Error());
			}

			Result = ResourceAPIs.EndUpdateResource(hUpdate, 0);
			if ((Result == 0))
			{
				throw new ResWriteCantEndException(Marshal.GetLastWin32Error());
			}
		}


		/// <summary>
		/// Delete a resource from the WinPE image.
		/// Important note: If the target resource was added via ADD or UPDATE
		/// in this class (and via BeginUpdateResource in general), this process
		/// fails with an error 87 (bad parameter).
		/// I have no idea why at this point.
		/// If the resource in question already existed within the file (defined at
		/// compile time, for instance), then there usually is no problem deleting it.
		/// </summary>
		/// <param name="ResType"></param>
		/// <param name="ResName"></param>
		/// <param name="wLanguage"></param>
		/// <remarks></remarks>
		public void Delete(object ResType, object ResName, short wLanguage = 0)
		{
			int Result;
			// Warning!!! Optional parameters not supported
			IntPtr hUpdate = ResourceAPIs.BeginUpdateResource(this.FileName, 0);
			if (hUpdate.ToInt32() == 0)
			{
				throw new ResWriteCantOpenException(this.FileName, Marshal.GetLastWin32Error());
			}

			if (ResType is string)
			{
				// Restype is a resource name
				if (ResName is string)
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, ResType.ToString().ToUpper(), ResName.ToString().ToUpper(), wLanguage, IntPtr.Zero, 0);
				}
				else
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, ResType.ToString().ToUpper(), int.Parse(ResName.ToString()), wLanguage, IntPtr.Zero, 0);
				}

			}
			else
			{
				// Restype is a numeric resource ID
				if (ResName is string)
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, int.Parse(ResType.ToString()), ResName.ToString().ToUpper(), wLanguage, IntPtr.Zero, 0);
				}
				else
				{
					Result = ResourceAPIs.UpdateResource(hUpdate, int.Parse(ResType.ToString()), int.Parse(ResName.ToString()), wLanguage, IntPtr.Zero, 0);
				}

			}

			if (Result == 0)
			{
				throw new ResWriteCantDeleteException(Marshal.GetLastWin32Error());
			}

			Result = ResourceAPIs.EndUpdateResource(hUpdate, 0);
			if (Result == 0)
			{
				throw new ResWriteCantEndException(Marshal.GetLastWin32Error());
			}

		}

		public void Dispose()
		{
			// Nothing really to dispose of

		}
	}


	#region " Resource API Declarations"
	/// <summary>
	/// this is an internal module for Resource API functions I use above
	/// to write the linemap into the resources of an exe/dll
	/// </summary>
	/// <remarks></remarks>
	public static class ResourceAPIs
	{
		[DllImport("KERNEL32.DLL", EntryPoint = "BeginUpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern IntPtr BeginUpdateResource([MarshalAs(UnmanagedType.LPWStr)]
string pFileName, Int32 bDeleteExistingResources);


		[DllImport("KERNEL32.DLL", EntryPoint = "EndUpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 EndUpdateResource(IntPtr hUpdate, Int32 fDiscard);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, Int32 lpType, [MarshalAs(UnmanagedType.LPWStr)]
string lpName, Int16 wLanguage, byte[] lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, [MarshalAs(UnmanagedType.LPWStr)]
string lpType, [MarshalAs(UnmanagedType.LPWStr)]
string lpName, Int16 wLanguage, byte[] lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, [MarshalAs(UnmanagedType.LPWStr)]
string lpType, Int32 lpName, Int16 wLanguage, byte[] lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, Int32 lpType, Int32 lpName, Int16 wLanguage, byte[] lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, Int32 lpType, [MarshalAs(UnmanagedType.LPWStr)]
string lpName, Int16 wLanguage, IntPtr lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, [MarshalAs(UnmanagedType.LPWStr)]
string lpType, [MarshalAs(UnmanagedType.LPWStr)]
string lpName, Int16 wLanguage, IntPtr lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, [MarshalAs(UnmanagedType.LPWStr)]
string lpType, Int32 lpName, Int16 wLanguage, IntPtr lpData, Int32 cbData);


		[DllImport("KERNEL32.DLL", EntryPoint = "UpdateResourceW", SetLastError = true, CharSet = CharSet.Unicode)]
		public static extern Int32 UpdateResource(IntPtr hUpdate, Int32 lpType, Int32 lpName, Int16 wLanguage, IntPtr lpData, Int32 cbData);
	}
	#endregion
}