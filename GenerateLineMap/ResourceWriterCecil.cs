#region MIT License
/*
    MIT License

    Copyright (c) 2016 Darin Higgins

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

using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using Mono.Cecil;
using System.Resources;


/// <summary>
/// 
/// </summary>
/// <remarks></remarks>
public class ResourceWriterCecil
{
	private bool _Inited;
	private AssemblyDefinition _Asm;
	private Mono.Collections.Generic.Collection<Resource> _Resources;


	public ResourceWriterCecil()
	{
	}


	public ResourceWriterCecil(string AssemblyFileName)
	{
		this.FileName = AssemblyFileName;
	}


	/// <summary>
	/// The filename of the assembly to modify
	/// </summary>
	/// <value></value>
	/// <returns></returns>
	/// <remarks></remarks>
	public string FileName
	{
		get { return rFilename; }
		set { rFilename = value; }
	}

	private string rFilename;

	private void InitAssembly()
	{
		if (_Inited == false & Strings.Len(this.FileName) > 0)
		{
			_Inited = true;

			// load up the assembly to be modified
			_Asm = AssemblyDefinition.ReadAssembly(this.FileName);

			//Gets all types which are declared in the Main Module of the asm to be modified
			_Resources = _Asm.MainModule.Resources;
		}
	}


	/// <summary>
	/// Adds a given resourcename and data to the app resources in the target assembly
	/// </summary>
	/// <param name="ResourceName"></param>
	/// <param name="ResourceData"></param>
	/// <remarks></remarks>
	public void Add(string ResourceName, byte[] ResourceData)
	{
		// make sure the writer is initialized
		InitAssembly();

		// have to enumerate this way
		for (var x = 0; x <= _Resources.Count - 1; x++)
		{
			var res = _Resources[x];
			if (res.Name.Contains(".Resources.resources"))
			{
				// Have to assume this is the root application's .net resources.
				// That might not be the case though.

				// cast as embeded resource to get at the data
				var EmbededResource = (Mono.Cecil.EmbeddedResource)res;

				// a Resource reader is required to read the resource data
				var ResReader = new ResourceReader(new MemoryStream(EmbededResource.GetResourceData()));

				// Use this output stream to capture all the resource data from the 
				// existing resource block, so we can add the new resource into it
				var MemStreamOut = new MemoryStream();
				var ResWriter = new System.Resources.ResourceWriter(MemStreamOut);
				var ResEnumerator = ResReader.GetEnumerator();
				byte[] resdata = null;
				while (ResEnumerator.MoveNext())
				{
					var resname = (string)ResEnumerator.Key;
					string restype = "";
					// if we come across a resource named the same as the one
					// we're about to add, skip it
					if (Strings.StrComp(resname, ResourceName, CompareMethod.Text) != 0)
					{
						ResReader.GetResourceData(resname, out restype, out resdata);
						ResWriter.AddResourceData(resname, restype, resdata);
					}
				}

				// add the new resource data here
				ResWriter.AddResourceData(ResourceName, "ResourceTypeCode.ByteArray", ResourceData);
				// gotta call this to render the memory stream
				ResWriter.Generate();

				// update the resource
				var buf = MemStreamOut.ToArray();
				var NewEmbedRes = new EmbeddedResource(res.Name, res.Attributes, buf);
				_Resources.Remove(res);
				_Resources.Add(NewEmbedRes);
				// gotta bail out, there can't be 2 embedded resource chunks, right?
				break; // TODO: might not be correct. Was : Exit For
			}
		}
	}


	/// <summary>
	/// Save all changes now.
	/// </summary>
	/// <remarks></remarks>
	public void Save(string OutFileName = "")
	{
		if (string.IsNullOrEmpty(OutFileName))
			OutFileName = this.FileName;
		_Asm.Write(OutFileName);
	}
}
