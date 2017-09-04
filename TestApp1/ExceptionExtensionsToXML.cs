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
using System.IO;
using System.Runtime.Serialization;
using System.Xml;


/// <summary>
/// This file specifically implements extension methods on the Exception class
/// to convert exceptions to XML format
/// 
/// It does this by converting the Exception to a SerializableException and 
/// then using standard XML serialization functions on it
/// </summary>
namespace ExceptionExtensions
{
	public static partial class ExceptionExtensions
	{
		/// <summary>
		/// translate exception object to string, with additional system info
		/// </summary>
		/// <param name="ex"></param>
		/// <returns></returns>
		public static string ToXML(this Exception ex)
		{
			return new SerializableException(ex).ToXML();
		}


		/// <summary>
		/// This method provides the default ToString rendering
		/// </summary>
		/// <returns></returns>
		private static string ToXML(this SerializableException sx)
		{
			var serializer = new DataContractSerializer(sx.GetType());
			using (var sw = new StringWriter())
			{
				using (var writer = new XmlTextWriter(sw))
				{
					writer.Formatting = Formatting.Indented; // indent the Xml so it's human readable
					serializer.WriteObject(writer, sx);
					writer.Flush();
					return sw.ToString();
				}
			}
		}
	}
}
