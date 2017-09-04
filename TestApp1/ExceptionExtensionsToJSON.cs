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
# endregion
 
 using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;


/// <summary>
/// This file specifically implements extension methods on the Exception class
/// to convert exceptions to JSON format
/// 
/// It does this by converting the Exception to a SerializableException and 
/// then using standard JSON serialization functions on it
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
		public static string ToJSON(this Exception ex)
		{
			return new SerializableException(ex).ToJSON();
		}


		/// <summary>
		/// This method provides the default ToString rendering
		/// </summary>
		/// <returns></returns>
		private static string ToJSON(this SerializableException sx)
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;

			try
			{
				// write JSON in human readable, indented format
				string json;
				using (MemoryStream ms = new MemoryStream())
				{
					using (var writer = JsonReaderWriterFactory.CreateJsonWriter(ms, Encoding.UTF8, true, true, "  "))
					{
						DataContractJsonSerializer js = new DataContractJsonSerializer(
							typeof(SerializableException), 
							new DataContractJsonSerializerSettings()
							{
								UseSimpleDictionaryFormat = true
							});
						js.WriteObject(writer, sx);
						writer.Flush();
						ms.Position = 0;
						using (StreamReader sr = new StreamReader(ms))
						{
							json = sr.ReadToEnd();
						}
					}
				}
				return json;
			}
			catch (Exception exception)
			{
				Debug.WriteLine(exception.ToString());
				return null;
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
	}
}
