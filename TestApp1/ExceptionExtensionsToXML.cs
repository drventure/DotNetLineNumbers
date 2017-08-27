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
