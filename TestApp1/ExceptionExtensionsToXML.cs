using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;


/// <summary>
/// This file specifically implements extension methods on the Exception class
/// to convert exceptions ToString()
/// 
/// It does this by retrieving the exception as a SerializableException
/// and then rendering that ToString();
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
			var xmlSerializer = new XmlSerializer(sx.GetType());

			using (var ms = new MemoryStream())
			{
				using (var xw = XmlWriter.Create(ms,
					new XmlWriterSettings()
					{
						Encoding = new UTF8Encoding(false),
						Indent = true,
						NewLineOnAttributes = true,
					}))
				{
					xmlSerializer.Serialize(xw, sx);
					return Encoding.UTF8.GetString(ms.ToArray());
				}
			}
		}
	}
}
