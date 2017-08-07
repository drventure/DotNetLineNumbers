using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml.Serialization;

using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using ExceptionExtensions;


namespace GenerateLineMapUnitTests
{
	[TestClass]
	public class XMLSerializerTests
	{
		private readonly ConcurrentDictionary<Type, XmlSerializer> XmlSerializers =
			new ConcurrentDictionary<Type, XmlSerializer>();


		/// <summary>
		/// Serializes the object using XML format.
		/// </summary>
		private string SerializeAsXml(object obj)
		{
			var type = obj.GetType();
			var xmlSerializer = GetXmlSerializer(type);

			using (var memStream = new MemoryStream())
			{
				xmlSerializer.Serialize(memStream, obj);
				return Encoding.Default.GetString(memStream.ToArray());
			}
		}

		private XmlSerializer GetXmlSerializer(Type type)
		{
			// gets the xml serializer from the concurrent dictionary, if it doesn't exist
			// then add one for the specified type
			return XmlSerializers.GetOrAdd(type, t => new XmlSerializer(t));
		}


		[TestMethod]
		public void SerializeToXMLTest()
		{
			var ex = new SerializableException();
			ex.Properties["Type"] = "System.Exception";
			ex.Properties["DateAndTime"] = DateTime.Now.ToString();
			ex.Properties["Message"] = "Just a Test";

			var ex2 = new SerializableException();
			ex2.Properties["Type"] = "System.Exception2";
			ex2.Properties["DateAndTime"] = DateTime.Now.ToString();
			ex2.Properties["Message"] = "Just a Test2";

			ex.InnerException = ex2;

			var ex3 = new SerializableException();
			ex3.Properties["Type"] = "System.Exception3";
			ex3.Properties["DateAndTime"] = DateTime.Now.ToString();
			ex3.Properties["Message"] = "Just a Test3";

			ex2.InnerException = ex3;

			var buf = SerializeAsXml(ex);
			Debug.WriteLine(buf);
		}


		[TestMethod]
		public void SerializableExceptionToXMLTest()
		{
			try
			{
				throw new FileNotFoundException("3rd File was not found", "ThirdFileNotFound.txt",
					new FileNotFoundException("2nd File was not found", "SecondFileNotFound.txt",
					new FileNotFoundException("1st File was not found", "FirstFileNotFound.txt")));
			}
			catch (Exception ex)
			{
				var buf = ex.ToXML();
				Debug.WriteLine(buf);
				buf.Should().Contain("1st File was not found");
			}
		}

	}
}
