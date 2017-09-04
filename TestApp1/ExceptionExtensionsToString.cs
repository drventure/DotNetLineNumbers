using System;
using System.Collections;
using System.Linq;
using System.Text;

using ExceptionExtensions;
using ExceptionExtensions.Internal;
using System.Diagnostics;
using System.Reflection;

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
	}


	public struct ExceptionOptions
	{
		public ExceptionOptions(int indentSpaces = 4, bool omitNullProperties = true)
		{
			this.CurrentIndentLevel = 0;
			this.IndentSpaces = indentSpaces;
			this.OmitNullProperties = omitNullProperties;
		}

		public static readonly ExceptionOptions Default = new ExceptionOptions()
		{
			CurrentIndentLevel = 0,
			IndentSpaces = 4,
			OmitNullProperties = true
		};


		internal ExceptionOptions(ExceptionOptions options, int currentIndent)
		{
			this.CurrentIndentLevel = currentIndent;
			this.IndentSpaces = options.IndentSpaces;
			this.OmitNullProperties = options.OmitNullProperties;
		}

		internal string Indent { get { return new string(' ', this.IndentSpaces * this.CurrentIndentLevel); } }

		internal int CurrentIndentLevel { get; set; }

		public int IndentSpaces { get; set; }

		public bool OmitNullProperties { get; set; }
	}
}


/// <summary>
/// Seperate namespace to allow extension of the StringBuilder and Stacktrace without polluting it elsewhere in the host app
/// </summary>
namespace ExceptionExtensions.Internal
{
	public static class StringBuilderExtensions
	{
		public static void AppendLine(this StringBuilder sb, string caption, Func<string> getValue)
		{
			sb.Append(caption);
			try
			{
				string text = getValue();
				sb.AppendLine(text);
			}
			catch (Exception ex)
			{
				sb.AppendLine(ex.Message);
			}
		}


		public static void AppendLine(this StringBuilder sb, Func<string> getValue)
		{
			sb.AppendLine(string.Empty, getValue);
		}


		public static void AppendValue(this StringBuilder sb, string propertyName, object value, ExceptionOptions options)
		{
			if (value is Exception)
			{
				var innerException = (Exception)value;
				sb.AppendException(propertyName, innerException, options);
			}
			else if (value is StackTrace)
			{
				sb.Append(string.Format("{0}{1}:", options.Indent, propertyName).PadRight(23));
				sb.AppendLine(((StackTrace)value).ToString(true));
			}
			else if (value is IEnumerable && !(value is string))
			{
				var collection = (IEnumerable)value;
				if (collection.GetEnumerator().MoveNext())
				{
					sb.AppendCollection(propertyName, collection, options);
				}
			}
			else
			{
				sb.Append(string.Format("{0}{1}:", options.Indent, propertyName).PadRight(23));
				if (value is DictionaryEntry)
				{
					DictionaryEntry dictionaryEntry = (DictionaryEntry)value;
					sb.AppendLine(string.Format("{0} : {1}", dictionaryEntry.Key, dictionaryEntry.Value));
				}
				else if (propertyName == "HResult")
				{
					sb.AppendLine(string.Format("0x{0:X}", (int)value));
				}
				else
				{
					sb.AppendLine(value.ToString());
				}
			}
		}


		public static void AppendCollection(this StringBuilder sb, string propertyName, IEnumerable collection, ExceptionOptions options)
		{
			sb.AppendLine(string.Format("{0}{1}:", options.Indent, propertyName).PadRight(23));

			var innerOptions = new ExceptionOptions(options, options.CurrentIndentLevel + 1);

			var i = 0;
			foreach (var item in collection)
			{
				var innerPropertyName = string.Format("[{0}]", i);

				if (item is Exception)
				{
					var innerException = (Exception)item;
					sb.AppendException(innerPropertyName, innerException, innerOptions);
				}
				else
				{
					sb.AppendValue(innerPropertyName, item, innerOptions);
				}

				++i;
			}
		}


		public static void AppendException(this StringBuilder sb, string propertyName, Exception ex, ExceptionOptions options)
		{
			var innerExceptionString = ex.ToExpandedString(new ExceptionOptions(options, options.CurrentIndentLevel + 1));

			sb.AppendLine(string.Format("{0}{1}: ", options.Indent, propertyName).PadRight(23));
			sb.AppendLine(innerExceptionString);
		}
	}
}
