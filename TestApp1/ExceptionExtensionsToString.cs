using System;
using System.Collections;
using System.Linq;
using System.Text;

using ExceptionExtensions;
using ExceptionToString.Internal;


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
		public static string ToString(this Exception ex, bool extended)
		{
			return ex.ToString(ExceptionOptions.Default);
		}


		/// <summary>
		/// translate exception object to string, with additional system info
		/// </summary>
		/// <param name="ex"></param>
		/// <returns></returns>
		public static string ToString(this Exception ex, ExceptionOptions options)
		{
			return new SerializableException(ex).ToString(options);
		}


		/// <summary>
		/// This method provides the default ToString rendering
		/// </summary>
		/// <returns></returns>
		private static string ToString(this SerializableException sx, ExceptionOptions options)
		{
			try
			{
				StringBuilder sb = new StringBuilder();

				var cx = sx;

				// grab some extended information for the exception
				do
				{
					if (cx == sx)
					{
						sb.AppendLine("Exception:");
					}

					// gather up all the properties of the Exception, plus the extended info above
					// sort it, and render to a stringbuilder
					foreach (var item in cx.Properties
						.OrderByDescending(x => string.Equals(x.Key, "Type", StringComparison.Ordinal))
						.ThenByDescending(x => string.Equals(x.Key, "Message", StringComparison.Ordinal))
						.ThenByDescending(x => string.Equals(x.Key, "Source", StringComparison.Ordinal))
						.ThenBy(x => string.Equals(x.Key, nameof(cx.InnerException), StringComparison.Ordinal))
						.ThenBy(x => string.Equals(x.Key, nameof(AggregateException.InnerExceptions), StringComparison.Ordinal))
						.ThenBy(x => x.Key))
					{
						object value = item.Value;
						if (value != null) value = value.ToString();
						if (value == null || (value is string && string.IsNullOrEmpty((string)value)))
						{
							if (options.OmitNullProperties)
							{
								continue;
							}
							else
							{
								value = string.Empty;
							}
						}

						sb.AppendValue(item.Key, value, options);
					}
					cx = sx.InnerException;
				} while (cx != null);

				return sb.ToString().TrimEnd('\r', '\n');
			}
			catch (Exception ex2)
			{
				return string.Format("Error '{0}' while generating exception description", ex2.Message);
			}
		}


		private static string IndentString(string value, ExceptionOptions options)
		{
			return value.Replace(Environment.NewLine, Environment.NewLine + options.Indent);
		}
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
namespace ExceptionToString.Internal
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


		public static void AppendCollection(this StringBuilder sb, string propertyName, IEnumerable collection, ExceptionOptions options)
		{
			sb.AppendLine(string.Format("{0, -23}", string.Format("{0}{1}:", options.Indent, propertyName)));

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


		public static void AppendValue(this StringBuilder sb, string propertyName, object value, ExceptionOptions options)
		{
			if (value is Exception)
			{
				var innerException = (Exception)value;
				sb.AppendException(propertyName, innerException, options);
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
				sb.Append(string.Format("{0, -23}", string.Format("{0}{1}:", options.Indent, propertyName)));
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


		public static void AppendException(this StringBuilder sb, string propertyName, Exception ex, ExceptionOptions options)
		{
			var innerExceptionString = ex.ToString(new ExceptionOptions(options, options.CurrentIndentLevel + 1));

			sb.AppendLine(string.Format("{0, -23}", string.Format("{0}{1}: ", options.Indent, propertyName)));
			sb.AppendLine(innerExceptionString);
		}
	}
}
