using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Build.Utilities;

namespace GenerateLineMap
{
	/// <summary>
	/// Public Log class to implement ILog interface for integration with MSBuild logging.
	/// </summary>
	/// <exclude />
	public static class Log 
	{
		/// <summary>
		/// Expose a logger to integrate with MSBuild logging
		/// </summary>
		public static ILog Logger;

		/// <summary>
		/// Log a standard build message
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public static void LogMessage(string message, params object[] messageargs)
		{
			Logger.LogMessage(message, messageargs);
		}

		/// <summary>
		/// Log a standard warning message
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public static void LogWarning(string message, params object[] messageargs)
		{
			Logger.LogWarning(message, messageargs);
		}

		/// <summary>
		/// log a standard error message
		/// </summary>
		/// <param name="ex"></param>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public static void LogError(Exception ex, string message, params object[] messageargs)
		{
			Logger.LogError(ex, message, messageargs);
		}
	}


	/// <summary>
	/// Interface that must be implemented to support integration with MSBuild logging.
	/// </summary>
	/// <exclude />
	public interface ILog
	{
		/// <summary>
		/// Log a standard message
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		void LogMessage(string message, params object[] messageargs);

		/// <summary>
		/// Log a standard warning message
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		void LogWarning(string message, params object[] messageargs);

		/// <summary>
		/// Log a standard error
		/// </summary>
		/// <param name="ex"></param>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		void LogError(Exception ex, string message, params object[] messageargs);
	}


	/// <summary>
	/// Public class required to support integration with MSBuild logging.
	/// </summary>
	/// <exclude />
	public class MSBuildLogger : ILog
	{
		private TaskLoggingHelper _log;

		/// <summary>
		/// Default Constructor
		/// </summary>
		/// <param name="log"></param>
		public MSBuildLogger(TaskLoggingHelper log)
		{
			_log = log;
		}


		/// <summary>
		/// Log an error 
		/// </summary>
		/// <param name="ex"></param>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public void LogError(Exception ex, string message, params object[] messageargs)
		{
			if (!string.IsNullOrEmpty(message)) _log.LogError(message, messageargs);
			_log.LogErrorFromException(ex, true);
		}

		/// <summary>
		/// Log a Message
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public void LogMessage(string message, params object[] messageargs)
		{
			_log.LogMessage(message, messageargs);
		}

		/// <summary>
		/// Log a warning
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public void LogWarning(string message, params object[] messageargs)
		{
			_log.LogWarning(message, messageargs);
		}
	}


	/// <summary>
	/// Alternate logging class used when GenerateLineMap is run directly from the command line.
	/// </summary>
	/// <exclude />
	public class ConsoleLogger : ILog
	{
		private string Combine(string message, params object[] messageargs)
		{
			if (messageargs == null) return message;
			return string.Format(message, messageargs);
		}

		/// <summary>
		/// Log an error
		/// </summary>
		/// <param name="ex"></param>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public void LogError(Exception ex, string message, params object[] messageargs)
		{
			if (!string.IsNullOrEmpty(message)) Console.WriteLine("ERROR: " + Combine(message, messageargs));
			Console.WriteLine(ex.ToString());
		}

		/// <summary>
		/// Log a message
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public void LogMessage(string message, params object[] messageargs)
		{
			Console.WriteLine(Combine(message, messageargs));
		}

		/// <summary>
		/// Log a warning
		/// </summary>
		/// <param name="message"></param>
		/// <param name="messageargs"></param>
		public void LogWarning(string message, params object[] messageargs)
		{
			Console.WriteLine("WARNING: " + Combine(message, messageargs));
		}
	}
}
