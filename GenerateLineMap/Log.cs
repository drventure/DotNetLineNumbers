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
	public static class Log 
	{
		public static ILog Logger;

		public static void LogMessage(string message, params object[] messageargs)
		{
			Logger.LogMessage(message, messageargs);
		}
		public static void LogWarning(string message, params object[] messageargs)
		{
			Logger.LogWarning(message, messageargs);
		}

		public static void LogError(Exception ex, string message, params object[] messageargs)
		{
			Logger.LogError(ex, message, messageargs);
		}
	}


	public interface ILog
	{
		void LogMessage(string message, params object[] messageargs);
		void LogWarning(string message, params object[] messageargs);
		void LogError(Exception ex, string message, params object[] messageargs);
	}


	public class MSBuildLogger : ILog
	{
		private TaskLoggingHelper _log;

		public MSBuildLogger(TaskLoggingHelper log)
		{
			_log = log;
		}


		public void LogError(Exception ex, string message, params object[] messageargs)
		{
			if (!string.IsNullOrEmpty(message)) _log.LogError(message, messageargs);
			_log.LogErrorFromException(ex, true);
		}

		public void LogMessage(string message, params object[] messageargs)
		{
			_log.LogMessage(message, messageargs);
		}

		public void LogWarning(string message, params object[] messageargs)
		{
			_log.LogWarning(message, messageargs);
		}
	}


	public class ConsoleLogger : ILog
	{
		private string Combine(string message, params object[] messageargs)
		{
			if (messageargs == null) return message;
			return string.Format(message, messageargs);
		}

		public void LogError(Exception ex, string message, params object[] messageargs)
		{
			if (!string.IsNullOrEmpty(message)) Console.WriteLine("ERROR: " + Combine(message, messageargs));
			Console.WriteLine(ex.ToString());
		}

		public void LogMessage(string message, params object[] messageargs)
		{
			Console.WriteLine(Combine(message, messageargs));
		}

		public void LogWarning(string message, params object[] messageargs)
		{
			Console.WriteLine("WARNING: " + Combine(message, messageargs));
		}
	}


}
