using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGeneratorCore
	{
	class Utilities
		{
		public static void WriteErrorsToConsole(List<String> parErrors)
			{
			if(parErrors == null)
				{
				return;
				}
			foreach(string error in parErrors)
				{
				Console.WriteLine("\t\t\t * {0}", error);
				}

			}

		/// <summary>
		/// Check fi a specific process is running and return True or False depending on the result,
		/// </summary>
		/// <param name="parProcessName">Pass the EXACT Name of the process that need to be checked as a string.</param>
		/// <returns>True if the process is running, False if it is NOT running.</returns>
		public static bool IsProcessRunning(string parProcessName)
			{
			return Process.GetProcessesByName(processName: parProcessName).Length > 0 ? true : false;
			}
		}
	}
	
