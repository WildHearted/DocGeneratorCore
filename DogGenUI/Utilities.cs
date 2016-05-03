using System;
using System.Collections.Generic;
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
		
		}
	}
