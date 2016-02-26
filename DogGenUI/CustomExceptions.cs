using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGenerator
	{
	public class InvalidTableFormatException: Exception
		{
		public InvalidTableFormatException()
			{

			}
		public InvalidTableFormatException(string message)
			: base(message)
			{

			}
		public InvalidTableFormatException(string message, Exception innerException)
			: base(message, innerException)
			{

			}
		}
	}
