using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGeneratorCore
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
	public class InvalidRichTextFormatException:Exception
		{
		public InvalidRichTextFormatException()
			{

			}
		public InvalidRichTextFormatException(string message)
			: base(message)
			{

			}
		public InvalidRichTextFormatException(string message, Exception innerException)
			: base(message, innerException)
			{

			}
		}

	public class GeneralException:Exception
		{
		public GeneralException()
			{

			}
		public GeneralException(string message)
			: base(message)
			{

			}
		public GeneralException(string message, Exception innerException)
			: base(message, innerException)
			{

			}
		}

	}
