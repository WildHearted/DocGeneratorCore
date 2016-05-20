using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocGeneratorCore
	{
	//+ InvalidTableFormatException
	public class InvalidTableFormatException : Exception
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
	//+ InvalidRichTextException
	public class InvalidRichTextFormatException : Exception
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

	//+ GeneralException
	public class GeneralException : Exception
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

	//+ UnableToCreateDocumentException
	public class UnableToCreateDocumentException : Exception
		{
		public UnableToCreateDocumentException()
			{

			}
		public UnableToCreateDocumentException(string message)
			: base(message)
			{
	
			}
		public UnableToCreateDocumentException(string message, Exception innerException)
			: base(message, innerException)
			{

			}
		}

	//+ NoContentSpecifiedException
	public class NoContentSpecifiedException : Exception
		{
		public NoContentSpecifiedException()
			{

			}
		public NoContentSpecifiedException(string message)
			: base(message)
			{

			}
		public NoContentSpecifiedException(string message, Exception innerException)
			: base(message, innerException)
			{

			}
		}

	//+ DocumentUpload Exception
	public class DocumentUploadException : Exception
		{
		public DocumentUploadException()
			{

			}
		public DocumentUploadException(string message)
			: base(message)
			{

			}
		public DocumentUploadException(string message, Exception innerException)
			: base(message, innerException)
			{

			}
		}

	}
