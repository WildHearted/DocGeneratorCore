using System;

namespace DocGeneratorCore
	{

	//===g
	//+ LocalDatabase Exception
	public class LocalDatabaseExeption : Exception
		{
		public LocalDatabaseExeption()
			{
			}
		public LocalDatabaseExeption(string message) : base(message)
			{
			}
		public LocalDatabaseExeption(string message, Exception innerException) : base(message, innerException)
			{
			}
		}


	//===g
	//+ InvalidContentFormatException
	public class InvalidContentFormatException:Exception
		{
		public InvalidContentFormatException()
			{ }
		public InvalidContentFormatException(string message) : base(message)
			{ }
		public InvalidContentFormatException(string message, Exception innerException) : base(message, innerException)
			{ }
		}

	//===g
	//+ InvalidTableFormatException
	public class InvalidTableFormatException : Exception
		// The invalid Table Format Exception will translate into the InvalidContentFormatException in the HTMLdecoder.DecodeHTML method
		{
		public InvalidTableFormatException()
			{ }
		public InvalidTableFormatException(string message) : base(message) 
			{ }
		public InvalidTableFormatException(string message, Exception innerException) : base(message, innerException)
			{ }
		}

	//===g
	//+ InvalidImageFormatException
	public class InvalidImageFormatException:Exception
		// The invalid Table Format Exception will translate into the InvalidContentFormatException in the HTMLdecoder.DecodeHTML method
		{
		public InvalidImageFormatException()
			{ }
		public InvalidImageFormatException(string message) : base(message)
			{ }
		public InvalidImageFormatException(string message, Exception innerException) : base(message, innerException)
			{ }
		}

	//===g
	//+ InvalidRichTextException
	public class InvalidRichTextFormatException : Exception
	{
	public InvalidRichTextFormatException()
		{ }
	public InvalidRichTextFormatException(string message) : base(message)
		{ }
	public InvalidRichTextFormatException(string message, Exception innerException) : base(message, innerException)
		{ }
	}

	//===g
	//+ GeneralException
	public class GeneralException : Exception
		{
		public GeneralException()
			{ }
		public GeneralException(string message) : base(message)
			{ }
		public GeneralException(string message, Exception innerException) : base(message, innerException)
			{ }
		}

	//===g
	//+ UnableToCreateDocumentException
	public class UnableToCreateDocumentException : Exception
		{
		public UnableToCreateDocumentException()
			{ }
		public UnableToCreateDocumentException(string message) : base(message)
			{ }
		public UnableToCreateDocumentException(string message, Exception innerException) : base(message, innerException)
			{ }
		}
	
	//===g
	//+ NoContentSpecifiedException
	public class NoContentSpecifiedException : Exception
		{
		public NoContentSpecifiedException()
			{ }
		public NoContentSpecifiedException(string message) : base(message)
			{ }
		public NoContentSpecifiedException(string message, Exception innerException) : base(message, innerException)
			{ }
		}
	
	//===g
	//+ DocumentUpload Exception
	public class DocumentUploadException : Exception
		{
		public DocumentUploadException()
			{ }
		public DocumentUploadException(string message) : base(message)
			{ }
		public DocumentUploadException(string message, Exception innerException) : base(message, innerException)
			{ }
		}

	}
