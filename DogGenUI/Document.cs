using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DogGenUI
	{/// <summary>
	 ///	Mapped to the following columns in the [Document Collection Library]  of SharePoint:
	 ///	- values less then 10 is mappaed to [Generate Service Framework Documents]
	 ///	- values between 20 and 49 is mapped to [Generate Internal Documents]
	 ///  - values greater than 50 is mapped to [Generate External Documents] 
	 ///  - values 
	 /// </summary>
	public enum enumDocumentTypes
		{
		Service_Framework_Document_DRM_sections=1,
		Service_Framework_Document_DRM_inline=2,
		ISD_Document_DRM_Sections=20,
		ISD_Document_DRM_Inline=21,
		RACI_Workbook_per_Role=25,
		RACI_Matrix_Workbook_per_Deliverable=26,
		Content_Status_Workbook=30,
		Activity_Effort_Workbook=35,
		Internal_Technology_Coverage_Dashboard=40,
		CSD_Document_DRM_Sections=50,
		CSD_Document_DRM_Inline=51,
		CSD_based_on_Client_Requirements_Mapping=52,
		Client_Requirement_Mapping_Workbook=60,
		Contract_SoW_Service_Description=70,
		Pricing_Addendum_Document=71,
		External_Technology_Coverage_Dashboard=80
		}
	public enum enumDocumentStatusses
		{
		New=0,
		Building=1,
		Failed=3,
		Completed=5,
		Uploading=7,
		Uploaded=9,
		}

	public enum enumCSD_Document_DRM_inline_Options
		{
		Section_A_Introductories=144,
		Introduction=145,
		Executive_Summary=146,
		Section_B_Service_Portfolio_Heading=147,
		Service_Portfolio_Description=148,
		Service_Family_Heading=149,
		Service_Family_Description=150,
		Service_Product_Heading=151,
		Service_Product_Description=152,
		Service_Feature_Heading=153,
		Service_Feature_Description=154,
		Deliverables_Reports_Meetings=155,
		Deliverable_Report_Meeting_Heading=156,
		Deliverable_Report_Meeting_Description=157,
		Dimension_Datas_Obligations=158,
		Clients_Responsibilities=159,
		Exclusions=160,
		Governance_Controls=161,
		Service_Levels=162,
		Service_Level_Heading=163,
		Service_Level_Commitments_Table=164,
		Section_C_Actonyms_and_Glossary_of_Terms=165,
		Acronyms=166,
		Glossary_of_Terms=167
		}

	class Document
		{
		private int _id = 0;
		public int ID
			{
			get
				{
				return _id;
				}
			set
				{
				_id = value;
				}
			}
		private enumDocumentTypes _documentType;
		public enumDocumentTypes DocumentType
			{
			set
				{
				_documentType = value;
				}
			get
				{
				return _documentType;
				}
			}
		private int _documentCollectionID = 0;
		public int DocumentCollectionID
			{
			get
				{
				return _documentCollectionID;
				}
			set
				{
				_documentCollectionID = value;
				}
			}
		private enumDocumentStatusses _documentStatus = enumDocumentStatusses.New;
		public enumDocumentStatusses DocumentStatus
			{
			get
				{
				return _documentStatus;
				}
			set
				{
				_documentStatus = value;
				}
			}
		public bool Publish()
			{
			return false;
			}
		}
	}