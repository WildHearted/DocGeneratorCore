using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Runtime.Caching;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{/// <summary>
	 ///	Mapped to the following columns in the [Document Collection Library]  of SharePoint:
	 ///	- values less then 10 is mappaed to [Generate Service Framework Documents]
	 ///	- values between 20 and 49 is mapped to [Generate Internal Documents]
	 /// - values greater than 50 is mapped to [Generate External Documents] 
	 /// - values 
	 /// </summary>
	public enum enumDocumentTypes
		{
		Service_Framework_Document_DRM_sections=1,	// class defined
		Service_Framework_Document_DRM_inline=2,	// class defined
		ISD_Document_DRM_Sections=20,				// class defined
		ISD_Document_DRM_Inline=21,				// class defined
		RACI_Workbook_per_Role=25,				// class defined
		RACI_Matrix_Workbook_per_Deliverable=26,	// class defined
		Content_Status_Workbook=30,				// class defined
		Activity_Effort_Workbook=35,				// no Class - removed from scope
		Internal_Technology_Coverage_Dashboard=40,	// class defined
		CSD_Document_DRM_Sections=50,				// class defined
		CSD_Document_DRM_Inline=51,				// class defined
		CSD_based_on_Client_Requirements_Mapping=52,	// class defined
		Client_Requirement_Mapping_Workbook=60,		// class defined
		Contract_SoW_Service_Description=70,		// class defined
		Pricing_Addendum_Document=71,				// class defined
		External_Technology_Coverage_Dashboard=80	// class defined
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

	class Document_Workbook
		{
		// Object Fields
		public string text2Write = "";

		// Object Properties
		private int _id = 0;
		public int ID
			{
			get{return this._id;}
			set{this._id = value;}
			}
		private enumDocumentTypes _documentType;
		public enumDocumentTypes DocumentType
			{
			set{this._documentType = value;}
			get{return this._documentType;}
			}
		private int _documentCollectionID = 0;
		public int DocumentCollectionID
			{
			get{return this._documentCollectionID;}
			set{this._documentCollectionID = value;}
			}
		private string _IntroductionRichText;
		public string IntroductionRichText
			{
			get{return this._IntroductionRichText;}
			set{this._IntroductionRichText = value;}
			}
		private string _ExecutiveSummaryRichText;
		public string ExecutiveSummaryRichText
			{
			get{return this._ExecutiveSummaryRichText;}
			set{this._ExecutiveSummaryRichText = value;}
			}
		private string _DocumentAcceptanceRichText;
		public String DocumentAcceptanceRichText
			{
			get{return this._DocumentAcceptanceRichText;}
			set{this._DocumentAcceptanceRichText = value;}
			}
		private enumDocumentStatusses _documentStatus = enumDocumentStatusses.New;
		public enumDocumentStatusses DocumentStatus
			{
			get{return this._documentStatus;}
			set{this._documentStatus = value;}
			}
		private bool _hyperlink_View = false;
		public bool Hyperlink_View
			{
			get{return this._hyperlink_View;}
			set{this._hyperlink_View = value;}
			}
		private bool _hyperlinkEdit = false;
		public bool HyperlinkEdit
			{
			get{return this._hyperlinkEdit;}
			set{this._hyperlinkEdit = value;}
			}
		private string _template = "";
		public string Template
			{
			get{return this._template;}
			set{this._template = value;}
			}
		private List<Hierarchy> _selectedNodes;
		/// <summary>
		/// This property is a List of Hierarchy objects which represent the nodes (content) that need to be included in the generated document.
		/// </summary>
		public List<Hierarchy> SelectedNodes
			{
			get{return this._selectedNodes;}
			set{this._selectedNodes = value;}
			}
		private List<string> _errorMessages = new List<string>();
		/// <summary>
		/// This property is a list of strings that will contain all the error messages why this specific 
		/// Document instance cannot be generated.
		/// </summary>
		public List<string> ErrorMessages
			{
			get{return _errorMessages;}
			private set{_errorMessages = value;}
			}
		// Methods:
		/// <summary>
		/// Use this method whenever an error occurs while preparing a Document object before it is generated,
		/// to add each fo the errors to the list of errors. 
		/// </summary>
		/// <param name="parErrorString"></param>
		public void LogError(string parErrorString)
			{
			//List<string> listNewErrors = new List<string>();
			//listNewErrors.Add(parErrorString);
			this.ErrorMessages.Add(parErrorString);
			}

		/// <summary>
		/// This method is used to publish the document to the document collection once it has been created.
		/// </summary>
		/// <returns>Returns True if successfully published else returns False.</returns>
		public bool Publish()
			{
			//TODO: Develop the Document Publish method
			return false;
			}
		}

	/// <summary>
	/// This is the base class for all documents. 
	/// The LOWEST level sub-class must alwasy be used to configure/setup generatable documents.
	/// </summary>
	class Document : Document_Workbook
		{
		private bool _introductories_Section = false;
		public bool Introductory_Section
			{
			get{return this._introductories_Section;}
			set{this._introductories_Section = value;}
			}
		private bool _introduction = false;
		public bool Introduction
			{
			get{return this._introduction;}
			set{this._introduction = value;}
			}
		private bool _executive_Summary = false;
		public bool Executive_Summary
			{
			get{return this._executive_Summary;}
			set{this._executive_Summary = value;}
			}
		private bool _Acronyms_Glossary_of_Terms_Section = false;
		public bool Acronyms_Glossary_of_Terms_Section
			{
			get{return this._Acronyms_Glossary_of_Terms_Section;}
			set{this._Acronyms_Glossary_of_Terms_Section = value;}
			}
		private bool _acronyms = false;
		public bool Acronyms
			{
			get{return this._acronyms;}
			set{this._acronyms = value;}
			}
		private List<TermAndAcronym> _termsAndAcronymList = new List<TermAndAcronym>();
		public List<TermAndAcronym> TermAndAcronymList
			{
			get{return this._termsAndAcronymList;}
			set{this._termsAndAcronymList = value;}
			}
		/// <summary>
		/// 
		/// </summary>
		private bool _glossary_of_Terms = false;
		public bool Glossary_of_Terms
			{
			get{return this._glossary_of_Terms;}
			set{this._glossary_of_Terms = value;}
			}
		/// <summary>
		/// 
		/// </summary>
		private UInt32 _pageHeight = 0;
		public UInt32 PageHight
			{
			get{return this._pageHeight;}
			set{this._pageHeight = value;}
			}
		/// <summary>
		/// 
		/// </summary>
		private UInt32 _pageWidth = 0;
		public UInt32 PageWith
			{
			get{return this._pageWidth;}
			set{this._pageWidth = value;}
			}

		private bool _colorCodingLayer1 = false;
		public bool ColorCodingLayer1
			{
			get{return this._colorCodingLayer1;}
			set{this._colorCodingLayer1 = value;}
			}

		private bool _colorCodingLayer2 = false;
		public bool ColorCodingLayer2
			{
			get{return this._colorCodingLayer2;}
			set{this._colorCodingLayer2 = value;}
			}

		private bool _colorCodingLayer3 = false;
		public bool ColorCodingLayer3
			{
			get{return this._colorCodingLayer3;}
			set{this._colorCodingLayer3 = value;}
			}

		}
	
	/// <summary>
	/// This class is the sub class on which all Workbooks are based.
	/// </summary>
	class Workbook : Document_Workbook
		{
		// deliberately kept open for future population if and when needed
		}

	/// <summary>
	/// This class handles the RACI Workbook per Role
	/// </summary>
	class RACI_Workbook_per_Role : Workbook
		{
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Add code for RACI_Workbook_per_Role 
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class handles the Client_Requirements_Mapping_Workbook
	/// </summary>
	class Client_Requirements_Mapping_Workbook : Workbook
		{
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to be added for Client_Requirements_Mapping_Workbook's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class handles the RACI Matrix Workbook per Deliverable
	/// </summary>
	class RACI_Matrix_Workbook_per_Deliverable : Workbook
		{
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to be added for RACI_Matrix_Workbook_per_Deliverable's Generate method...
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}
	
	/// <summary>
	/// This class handles the Content Status Workbook
	/// </summary>
	class Content_Status_Workbook : Workbook
		{
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for Content_Status_Workbook's Generate method
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class handles the Internal Technology coverage Dashbord Workbook
	/// </summary>
	class Internal_Technology_Coverage_Dashboard_Workbook : Workbook
		{
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for Internal_Technology_Coverage_Dashboard_Workbook's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class handles the External Technology coverage Dashbord Workbook
	/// </summary>
	class External_Technology_Coverage_Dashboard_Workbook : Workbook
		{
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for External_Technology_Coverage_Dashboard_Workbook's Generate Method
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class is used to set all the properties for a
	/// CLient Service Description (CSD) based on a Client Requirements Mapping (CRM) Document.
	/// It inherits from the Document class.
	/// </summary>
	class CSD_based_on_ClientRequirementsMapping : Document
		{
		private bool _csd_Doc_based_on_CRM = false;
		public bool CSD_Doc_based_on_CRM
			{
			get{return this._csd_Doc_based_on_CRM;}
			set{this._csd_Doc_based_on_CRM = value;}
			}
		private int _crm_Mapping = 0;
		/// <summary>
		/// This property reference the ID value of the SharePoint Mappings entry which is used to generate the Document
		/// </summary>
		public int CRM_Mapping
			{
			get{return this._crm_Mapping;}
			set{this._crm_Mapping = value;}
			}
		private bool _requirements_Section = false;
		public bool Requirements_Section
			{
			get{return this._requirements_Section;}
			set{this._requirements_Section = value;}
			}
		private bool _tower_of_Service_Heading = false;
		public bool Tower_of_Service_Heading
			{
			get{return _tower_of_Service_Heading;}
			set{this._tower_of_Service_Heading = value;}
			}
		private bool _requirement_Heading = false;
		public bool Requirement_Heading
			{
			get{return this._requirement_Heading;}
			set{this._requirement_Heading = value;}
			}
		private bool _requirement_Reference = false;
		public bool Requirement_Reference
			{
			get{return this._requirement_Reference = false;}
			set{this._requirement_Reference = value;}
			}
		private bool _requirement_Text = false;
		public bool Requirement_Text
			{
			get{return this._requirement_Text;}
			set{this._requirement_Text = value;}
			}
		private bool _requirement_Service_Level = false;
		public bool Requirement_Service_Level
			{
			get{return this._requirement_Service_Level;}
			set{this._requirement_Service_Level = value;}
			}
		private bool _risks = false;
		public bool Risks
			{
			get{return this._risks;}
			set{this._risks = value;}
			}
		private bool _risk_Heading = false;
		public bool Risk_Heading
			{
			get{return this._risk_Heading;}
			set{this._risk_Heading = value;}
			}
		private bool _risk_Description = false;
		public bool Risk_Description
			{
			get{return this._risk_Description;}
			set{this._risk_Description = value;}
			}
		private bool _assumptions = false;
		public bool Assumptions
			{
			get{return this._assumptions;}
			set{this._assumptions = value;}
			}
		private bool _assumption_Heading;
		public bool Assumption_Heading
			{
			get{return this._assumption_Heading;}
			set{this._assumption_Heading = value;}
			}
		private bool _assumption_Description = false;
		public bool Assumption_Description
			{
			get{return this._assumption_Description;}
			set{this._assumption_Description = value;}
			}
		private bool _deliverables_Reports_and_Meetings = false;
		public bool Deliverable_Reports_and_Meetings
			{
			get{return this._deliverables_Reports_and_Meetings;}
			set{this._deliverables_Reports_and_Meetings = value;}
			}
		private bool _drm_Heading = false;
		public bool DRM_Heading
			{
			get{return this._drm_Heading;}
			set{this._drm_Heading = value;}
			}
		private bool _drm_Description = false;
		public bool DRM_Descrioption
			{
			get{return this._drm_Description;}
			set{this._drm_Description = value;}
			}
		private bool _dds_DRM_Obligations = false;
		public bool DDs_DRM_Obligations
			{
			get{return this._dds_DRM_Obligations;}
			set{this._dds_DRM_Obligations = value;}
			}
		private bool _clients_DRM_Responsibilities = false;
		public bool Clients_DRM_Responsibiities
			{
			get{return this._clients_DRM_Responsibilities;}
			set{this._clients_DRM_Responsibilities = value;}
			}
		private bool _drm_Exclusions = false;
		public bool DRM_Exclusions
			{
			get{return this._drm_Exclusions;}
			set{this._drm_Exclusions = value;}
			}
		private bool _drm_Governance_Controls = false;
		public bool DRM_Governance_Controls
			{
			get{return this._drm_Governance_Controls;}
			set{this._drm_Governance_Controls = value;}
			}
		private bool _service_Levels = false;
		public bool Service_Levels
			{
			get{return this._service_Levels;}
			set{this._service_Levels = value;}
			}
		private bool _service_Level_Heading = false;
		public bool Service_Level_Heading
			{
			get{return this._service_Level_Heading;}
			set{this._service_Level_Heading = value;}
			}
		private bool _service_Level_Commitments_Table = false;
		public bool Service_Level_Commitments_Table
			{
			get{return this._service_Level_Commitments_Table;}
			set{this._service_Level_Commitments_Table = value;}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for CSD_based_on_ClientRequirementsMapping's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}

		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 168:
								this.Introductory_Section = true;
								break;
							case 169:
								this.Introduction = true;
								break;
							case 170:
								this.Executive_Summary = true;
								break;
							case 171:
								this.Requirements_Section = true;
								break;
							case 172:
								this.Tower_of_Service_Heading = true;
								break;
							case 173:
								this.Requirement_Heading = true;
								break;
							case 174:
								this.Requirement_Reference = true;
								break;
							case 175:
								this.Requirement_Text = true;
								break;
							case 176:
								this.Requirement_Service_Level = true;
								break;
							case 177:
								this.Risks = true;
								break;
							case 178:
								this.Risk_Heading = true;
								break;
							case 179:
								this.Risk_Description = true;
								break;
							case 180:
								this.Assumptions = true;
								break;
							case 181:
								this.Assumption_Heading = true;
								break;
							case 182:
								this.Deliverable_Reports_and_Meetings = true;
								break;
							case 183:
								this.DRM_Heading = true;
								break;
							case 184:
								this.DRM_Descrioption = true;
								break;
							case 185:
								this.DDs_DRM_Obligations = true;
								break;
							case 186:
								this.Clients_DRM_Responsibiities = true;
								break;
							case 187:
								this.DRM_Exclusions = true;
								break;
							case 188:
								this.DRM_Governance_Controls = true;
								break;
							case 189:
								this.Service_Levels = true;
								break;
							case 190:
								this.Service_Level_Heading = true;
								break;
							case 191:
								this.Service_Level_Commitments_Table = true;
								break;
							case 192:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 193:
								this.Acronyms = true;
								break;
							case 194:
								this.Glossary_of_Terms = true;
								break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		} // end of CSD_ClientRequirementsMapping_Document class

	/// <summary>
	/// This class inherits from the Document class and contain all the common properties and methods that
	/// the Predefined product documents have.
	/// </summary>
	class PredefinedProduct_Document : Document
		{
		private bool _service_Portfolio_Section = false;
		public bool Service_Portfolio_Section
			{
			get{return this._service_Portfolio_Section;}
			set{this._service_Portfolio_Section = value;}
			}
		private bool _service_Portfolio_Description = false;
		public bool Service_Portfolio_Description
			{
			get{return this._service_Portfolio_Description;}
			set{this._service_Portfolio_Description = value;}
			}
		private bool _service_Family_Heading = false;
		public bool Service_Family_Heading
			{
			get{return this._service_Family_Heading;}
			set{this._service_Family_Heading = value;}
			}
		private bool _service_Family_Description = false;
		public bool Service_Family_Description
			{
			get
				{
				return this._service_Family_Description;
				}
			set
				{
				this._service_Family_Description = value;
				}
			}
		private bool _service_Product_Heading = false;
		public bool Service_Product_Heading
			{
			get
				{
				return this._service_Product_Heading;
				}
			set
				{
				this._service_Product_Heading = value;
				}
			}
		private bool _service_Product_Description = false;
		public bool Service_Product_Description
			{
			get
				{
				return this._service_Product_Description;
				}
			set
				{
				this._service_Product_Description = value;
				}
			}
		private bool _drm_Heading = false;
		public bool DRM_Heading
			{
			get
				{
				return this._drm_Heading;
				}
			set
				{
				this._drm_Heading = value;
				}
			}
		private bool _Deliverables_Reports_Meetings = false;
		public bool Deliverables_Reports_Meetings
			{
			get
				{
				return this._Deliverables_Reports_Meetings;
				}
			set
				{
				this._Deliverables_Reports_Meetings = value;
				}
			}
		private bool _service_Levels = false;
		public bool Service_Levels
			{
			get
				{
				return this._service_Levels;
				}
			set
				{
				this._service_Levels = value;
				}
			}
		private bool _service_Level_Heading = false;
		public bool Service_Level_Heading
			{
			get
				{
				return this._service_Level_Heading;
				}
			set
				{
				this._service_Level_Heading = value;
				}
			}
		private bool _service_Level_Commitments_Table = false;
		public bool Service_Level_Commitments_Table
			{
			get
				{
				return this._service_Level_Commitments_Table;
				}
			set
				{
				this._service_Level_Commitments_Table = value;
				}
			}
		} // end of PredefinedProduct_Document class
	
	/// <summary>
	/// This class inherits from the PredefinedProduct_Document class and contain all the common properties and methods that
	/// the External (Client Facing) documents have.
	/// </summary>
	class External_Document : PredefinedProduct_Document
		{
		private bool _service_Feature_Heading = false;
		public bool Service_Feature_Heading
			{
			get
				{
				return this._service_Feature_Heading;
				}
			set
				{
				this._service_Feature_Heading = value;
				}
			}
		private bool _service_Feature_Description = false;
		public bool Service_Feature_Description
			{
			get
				{
				return this._service_Feature_Description;
				}
			set
				{
				this._service_Feature_Description = value;
				}
			}
		} // End of the External_Document class

	/// <summary>
	/// This class inherits from the PredefinedProduct_Document class and contain all the common properties and methods that the Internal documents have.
	/// </summary>
	class Internal_Document : PredefinedProduct_Document
		{
		private bool _service_Product_Key_Client_Benefits = false;
		public bool Service_Product_Key_Client_Benefits
			{
			get
				{
				return this._service_Product_Key_Client_Benefits;
				}
			set
				{
				this._service_Product_Key_Client_Benefits = value;
				}
			}
		private bool _service_Product_Key_DD_Benefits = false;
		public bool Service_Product_KeyDD_Benefits
			{
			get
				{
				return this._service_Product_Key_DD_Benefits;
				}
			set
				{
				this._service_Product_Key_DD_Benefits = value;
				}
			}
		private bool _service_Element_Heading = false;
		public bool Service_Element_Heading
			{
			get
				{
				return this._service_Element_Heading;
				}
			set
				{
				this._service_Element_Heading = value;
				}
			}
		private bool _service_Element_Description = false;
		public bool Service_Element_Description
			{
			get
				{
				return this._service_Element_Description;
				}
			set
				{
				this._service_Element_Description = value;
				}
			}
		private bool _service_Element_Objectives = false;
		public bool Service_Element_Objectives
			{
			get
				{
				return this._service_Element_Objectives;
				}
			set
				{
				this._service_Element_Objectives = value;
				}
			}
		private bool _service_Element_Key_Client_Benefits = false;
		public bool Service_Element_Key_Client_Benefits
			{
			get
				{
				return this._service_Element_Key_Client_Benefits;
				}
			set
				{
				this._service_Element_Key_Client_Benefits = value;
				}
			}
		private bool _service_Element_Key_Client_Advantages = false;
		public bool Service_Element_Key_Client_Advantages
			{
			get
				{
				return this._service_Element_Key_Client_Advantages;
				}
			set
				{
				this._service_Element_Key_Client_Advantages = value;
				}
			}
		private bool _service_Element_Key_DD_Benefits = false;
		public bool Service_Element_Key_DD_Benefits
			{
			get
				{
				return this._service_Element_Key_DD_Benefits;
				}
			set
				{
				this._service_Element_Key_DD_Benefits = value;
				}
			}
		private bool _service_Element_Critical_Success_Factors = false;
		public bool Service_Element_Critical_Success_Factors
			{
			get
				{
				return this._service_Element_Critical_Success_Factors;
				}
			set
				{
				this._service_Element_Critical_Success_Factors = value;
				}
			}
		private bool _service_Element_Key_Performance_Indicators = false;
		public bool Service_Element_Key_Performance_Indicators
			{
			get
				{
				return this._service_Element_Key_Performance_Indicators;
				}
			set
				{
				this._service_Element_Key_Performance_Indicators = value;
				}
			}
		private bool _service_Element_High_Level_Process = false;
		public bool Service_Element_High_Level_Process
			{
			get
				{
				return this._service_Element_High_Level_Process;
				}
			set
				{
				this._service_Element_High_Level_Process = value;
				}
			}
		private bool _activities = false;
		public bool Activities
			{
			get
				{
				return this._activities;
				}
			set
				{
				this._activities = value;
				}
			}
		private bool _activity_Heading = false;
		public bool Activity_Heading
			{
			get
				{
				return this._activity_Heading;
				}
			set
				{
				this._activity_Heading = value;
				}
			}
		private bool _activity_Description_Table = false;
		public bool Activity_Description_Table
			{
			get
				{
				return this._activity_Description_Table;
				}
			set
				{
				this._activity_Description_Table = value;
				}
			}
		private bool _document_Acceptance_Section = false;
		public bool Document_Acceptance_Section
			{
			get
				{
				return this._document_Acceptance_Section;
				}
			set
				{
				this._document_Acceptance_Section = value;
				}
			}
		} // End of the Internal_Document class

	/// <summary>
	/// This class contains all the Client Service Description (CSD) with inline DRM (Deliverable Report Meeting).
	/// </summary>
	class CSD_Document_DRM_Inline : External_Document
		{
		private bool _drm_Description = false;
		public bool DRM_Description
			{
			get
				{
				return _drm_Description;
				}
			set
				{
				_drm_Description = value;
				}
			}
		private bool _drm_Inputs = false;
		public bool DRM_Inputs
			{
			get
				{
				return _drm_Inputs;
				}
			set
				{
				_drm_Inputs = value;
				}
			}
		private bool _drm_Outputs = false;
		public bool DRM_Outputs
			{
			get
				{
				return _drm_Outputs;
				}
			set
				{
				_drm_Outputs = value;
				}
			}
		private bool _dds_DRM_Obligations = false;
		public bool DDS_DRM_Obligations
			{
			get
				{
				return _dds_DRM_Obligations;
				}
			set
				{
				_dds_DRM_Obligations = value;
				}
			}
		private bool _clients_DRM_Responsibilities = false;
		public bool Clients_DRM_Responsibilities
			{
			get
				{
				return _clients_DRM_Responsibilities;
				}
			set
				{
				_clients_DRM_Responsibilities = value;
				}
			}
		private bool _drm_Exclusions = false;
		public bool DRM_Exclusions
			{
			get
				{
				return _drm_Exclusions;
				}
			set
				{
				_drm_Exclusions = value;
				}
			}
		private bool _drm_Governance_Controls = false;
		public bool DRM_Governance_Controls
			{
			get
				{
				return _drm_Governance_Controls;
				}
			set
				{
				_drm_Governance_Controls = value;
				}
			}
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added to CSD_Document_DRM_Inline's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}

		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 144:
								this.Introductory_Section = true;
								break;
							case 145:
								this.Introduction = true;
								break;
							case 146:
								this.Executive_Summary = true;
								break;
							case 147:
								this.Service_Portfolio_Section = true;
								break;
							case 148:
								this.Service_Portfolio_Description = true;
								break;
							case 149:
								this.Service_Family_Heading = true;
								break;
							case 150:
								this.Service_Family_Description = true;
								break;
							case 151:
								this.Service_Product_Heading = true;
								break;
							case 152:
								this.Service_Product_Description = true;
								break;
							case 153:
								this.Service_Feature_Heading = true;
								break;
							case 154:
								this.Service_Feature_Description = true;
								break;
							case 155:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 156:
								this.DRM_Heading = true;
								break;
							case 157:
								this.DRM_Description = true;
								break;
							case 158:
								this.DDS_DRM_Obligations = true;
								break;
							case 159:
								this.Clients_DRM_Responsibilities = true;
								break;
							case 160:
								this.DRM_Exclusions = true;
								break;
							case 161:
								this.DRM_Governance_Controls = true;
								break;
							case 162:
								this.Service_Levels = true;
								break;
							case 163:
								this.Service_Level_Heading = true;
								break;
							case 164:
								this.Service_Level_Commitments_Table = true;
								break;
							case 165:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 166:
								this.Acronyms = true;
								break;
							case 167:
								this.Glossary_of_Terms = true;
								break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		} // end of CSD_inline DRM class

	class Pricing_Addendum_Document : Document
		{
		private int _pricing_Worksbook_Id = 0;
		public int Pricing_Workbook_Id
			{
			get
				{
				return _pricing_Worksbook_Id;
				}
			set
				{
				_pricing_Worksbook_Id = value;
				}
			}
		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for Pricing_Addendum_Document's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		}

	/// <summary>
	/// This class contains all the Client Service Description (CSD) with inline DRM (Deliverable Report Meeting).
	/// </summary>
	class Internal_DRM_Inline : Internal_Document
		{
		private bool _drm_Description = false;
		public bool DRM_Description
			{
			get{return this._drm_Description;}
			set{this._drm_Description = value;}
			}
		private bool _drm_Inputs = false;
		public bool DRM_Inputs
			{
			get{return this._drm_Inputs;}
			set{this._drm_Inputs = value;}
			}
		private bool _drm_Outputs = false;
		public bool DRM_Outputs
			{
			get{return this._drm_Outputs;}
			set{this._drm_Outputs = value;}
			}
		private bool _dds_DRM_Obligations = false;
		public bool DDS_DRM_Obligations
			{
			get{return this._dds_DRM_Obligations;}
			set{this._dds_DRM_Obligations = value;}
			}
		private bool _clients_DRM_Responsibilities = false;
		public bool Clients_DRM_Responsibilities
			{
			get{return this._clients_DRM_Responsibilities;}
			set{this._clients_DRM_Responsibilities = value;}
			}
		private bool _drm_Exclusions = false;
		public bool DRM_Exclusions
			{
			get{return this._drm_Exclusions;}
			set{this._drm_Exclusions = value;}
			}
		private bool _drm_Governance_Controls = false;
		public bool DRM_Governance_Controls
			{
			get{return this._drm_Governance_Controls;}
			set{this._drm_Governance_Controls = value;}
			}

		} // end of CSD_inline DRM class

	/// <summary>
	/// This class contains all the properties and methods for Internal DRM (Deliverable Report Meeting) Sections object
	/// </summary>
	class Internal_DRM_Sections : Internal_Document
		{
		private bool _drm_Summary = false;
		public bool DRM_Summary
			{
			get{return this._drm_Summary;}
			set{this._drm_Summary = value;}
			}
		private bool _drm_Section = false;
		public bool DRM_Section
			{
			get{return this._drm_Section;}
			set{this._drm_Section = value;}
			}
		private bool _deliverables = false;
		public bool Deliverables
			{get{return this._deliverables;}
			set{this._deliverables = value;}
			}
		private bool _deliverable_Heading = false;
		public bool Deliverable_Heading
			{
			get{return this._deliverable_Heading;}
			set{this._deliverable_Heading = value;}
			}
		private bool _deliverable_Description = false;
		public bool Deliverable_Description
			{
			get{return this._deliverable_Description;}
			set{this._deliverable_Description = value;}
			}
		private bool _deliverable_Inputs = false;
		public bool Deliverable_Inputs
			{
			get{return this._deliverable_Inputs;}
			set{this._deliverable_Inputs = value;}
			}
		private bool _deliverable_Outputs = false;
		public bool Deliverable_Outputs
			{
			get{return this._deliverable_Outputs;}
			set{this._deliverable_Outputs = value;}
			}
		private bool _dds_Deliverable_Obligations = false;
		public bool DDs_Deliverable_Obligations
			{
			get{return this._dds_Deliverable_Obligations;}
			set{this._dds_Deliverable_Obligations = value;}
			}
		private bool _clients_Deliverable_Responsibilities = false;
		public bool Clients_Deliverable_Responsibilities
			{
			get{return this._clients_Deliverable_Responsibilities;}
			set{this._clients_Deliverable_Responsibilities = value;}
			}
		private bool _deliverable_Exclusions = false;
		public bool Deliverable_Exclusions
			{
			get{return this._deliverable_Exclusions;}
			set{this._deliverable_Exclusions = value;}
			}
		private bool _deliverable_Governance_Controls = false;
		public bool Deliverable_Governance_Controls
			{
			get{return this._deliverable_Governance_Controls;}
			set{this._deliverable_Governance_Controls = value;}
			}
		private bool _reports = false;
		public bool Reports
			{
			get{return this._reports;}
			set{this._reports = value;}
			}
		private bool _report_Heading = false;
		public bool Report_Heading
			{
			get{return this._report_Heading;}
			set{this._report_Heading = value;}
			}
		private bool _report_Description = false;
		public bool Report_Description
			{
			get{return this._report_Description;}
			set{this._report_Description = value;}
			}
		private bool _report_Inputs = false;
		public bool Report_Inputs
			{
			get{return this._report_Inputs;}
			set{this._report_Inputs = value;}
			}
		private bool _report_Outputs = false;
		public bool Report_Outputs
			{
			get{return this._report_Outputs;}
			set{this._report_Outputs = value;}
			}
		private bool _dds_Report_Obligations = false;
		public bool DDs_Report_Obligations
			{
			get{return this._dds_Report_Obligations;}
			set{this._dds_Report_Obligations = value;}
			}
		private bool _clients_Report_Responsibilities = false;
		public bool Clients_Report_Responsibilities
			{
			get{return this._clients_Report_Responsibilities;}
			set{this._clients_Report_Responsibilities = value;}
			}
		private bool _report_Exclusions = false;
		public bool Report_Exclusions
			{
			get{return this._report_Exclusions;}
			set{this._report_Exclusions = value;}
			}
		private bool _report_Governance_Controls = false;
		public bool Report_Governance_Controls
			{
			get{return this._report_Governance_Controls;}
			set{this._report_Governance_Controls = value;}
			}
		private bool _meetings = false;
		public bool Meetings
			{
			get{return this._meetings;}
			set{this._meetings = value;}
			}
		private bool _meeting_Heading = false;
		public bool Meeting_Heading
			{
			get{return this._meeting_Heading;}
			set{this._meeting_Heading = value;}
			}
		private bool _meeting_Description = false;
		public bool Meeting_Description
			{
			get{return this._meeting_Description;}
			set{this._meeting_Description = value;}
			}
		private bool _meeting_Inputs = false;
		public bool Meeting_Inputs
			{
			get{return this._meeting_Inputs;}
			set{this._meeting_Inputs = value;}
			}
		private bool _meeting_Outputs = false;
		public bool Meeting_Outputs
			{
			get{return this._meeting_Outputs;}
			set{this._meeting_Outputs = value;}
			}
		private bool _dds_meeting_Obligations = false;
		public bool DDs_Meeting_Obligations
			{
			get{return this._dds_meeting_Obligations;}
			set{this._dds_meeting_Obligations = value;}
			}
		private bool _clients_Meeting_Responsibilities = false;
		public bool Clients_Meeting_Responsibilities
			{
			get{return this._clients_Meeting_Responsibilities;}
			set{this._clients_Meeting_Responsibilities = value;}
			}
		private bool _meeting_Exclusions = false;
		public bool Meeting_Exclusions
			{
			get{return this._meeting_Exclusions;}
			set{this._meeting_Exclusions = value;}
			}
		private bool _meeting_Governance_Controls = false;
		public bool Meeting_Governance_Controls
			{
			get{return this._meeting_Governance_Controls;}
			set{this._meeting_Governance_Controls = value;}
			}
		private bool _service_Level_Section = false;
		public bool Service_Level_Section
			{
			get{return this._service_Level_Section;}
			set{this._service_Level_Section = value;}
			}
		} // end of Internal_DRM_Sections class

	/// <summary>
	/// This class contains all the properties and methods for DRM (Deliverable Report Meeting) Sections
	/// </summary>
	class External_DRM_Sections : External_Document
		{
		private bool _drm_Summary = false;
		public bool DRM_Summary
			{
			get
				{
				return _drm_Summary;
				}
			set
				{
				_drm_Summary = value;
				}
			}
		private bool _drm_Section = false;
		public bool DRM_Section
			{
			get
				{
				return _drm_Section;
				}
			set
				{
				_drm_Section = value;
				}
			}
		private bool _deliverables = false;
		public bool Deliverables
			{
			get
				{
				return _deliverables;
				}
			set
				{
				_deliverables = value;
				}
			}
		private bool _deliverable_Heading = false;
		public bool Deliverable_Heading
			{

			get
				{
				return _deliverable_Heading;
				}
			set
				{
				_deliverable_Heading = value;
				}
			}
		private bool _deliverable_Description = false;
		public bool Deliverable_Description
			{
			get
				{
				return _deliverable_Description;
				}
			set
				{
				_deliverable_Description = value;
				}
			}
		private bool _deliverable_Inputs = false;
		public bool Deliverable_Inputs
			{
			get
				{
				return _deliverable_Inputs;
				}
			set
				{
				_deliverable_Inputs = value;
				}
			}
		private bool _deliverable_Outputs = false;
		public bool Deliverable_Outputs
			{
			get
				{
				return _deliverable_Outputs;
				}
			set
				{
				_deliverable_Outputs = value;
				}
			}
		private bool _dds_Deliverable_Obligations = false;
		public bool DDs_Deliverable_Obligations
			{
			get
				{
				return _dds_Deliverable_Obligations;
				}
			set
				{
				_dds_Deliverable_Obligations = value;
				}
			}
		private bool _clients_Deliverable_Responsibilities = false;
		public bool Clients_Deliverable_Responsibilities
			{
			get
				{
				return _clients_Deliverable_Responsibilities;
				}
			set
				{
				_clients_Deliverable_Responsibilities = value;
				}
			}
		private bool _deliverable_Exclusions = false;
		public bool Deliverable_Exclusions
			{
			get
				{
				return _deliverable_Exclusions;
				}
			set
				{
				_deliverable_Exclusions = value;
				}
			}
		private bool _deliverable_Governance_Controls = false;
		public bool Deliverable_Governance_Controls
			{
			get
				{
				return _deliverable_Governance_Controls;
				}
			set
				{
				_deliverable_Governance_Controls = value;
				}
			}
		private bool _reports = false;
		public bool Reports
			{
			get
				{
				return _reports;
				}
			set
				{
				_reports = value;
				}
			}
		private bool _report_Heading = false;
		public bool Report_Heading
			{

			get
				{
				return _report_Heading;
				}
			set
				{
				_report_Heading = value;
				}
			}
		private bool _report_Description = false;
		public bool Report_Description
			{
			get
				{
				return _report_Description;
				}
			set
				{
				_report_Description = value;
				}
			}
		private bool _report_Inputs = false;
		public bool Report_Inputs
			{
			get
				{
				return _report_Inputs;
				}
			set
				{
				_report_Inputs = value;
				}
			}
		private bool _report_Outputs = false;
		public bool Report_Outputs
			{
			get
				{
				return _report_Outputs;
				}
			set
				{
				_report_Outputs = value;
				}
			}
		private bool _dds_Report_Obligations = false;
		public bool DDs_Report_Obligations
			{
			get
				{
				return _dds_Report_Obligations;
				}
			set
				{
				_dds_Report_Obligations = value;
				}
			}
		private bool _clients_Report_Responsibilities = false;
		public bool Clients_Report_Responsibilities
			{
			get
				{
				return _clients_Report_Responsibilities;
				}
			set
				{
				_clients_Report_Responsibilities = value;
				}
			}
		private bool _report_Exclusions = false;
		public bool Report_Exclusions
			{
			get
				{
				return _report_Exclusions;
				}
			set
				{
				_report_Exclusions = value;
				}
			}
		private bool _report_Governance_Controls = false;
		public bool Report_Governance_Controls
			{
			get
				{
				return _report_Governance_Controls;
				}
			set
				{
				_report_Governance_Controls = value;
				}
			}
		private bool _meetings = false;
		public bool Meetings
			{
			get
				{
				return _meetings;
				}
			set
				{
				_meetings = value;
				}
			}
		private bool _meeting_Heading = false;
		public bool Meeting_Heading
			{

			get
				{
				return _meeting_Heading;
				}
			set
				{
				_meeting_Heading = value;
				}
			}
		private bool _meeting_Description = false;
		public bool Meeting_Description
			{
			get
				{
				return _meeting_Description;
				}
			set
				{
				_meeting_Description = value;
				}
			}
		private bool _meeting_Inputs = false;
		public bool Meeting_Inputs
			{
			get
				{
				return _meeting_Inputs;
				}
			set
				{
				_meeting_Inputs = value;
				}
			}
		private bool _meeting_Outputs = false;
		public bool Meeting_Outputs
			{
			get
				{
				return _meeting_Outputs;
				}
			set
				{
				_meeting_Outputs = value;
				}
			}
		private bool _dds_meeting_Obligations = false;
		public bool DDs_Meeting_Obligations
			{
			get
				{
				return _dds_meeting_Obligations;
				}
			set
				{
				_dds_meeting_Obligations = value;
				}
			}
		private bool _clients_Meeting_Responsibilities = false;
		public bool Clients_Meeting_Responsibilities
			{
			get
				{
				return _clients_Meeting_Responsibilities;
				}
			set
				{
				_clients_Meeting_Responsibilities = value;
				}
			}
		private bool _meeting_Exclusions = false;
		public bool Meeting_Exclusions
			{
			get
				{
				return _meeting_Exclusions;
				}
			set
				{
				_meeting_Exclusions = value;
				}
			}
		private bool _meeting_Governance_Controls = false;
		public bool Meeting_Governance_Controls
			{
			get
				{
				return _meeting_Governance_Controls;
				}
			set
				{
				_meeting_Governance_Controls = value;
				}
			}
		private bool _service_Level_Section = false;
		public bool Service_Level_Section
			{
			get
				{
				return _service_Level_Section;
				}
			set
				{
				_service_Level_Section = value;
				}
			}
		} // end of External_DRM_Sections class

	/// <summary>
	/// This class represent the Internal Service Definition (ISD) with inline DRM (Deliverable Report Meeting) 
	/// It inherits from the Internal_DRM_Inline Class.
	/// </summary>
	class ISD_Document_DRM_Inline : Internal_DRM_Inline
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 60:
								this.Introductory_Section = true;
								break;
							case 61:
								this.Introduction = true;
								break;
							case 62:
								this.Executive_Summary = true;
								break;
							case 63:
								this.Service_Portfolio_Section = true;
								break;
							case 64:
								this.Service_Portfolio_Description = true;
								break;
							case 65:
								this.Service_Family_Heading = true;
								break;
							case 66:
								this.Service_Family_Description = true;
								break;
							case 67:
								this.Service_Product_Heading = true;
								break;
							case 68:
								this.Service_Product_Description = true;
								break;
							case 69:
								this.Service_Product_Key_Client_Benefits = true;
								break;
							case 70:
								this.Service_Product_KeyDD_Benefits = true;
								break;
							case 71:
								this.Service_Element_Heading = true;
								break;
							case 72:
								this.Service_Element_Description= true;
								break;
							case 73:
								this.Service_Element_Objectives = true;
								break;
							case 74:
								this.Service_Element_Key_Client_Benefits = true;
								break;
							case 75:
								this.Service_Element_Key_Client_Advantages = true;
								break;
							case 76:
								this.Service_Element_Key_DD_Benefits = true;
								break;
							case 77:
								this.Service_Element_Critical_Success_Factors = true;
								break;
							case 78:
								this.Service_Element_Key_Performance_Indicators = true;
								break;
							case 79:
								this.Service_Element_High_Level_Process = true;
								break;
							case 80:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 81:
								this.DRM_Heading = true;
								break;
							case 82:
								this.DRM_Description = true;
								break;
							case 83:
								this.DRM_Inputs = true;
								break;
							case 84:
								this.DRM_Outputs= true;
								break;
							case 85:
								this.DDS_DRM_Obligations = true;
								break;
							case 86:
								this.Clients_DRM_Responsibilities = true;
								break;
							case 87:
								this.DRM_Exclusions = true;
								break;
							case 88:
								this.DRM_Governance_Controls = true;
								break;
							case 89:
								this.Service_Levels = true;
								break;
							case 90:
								this.Service_Level_Heading = true;
								break;
							case 91:
								this.Service_Level_Commitments_Table = true;
								break;
							case 92:
								this.Activities = true;
								break;
							case 93:
								this.Activity_Heading = true;
								break;
							case 94:
								this.Activity_Description_Table = true;
								break;
							case 95:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 96:
								this.Acronyms = true;
								break;
							case 97:
								this.Glossary_of_Terms = true;
								break;
							case 98:
								this.Document_Acceptance_Section = true;
                                        break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for ISD_Document_DRM_Inline's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		} // end of ISD_Document_DRM_Inline class

	/// <summary>
	/// This class represent the Framework with inline DRM (Deliverable Report Meeting) Document object
	/// It inherits from the Internal_DRM_Inline Class.
	/// </summary>
	class Services_Framework_Document_DRM_Inline : Internal_DRM_Inline
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 293:
								this.Introductory_Section = true;
								break;
							case 294:
								this.Introduction = true;
								break;
							case 295:
								this.Executive_Summary = true;
								break;
							case 296:
								this.Service_Portfolio_Section = true;
								break;
							case 297:
								this.Service_Portfolio_Description = true;
								break;
							case 298:
								this.Service_Family_Heading = true;
								break;
							case 299:
								this.Service_Family_Description = true;
								break;
							case 300:
								this.Service_Product_Heading = true;
								break;
							case 301:
								this.Service_Product_Description = true;
								break;
							case 302:
								this.Service_Product_Key_Client_Benefits = true;
								break;
							case 303:
								this.Service_Product_KeyDD_Benefits = true;
								break;
							case 304:
								this.Service_Element_Heading = true;
								break;
							case 305:
								this.Service_Element_Description = true;
								break;
							case 306:
								this.Service_Element_Objectives = true;
								break;
							case 307:
								this.Service_Element_Key_Client_Benefits = true;
								break;
							case 308:
								this.Service_Element_Key_Client_Advantages = true;
								break;
							case 309:
								this.Service_Element_Key_DD_Benefits = true;
								break;
							case 311:
								this.Service_Element_Critical_Success_Factors = true;
								break;
							case 312:
								this.Service_Element_Key_Performance_Indicators = true;
								break;
							case 313:
								this.Service_Element_High_Level_Process = true;
								break;
							case 314:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 315:
								this.DRM_Heading = true;
								break;
							case 316:
								this.DRM_Description = true;
								break;
							case 317:
								this.DRM_Inputs = true;
								break;
							case 318:
								this.DRM_Outputs = true;
								break;
							case 319:
								this.DDS_DRM_Obligations = true;
								break;
							case 320:
								this.Clients_DRM_Responsibilities = true;
								break;
							case 321:
								this.DRM_Exclusions = true;
								break;
							case 322:
								this.DRM_Governance_Controls = true;
								break;
							case 323:
								this.Service_Levels = true;
								break;
							case 324:
								this.Service_Level_Heading = true;
								break;
							case 325:
								this.Service_Level_Commitments_Table = true;
								break;
							case 326:
								this.Activities = true;
								break;
							case 327:
								this.Activity_Heading = true;
								break;
							case 328:
								this.Activity_Description_Table = true;
								break;
							case 329:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 330:
								this.Acronyms = true;
								break;
							case 331:
								this.Glossary_of_Terms = true;
								break;
							case 332:
								this.Document_Acceptance_Section = true;
                                        break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for Services_Framework_Document_DRM_Inline's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		} // end of Services_Framework_Document_DRM_Inline class

	/// <summary>
	/// This class represent the Client Service Description (CSD) with sperate DRM (Deliverable Report Meeting) sections
	/// It inherits from the DRM Sections Class.
	/// </summary>
	class CSD_Document_DRM_Sections : External_DRM_Sections
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 99:
								this.Introductory_Section = true;
								break;
							case 100:
								this.Introduction = true;
								break;
							case 101:
								this.Executive_Summary = true;
								break;
							case 102:
								this.Service_Portfolio_Section = true;
								break;
							case 103:
								this.Service_Portfolio_Description = true;
								break;
							case 104:
								this.Service_Family_Heading = true;
								break;
							case 105:
								this.Service_Family_Description = true;
								break;
							case 106:
								this.Service_Product_Heading = true;
								break;
							case 107:
								this.Service_Product_Description = true;
								break;
							case 108:
								this.Service_Feature_Heading = true;
								break;
							case 109:
								this.Service_Feature_Description = true;
								break;
							case 110:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 111:
								this.DRM_Heading = true;
								break;
							case 112:
								this.DRM_Summary = true;
								break;
							case 113:
								this.Service_Levels = true;
								break;
							case 114:
								this.Service_Level_Heading = true;
								break;
							case 115:
								this.Service_Level_Commitments_Table = true;
								break;
							case 116:
								this.DRM_Section = true;
								break;
							case 117:
								this.Deliverables = true;
								break;
							case 118:
								this.Deliverable_Heading = true;
								break;
							case 119:
								this.Deliverable_Description = true;
								break;
							case 120:
								this.DDs_Deliverable_Obligations = true;
								break;
							case 121:
								this.Clients_Deliverable_Responsibilities = true;
								break;
							case 122:
								this.Deliverable_Exclusions = true;
								break;
							case 123:
								this.Deliverable_Governance_Controls = true;
								break;
							case 124:
								this.Reports = true;
								break;
							case 125:
								this.Report_Heading = true;
								break;
							case 126:
								this.Report_Description = true;
								break;
							case 127:
								this.DDs_Report_Obligations = true;
								break;
							case 128:
								this.Clients_Report_Responsibilities = true;
								break;
							case 129:
								this.Report_Exclusions = true;
								break;
							case 130:
								this.Report_Governance_Controls = true;
								break;
							case 131:
								this.Meetings = true;
								break;
							case 132:
								this.Meeting_Heading = true;
								break;
							case 133:
								this.Meeting_Description = true;
								break;
							case 134:
								this.DDs_Meeting_Obligations = true;
								break;
							case 135:
								this.Clients_Meeting_Responsibilities = true;
								break;
							case 136:
								this.Meeting_Exclusions = true;
								break;
							case 137:
								this.Meeting_Governance_Controls = true;
								break;
							case 138:
								this.Service_Level_Section = true;
								break;
							case 139:
								this.Service_Level_Heading = true;
								break;
							case 140:
								this.Service_Level_Commitments_Table = true;
								break;
							case 141:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 142:
								this.Acronyms = true;
								break;
							case 143:
								this.Glossary_of_Terms = true;
								break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for CSD_Document_DRM_Sections's Generate method
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		} // end of CSD_Document_DRM_Sections class

	/// <summary>
	/// This class represent the Internal Service Definition (ISD) with sperate DRM (Deliverable Report Meeting) sections
	/// It inherits from the Internal DRM Sections Class.
	/// </summary>
	class ISD_Document_DRM_Sections : Internal_DRM_Sections
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 1:
								this.Introductory_Section = true;
								break;
							case 2:
								this.Introduction = true;
								break;
							case 3:
								this.Executive_Summary = true;
								break;
							case 4:
								this.Service_Portfolio_Section = true;
								break;
							case 5:
								this.Service_Portfolio_Description = true;
								break;
							case 6:
								this.Service_Family_Heading = true;
								break;
							case 7:
								this.Service_Family_Description = true;
								break;
							case 8:
								this.Service_Product_Heading = true;
								break;
							case 9:
								this.Service_Product_Description = true;
								break;
							case 10:
								this.Service_Product_Key_Client_Benefits = true;
								break;
							case 11:
								this.Service_Product_KeyDD_Benefits = true;
								break;
							case 12:
								this.Service_Element_Heading = true;
								break;
							case 13:
								this.Service_Element_Description = true;
								break;
							case 14:
								this.Service_Element_Objectives = true;
								break;
							case 15:
								this.Service_Element_Key_Client_Benefits = true;
								break;
							case 16:
								this.Service_Element_Key_Client_Advantages = true;
								break;
							case 17:
								this.Service_Element_Key_DD_Benefits = true;
								break;
							case 18:
								this.Service_Element_Critical_Success_Factors = true;
								break;
							case 19:
								this.Service_Element_Key_Performance_Indicators = true;
								break;
							case 20:
								this.Service_Element_High_Level_Process = true;
								break;
							case 21:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 27:
								this.DRM_Heading = true;
								break;
							case 22:
								this.DRM_Summary = true;
								break;
							case 23:
								this.Service_Levels = true;
								break;
							case 24:
								this.Service_Level_Heading = true;
								break;
							case 25:
								this.Service_Level_Commitments_Table = true;
								break;
							case 26:
								this.Activities = true;
								break;
							case 28:
								this.Activity_Heading = true;
								break;
							case 29:
								this.Activity_Description_Table = true;
								break;
							case 32:
								this.DRM_Section = true;
								break;
							case 33:
								this.Deliverables = true;
								break;
							case 34:
								this.Deliverable_Heading = true;
								break;
							case 35:
								this.Deliverable_Description = true;
								break;
							case 36:
								this.Deliverable_Inputs = true;
								break;
							case 37:
								this.Deliverable_Outputs = true;
								break;
							case 38:
								this.DDs_Deliverable_Obligations = true;
								break;
							case 39:
								this.Clients_Deliverable_Responsibilities = true;
								break;
							case 40:
								this.Deliverable_Exclusions = true;
								break;
							case 41:
								this.Deliverable_Governance_Controls = true;
								break;
							case 42:
								this.Reports = true;
								break;
							case 43:
								this.Report_Heading = true;
								break;
							case 44:
								this.Report_Description = true;
								break;
							case 45:
								this.DDs_Report_Obligations = true;
								break;
							case 46:
								this.Clients_Report_Responsibilities = true;
								break;
							case 47:
								this.Report_Exclusions = true;
								break;
							case 48:
								this.Report_Governance_Controls = true;
								break;
							case 49:
								this.Meetings = true;
								break;
							case 50:
								this.Meeting_Heading = true;
								break;
							case 51:
								this.Meeting_Description = true;
								break;
							case 52:
								this.DDs_Meeting_Obligations = true;
								break;
							case 53:
								this.Clients_Meeting_Responsibilities = true;
								break;
							case 54:
								this.Meeting_Exclusions = true;
								break;
							case 55:
								this.Meeting_Governance_Controls = true;
								break;
							case 56:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 57:
								this.Acronyms = true;
								break;
							case 58:
								this.Glossary_of_Terms = true;
								break;
							case 59:
								this.Document_Acceptance_Section = true;
								break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to added for ISD_Document_DRM_Sections's Generate method
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		} // end of ISD_Document_DRM_Sections class

	/// <summary>
	/// This class represent the Services Framework Document with sperate DRM (Deliverable Report Meeting) sections
	/// It inherits from the Internal DRM Sections Class.
	/// </summary>
	class Services_Framework_Document_DRM_Sections : Internal_DRM_Sections
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
	     /// which represents the options the user selected and transposing the values by
	     /// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 236:
								this.Introductory_Section = true;
								break;
							case 237:
								this.Introduction = true;
								break;
							case 238:
								this.Executive_Summary = true;
								break;
							case 239:
								this.Service_Portfolio_Section = true;
								break;
							case 240:
								this.Service_Portfolio_Description = true;
								break;
							case 241:
								this.Service_Family_Heading = true;
								break;
							case 242:
								this.Service_Family_Description = true;
								break;
							case 243:
								this.Service_Product_Heading = true;
								break;
							case 244:
								this.Service_Product_Description = true;
								break;
                                   case 245:
								this.Service_Product_Key_Client_Benefits = true;
								break;
							case 246:
								this.Service_Product_KeyDD_Benefits = true;
								break;
							case 247:
								this.Service_Element_Heading = true;
								break;
							case 248:
								this.Service_Element_Description = true;
								break;
							case 249:
								this.Service_Element_Objectives = true;
								break;
							case 250:
								this.Service_Element_Key_Client_Benefits = true;
								break;
							case 251:
								this.Service_Element_Key_Client_Advantages = true;
								break;
							case 252:
								this.Service_Element_Key_DD_Benefits = true;
								break;
							case 253:
								this.Service_Element_Critical_Success_Factors = true;
								break;
							case 254:
								this.Service_Element_Key_Performance_Indicators = true;
								break;
							case 255:
								this.Service_Element_High_Level_Process = true;
								break;
							case 256:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 257:
								this.DRM_Heading = true;
								break;
							case 258:
								this.DRM_Summary = true;
								break;
							case 359:
								this.Service_Levels = true;
								break;
							case 260:
								this.Service_Level_Heading = true;
								break;
							case 261:
								this.Service_Level_Commitments_Table = true;
								break;
							case 262:
								this.Activities = true;
								break;
							case 263:
								this.Activity_Heading = true;
								break;
							case 264:
								this.Activity_Description_Table = true;
								break;
							case 267:
								this.DRM_Section = true;
								break;
							case 268:
								this.Deliverables = true;
								break;
							case 269:
								this.Deliverable_Heading = true;
								break;
							case 333:
								this.Deliverable_Description = true;
                                        break;
							case 270:
								this.Deliverable_Inputs = true;
								break;
							case 271:
								this.Deliverable_Outputs = true;
								break;
							case 272:
								this.DDs_Deliverable_Obligations = true;
								break;
							case 273:
								this.Clients_Deliverable_Responsibilities = true;
								break;
							case 274:
								this.Deliverable_Exclusions = true;
								break;
							case 275:
								this.Deliverable_Governance_Controls = true;
								break;
							case 276:
								this.Reports = true;
								break;
							case 277:
								this.Report_Heading = true;
								break;
							case 278:
								this.Report_Description = true;
								break;
							case 279:
								this.DDs_Report_Obligations = true;
								break;
							case 280:
								this.Clients_Report_Responsibilities = true;
								break;
							case 281:
								this.Report_Exclusions = true;
								break;
							case 282:
								this.Report_Governance_Controls = true;
								break;
							case 283:
								this.Meetings = true;
								break;
							case 284:
								this.Meeting_Heading = true;
								break;
							case 285:
								this.Meeting_Description = true;
								break;
							case 286:
								this.DDs_Meeting_Obligations = true;
								break;
							case 287:
								this.Clients_Meeting_Responsibilities = true;
								break;
							case 288:
								this.Meeting_Exclusions = true;
								break;
							case 289:
								this.Meeting_Governance_Controls = true;
								break;
							case 290:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 291:
								this.Acronyms = true;
								break;
							case 292:
								this.Glossary_of_Terms = true;
								break;
							case 293:
								this.Document_Acceptance_Section = true;
								break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t Begin to generate {0}", this.DocumentType);
			DateTime timeStarted = DateTime.Now;
			string hyperlinkImageRelationshipID = "";
			string documentCollection_HyperlinkURL = "";
			string currentListURI = "";
			string currentHyperlinkViewEditURI = "";
			string currentContentLayer = "None";
			bool drmHeading = false;
			Table objActivityTable = new Table();
			Table objServiceLevelTable = new Table();
			Dictionary<int, string> dictDeliverables = new Dictionary<int, string>();
			Dictionary<int, string> dictReports = new Dictionary<int, string>();
			Dictionary<int, string> dictMeetings = new Dictionary<int, string>();
			Dictionary<int, string> dictSLAs = new Dictionary<int, string>();

			if(this.HyperlinkEdit)
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.EditFormURI + this.DocumentCollectionID;
				currentHyperlinkViewEditURI = Properties.AppResources.EditFormURI;
			if(this.Hyperlink_View)
				documentCollection_HyperlinkURL = Properties.AppResources.SharePointSiteURL +
					Properties.AppResources.List_DocumentCollectionLibraryURI +
					Properties.AppResources.DisplayFormURI + this.DocumentCollectionID;
				currentHyperlinkViewEditURI = Properties.AppResources.DisplayFormURI;
			int tableCaptionCounter = 1;
			int imageCaptionCounter = 1;
			int hyperlinkCounter = 4;

			//Initialize the Data access to SharePoint
			DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new
				Uri(Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri));
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			datacontexSDDP.MergeOption = System.Data.Services.Client.MergeOption.NoTracking;
			
			// define a new objOpenXMLdocument
			oxmlDocument objOXMLdocument = new oxmlDocument();
			// use CreateDocumentFromTemplate method to create a new MS Word Document based on the relevant template
			if(objOXMLdocument.CreateDocumentFromTemplate(parTemplateURL: this.Template, parDocumentType: this.DocumentType))
				{
				Console.WriteLine("\t\t objOXMLdocument:\n" +
				"\t\t\t+ LocalDocumentPath: {0}\n" +
				"\t\t\t+ DocumentFileName.: {1}\n" +
				"\t\t\t+ DocumentURI......: {2}", objOXMLdocument.LocalDocumentPath, objOXMLdocument.DocumentFilename, objOXMLdocument.LocalDocumentURI);
				}
			else
				{
				// if the creation failed.
				Console.WriteLine("An ERROR occurred and the new MS Word Document could not be created due to above stated ERROR conditions.");
				this.ErrorMessages.Add("Application was unable to create the document based on the template - Check the Output log.");
				return false;
				}

			if(this.SelectedNodes == null || this.SelectedNodes.Count < 1)
				{
				Console.WriteLine("\t\t\t *** There are 0 selected nodes to generate");
				this.ErrorMessages.Add("There are no Selected Nodes to generate.");
				return false;
				}
			// Create and open the new Document
			try  {
				// Open the MS Word document in Edit mode
				WordprocessingDocument objWPdocument = WordprocessingDocument.Open(path: objOXMLdocument.LocalDocumentURI, isEditable: true);
				// Define all open XML object to use for building the document
				MainDocumentPart objMainDocumentPart = objWPdocument.MainDocumentPart;
				Body objBody = objWPdocument.MainDocumentPart.Document.Body;          // Define the objBody of the document
				Paragraph objParagraph = new Paragraph();
				ParagraphProperties objParaProperties = new ParagraphProperties();
				Run objRun = new Run();
				RunProperties objRunProperties = new RunProperties();
				Text objText = new Text();
				HTMLdecoder objHTMLdecoder = new HTMLdecoder();
				objHTMLdecoder.WPbody = objBody;

				// Determine the Page Size for the current Body object.
				SectionProperties objSectionProperties = new SectionProperties();
				this.PageWith = Convert.ToUInt32(Properties.AppResources.DefaultPageWidth);
				this.PageHight = Convert.ToUInt32(Properties.AppResources.DefaultPageHeight);
				
				if(objBody.GetFirstChild<SectionProperties>() != null)
					{
					objSectionProperties = objBody.GetFirstChild<SectionProperties>();
					PageSize objPageSize = objSectionProperties.GetFirstChild<PageSize>();
					PageMargin objPageMargin = objSectionProperties.GetFirstChild<PageMargin>();
					if(objPageSize != null)
						{
						this.PageWith = objPageSize.Width;
						this.PageHight = objPageSize.Height;
						Console.WriteLine("\t\t Page width x height: {0} x {1} twips", this.PageWith, this.PageHight);
						}
					if(objPageMargin != null)
						{
						if(objPageMargin.Left != null)
							{
							this.PageWith -= objPageMargin.Left;
							Console.WriteLine("\t\t\t - Left Margin..: {0} twips", objPageMargin.Left);
							}
						if(objPageMargin.Right != null)
							{
							this.PageWith -= objPageMargin.Right;
							Console.WriteLine("\t\t\t - Right Margin.: {0} twips", objPageMargin.Right);
							}
						if(objPageMargin.Top != null)
							{
							string tempTop = objPageMargin.Top.ToString();
							Console.WriteLine("\t\t\t - Top Margin...: {0} twips", tempTop);
							this.PageHight -= Convert.ToUInt32(tempTop);
							}
						if(objPageMargin.Bottom != null)
							{
							string tempBottom = objPageMargin.Bottom.ToString();
							Console.WriteLine("\t\t\t - Bottom Margin: {0} twips", tempBottom);
							this.PageHight -= Convert.ToUInt32(tempBottom);
							}
						}
					}
				
				Console.WriteLine("\t\t Effective pageWidth x pageHeight.: {0} x {1} twips", this.PageWith, this.PageHight);

				// Check whether Hyperlinks need to be included
				if(this.HyperlinkEdit || this.Hyperlink_View)
					{
					//Insert and embed the hyperlink image in the document and keep the Image's Relationship ID in a variable for repeated use
					hyperlinkImageRelationshipID = oxmlDocument.InsertHyperlinkImage(parMainDocumentPart: ref objMainDocumentPart);
					}
				//--------------------------------------------------
				// Insert the Introductory Section
				if(this.Introductory_Section)
					{
					objParagraph = oxmlDocument.Insert_Section();
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_IntruductorySection_HeadingText, 
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}
				//--------------------------------------------------
				// Insert the Introduction
				if(this.Introduction)
					{
					objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Introduction_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					if(this.IntroductionRichText != null)
						{
						objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 1,
							parHTML2Decode: this.IntroductionRichText,
							parTableCaptionCounter: ref tableCaptionCounter,
							parImageCaptionCounter: ref imageCaptionCounter,
							parHyperlinkID: ref hyperlinkCounter,
							parPageHeightTwips: this.PageHight,
							parPageWidthTwips: this.PageWith);
						}
					}
				//--------------------------------------------------
				// Insert the Executive Summary
				if(this.Executive_Summary)
					{
					objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ExecutiveSummary_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					if(this.ExecutiveSummaryRichText != null)
						{
						objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 1,
							parHTML2Decode: this.ExecutiveSummaryRichText,
							parTableCaptionCounter: ref tableCaptionCounter,
							parImageCaptionCounter: ref imageCaptionCounter,
							parHyperlinkID: ref hyperlinkCounter,
							parPageHeightTwips: this.PageHight,
							parPageWidthTwips: this.PageWith);
						}

					}
				//--------------------------------------------------
				// Insert the user selected content
				if(this.SelectedNodes.Count <= 0)
					goto Glossary_and_Acronyms;
				foreach(Hierarchy node in this.SelectedNodes)
					{
					Console.WriteLine("Node: {0} - {1} {2} {3}", node.Sequence, node.Level, node.NodeType, node.NodeID);
					switch(node.NodeType)
						{
						case enumNodeTypes.FRA:	// Service Framework
						case enumNodeTypes.POR:  //Service Portfolio
							{
							if(this.Service_Portfolio_Section)
								{
								try
									{
									var rsPortfolios = 
									from dsPortfolio in datacontexSDDP.ServicePortfolios
									where dsPortfolio.Id == node.NodeID
									select new
										{dsPortfolio.Id,
										dsPortfolio.Title,
										dsPortfolio.ISDHeading,
										dsPortfolio.ISDDescription};

									var recPortfolio = rsPortfolios.FirstOrDefault();
									
									Console.WriteLine("\t\t + {0} - {1}", recPortfolio.Id , recPortfolio.Title);
									objParagraph = oxmlDocument.Insert_Section();
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recPortfolio.ISDHeading,
										parIsNewSection: true);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL + 
												Properties.AppResources.List_ServicePortfoliosURI + 
												currentHyperlinkViewEditURI + recPortfolio.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Porfolio Description
									if(this.Service_Portfolio_Description)
										{
										if(recPortfolio.ISDDescription != null)
											{
												currentListURI = Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServicePortfoliosURI +
													currentHyperlinkViewEditURI + recPortfolio.Id;
												objHTMLdecoder.DecodeHTML(
													parMainDocumentPart: ref objMainDocumentPart,
													parDocumentLevel: 0,
													parHTML2Decode: recPortfolio.ISDDescription,
													parTableCaptionCounter: ref tableCaptionCounter,
													parImageCaptionCounter: ref imageCaptionCounter,
													parHyperlinkID: ref hyperlinkCounter,
													parPageHeightTwips: this.PageHight,
													parPageWidthTwips: this.PageWith);
											}
										}
									} //Try
								catch(DataServiceQueryException exc)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Portfolio ID " + node.NodeID +
										" doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Section();
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Portfolio " + node.NodeID + " is missing.",
										parIsNewSection: true,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									}
								} // //if(this.Service_Portfolio_Section)
							break;
							}
						case enumNodeTypes.FAM:  // Service Family
							{
							if(this.Service_Family_Heading)
								{
								try
									{
									var rsFamilies =
										from rsFamily in datacontexSDDP.ServiceFamilies
										where rsFamily.Id == node.NodeID
										select new
											{
											rsFamily.Id,
											rsFamily.Title,
											rsFamily.ISDHeading,
											rsFamily.ISDDescription
											};

									var recFamily = rsFamilies.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recFamily.Id, recFamily.Title);
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recFamily.ISDHeading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_ServiceFamiliesURI +
											currentHyperlinkViewEditURI + recFamily.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Family Description
									if(this.Service_Family_Description)
										{
										if(recFamily.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServicePortfoliosURI +
												currentHyperlinkViewEditURI +
												recFamily.Id;
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: recFamily.ISDDescription,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									} // Try
								catch (DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Family ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									break;
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									}
								} // //if(this.Service_Portfolio_Section)
							break;
							}
						case enumNodeTypes.PRO:  // Service Product
							{
							if(this.Service_Product_Heading)
								{
								try
									{
									var rsProducts =
										from rsProduct in datacontexSDDP.ServiceProducts
										where rsProduct.Id == node.NodeID
										select new
											{
											rsProduct.Id,
											rsProduct.Title,
											rsProduct.ISDHeading,
											rsProduct.ISDDescription,
											rsProduct.KeyClientBenefits,
											rsProduct.KeyDDBenefits
											};

									var recProduct = rsProducts.FirstOrDefault();

									Console.WriteLine("\t\t + {0} - {1}", recProduct.Id, recProduct.Title);
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recProduct.ISDHeading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
											Properties.AppResources.List_ServiceProductsURI +
											currentHyperlinkViewEditURI + recProduct.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Product Description
									if(this.Service_Product_Description)
										{
										if(recProduct.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												recProduct.Id;

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: recProduct.ISDDescription,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Product_KeyDD_Benefits)
										{
										if(recProduct.KeyDDBenefits != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												recProduct.Id;
											Console.WriteLine("\t\t + {0} - {1}", recProduct.Id, Properties.AppResources.Document_Product_KeyDD_Benefits);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_KeyDD_Benefits,
												parIsNewSection: false);
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parClickLinkURL: Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id,
													parHyperlinkID: hyperlinkCounter);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recProduct.KeyDDBenefits,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}

									if(this.Service_Product_Key_Client_Benefits)
										{
										if(recProduct.KeyClientBenefits != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceProductsURI +
												currentHyperlinkViewEditURI +
												recProduct.Id;

											Console.WriteLine("\t\t + {0} - {1}", recProduct.Id,
												Properties.AppResources.Document_Product_ClientKeyBenefits);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Product_ClientKeyBenefits,
												parIsNewSection: false);
											// Check if a hyperlink must be inserted
											if(documentCollection_HyperlinkURL != "")
												{
												hyperlinkCounter += 1;
												Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
													parMainDocumentPart: ref objMainDocumentPart,
													parImageRelationshipId: hyperlinkImageRelationshipID,
													parClickLinkURL: Properties.AppResources.SharePointURL +
													Properties.AppResources.List_ServiceProductsURI +
													currentHyperlinkViewEditURI + recProduct.Id,
													parHyperlinkID: hyperlinkCounter);
												objRun.Append(objDrawing);
												}
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 2,
												parHTML2Decode: recProduct.KeyClientBenefits,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									}
								catch(DataServiceClientException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Service Product ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Service Family " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									}
								} //if(this.Service_Product_Heading)
							break;
							}
						case enumNodeTypes.ELE:  // Service Element
							{
							if(this.Service_Element_Heading)
								{
								try
									{
									// Obtain the Element info from SharePoint
									var rsElements =
										from dsElement in datacontexSDDP.ServiceElements
										where dsElement.Id == node.NodeID
										select new
											{dsElement.Id, dsElement.Title, dsElement.ISDHeading, dsElement.ISDDescription,
											dsElement.Objective, dsElement.KeyClientAdvantages, dsElement.KeyClientBenefits,
											dsElement.KeyDDBenefits, dsElement.KeyPerformanceIndicators, dsElement.CriticalSuccessFactors,
											dsElement.ProcessLink
											};
									
									var recElement = rsElements.FirstOrDefault();
									
									Console.WriteLine("\t\t + {0} - {1}", recElement.Id, recElement.Title);
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: recElement.ISDHeading,
										parIsNewSection: false);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI + recElement.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Service Service Element Description
									if(this.Service_Element_Description)
										{
										if(recElement.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recElement.ISDDescription,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Objectives)
										{
										if(recElement.Objective != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id, 
												Properties.AppResources.Document_Element_Objectives);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_Objectives,
												parIsNewSection: false);

											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.Objective,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parContentLayer: currentContentLayer,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}

									if(this.Service_Element_Critical_Success_Factors)
										{
										if(recElement.CriticalSuccessFactors != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_CriticalSuccessFactors);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_CriticalSuccessFactors,
												parIsNewSection: false);

											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.CriticalSuccessFactors,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_Client_Advantages)
										{
										if(recElement.KeyClientAdvantages != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_ClientKeyAdvantages);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_ClientKeyAdvantages,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.KeyClientAdvantages,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_Client_Benefits)
										{
										if(recElement.KeyClientBenefits != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_ClientKeyBenefits);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_ClientKeyBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.KeyClientBenefits,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_DD_Benefits)
										{
										if(recElement.KeyDDBenefits != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_KeyDDBenefits);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_KeyDDBenefits,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.KeyDDBenefits,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_Key_Performance_Indicators)
										{
										if(recElement.KeyPerformanceIndicators != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_KPI);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_KPI,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recElement.KeyPerformanceIndicators,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											}
										}
									if(this.Service_Element_High_Level_Process)
										{
										if(recElement.ProcessLink != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ServiceElementsURI +
												currentHyperlinkViewEditURI +
												recElement.Id;
											// Insert the heading
											Console.WriteLine("\t\t + {0} - {1}", recElement.Id,
												Properties.AppResources.Document_Element_KPI);
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_Element_HighLevelProcess,
												parIsNewSection: false);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											//TODO: Insert generate hypelink in oxmlEncoder

											}
										}
									drmHeading = false;
									}
                                        catch(DataServiceClientException)
										{
										// If the entry is not found - write an error in the document and record an error in the error log.
										this.LogError("Error: The Service Element ID " + node.NodeID
											+ " doesn't exist in SharePoint and couldn't be retrieved.");
										objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3);
										objRun = oxmlDocument.Construct_RunText(
											parText2Write: "Error: Service Element " + node.NodeID + " is missing.",
											parIsNewSection: false,
											parIsError: true);
										objParagraph.Append(objRun);
										}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								
								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
								} // if (this.Service_Element_Heading)
							break;
							}
						case enumNodeTypes.ELD:  // Deliverable associated with Element
						case enumNodeTypes.ELR:  // Report deliverable associated with Element
						case enumNodeTypes.ELM:  // Meeting deliverable associated with Element
							{
							if(this.DRM_Heading)
								{
								if(drmHeading == false)
									{
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: Properties.AppResources.Document_DeliverableReportsMeetings_Heading);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									drmHeading = true;
									}
								}
								try
									{
									// Obtain the Deliverable info from SharePoint
									var rsDeliverables =
										from dsDeliverable in datacontexSDDP.Deliverables
										where dsDeliverable.Id == node.NodeID
										select new
											{dsDeliverable.Id, dsDeliverable.Title, dsDeliverable.ISDHeading, dsDeliverable.ISDSummary
											};
									
									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									if(node.NodeType == enumNodeTypes.ELD)
										{
										if(dictDeliverables.ContainsKey(recDeliverable.Id) != true)
											dictDeliverables.Add(recDeliverable.Id, recDeliverable.ISDHeading);
										}
									else if(node.NodeType == enumNodeTypes.ELR)
										{
										if(dictReports.ContainsKey(recDeliverable.Id) != true)
											dictReports.Add(recDeliverable.Id, recDeliverable.ISDHeading);
										}
									else if(node.NodeType == enumNodeTypes.ELM)
										{
										if(dictMeetings.ContainsKey(recDeliverable.Id) != true)
											dictMeetings.Add(recDeliverable.Id, recDeliverable.ISDHeading);
										}
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
									}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									// Check if the user specified to include the Deliverable Description
									if(this.DRM_Summary)
										{
										if(recDeliverable.ISDSummary != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
											objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDSummary);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											}
										} // if(this.DeliverableSummary

									// Insert the Hyperlink to the relevant position in the DRM Section.
									objParagraph = oxmlDocument.Construct_BookmarkHyperlink(
										parBodyTextLevel: 5,
										parBookmarkValue: "Deliverable_" + recDeliverable.Id);
									objBody.Append(objParagraph);
									}
                                        catch(DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + node.NodeID
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + node.NodeID
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
                                        catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
							break;
							}

						case enumNodeTypes.EAC:  // Activity associated with Deliverable pertaining to Service Element
							{
							if(this.Activities)
								{
								objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 6);
								objRun = oxmlDocument.Construct_RunText(
									parText2Write: Properties.AppResources.Document_Activities_Heading);
								objParagraph.Append(objRun);
								objBody.Append(objParagraph);
								try
									{
									// Obtain the Deliverable info from SharePoint
									var rsActivities =
										from rsActivity in datacontexSDDP.Activities
										where rsActivity.Id == node.NodeID
										select new
											{
											rsActivity.Id, rsActivity.Title, rsActivity.ISDHeading, rsActivity.ISDDescription,
											rsActivity.ActivityInput, rsActivity.ActivityOutput, rsActivity.ActivityOptionalityValue,
											rsActivity.ActivityAssumptions
											};
									
									var recActivity = rsActivities.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recActivity.Id, recActivity.Title);

									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recActivity.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_ActvitiesURI +
												currentHyperlinkViewEditURI + recActivity.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Deliverable Description
									if(this.Activity_Description_Table)
										{
										// Initialize the Activities table
										if(objActivityTable.HasChildren)
											objActivityTable.RemoveAllChildren();

										objActivityTable = oxmlDocument.ConstructTable(
											parPageWidth: this.PageWith,
											parFirstRow: false,
											parNoVerticalBand: true,
											parNoHorizontalBand: true);
										TableRow objTableRow = new TableRow();
										TableCell objTableCell = new TableCell();
										string tableText = "";
										TableGrid objTableGrid = new TableGrid();
										List<UInt32> lstTableColumns = new List<UInt32>();
										lstTableColumns.Add(this.PageWith * 20 / 100);
										lstTableColumns.Add(this.PageWith * 80 / 100);
												
										objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
										// Append the TableGrid object instance to the Table object instance
										objActivityTable.Append(objTableGrid);

										// Create the Activity Description row for the table
										objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: false);
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [0], parIsFirstRow: false);
										// Add the Activity Description Title in the first Column
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = Properties.AppResources.Document_ActivityTable_RowTitle_Description;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										// Add the Activity Description value in the second Column
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [1], parIsFirstRow: false);
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = recActivity.ISDDescription;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										objActivityTable.Append(objTableRow);

										// Create the Activity Input row for the table
										objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: false);
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [0], parIsFirstRow: false);
										// Add the Activity Description Title in the first Column
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = Properties.AppResources.Document_ActivityTable_RowTitle_Inputs;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										// Add the Activity Description value in the second Column
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [1], parIsFirstRow: false);
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = recActivity.ActivityInput;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										objActivityTable.Append(objTableRow);

										// Create the Activity Outputs row for the table
										objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: false);
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [0], parIsFirstRow: false);
										// Add the Activity Description Title in the first Column
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = Properties.AppResources.Document_ActivityTable_RowTitle_Outputs;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										// Add the Activity Description value in the second Column
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [1], parIsFirstRow: false);
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = recActivity.ActivityOutput;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										objActivityTable.Append(objTableRow);

										// Create the Activity Assumptions row for the table
										objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: false);
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [0], parIsFirstRow: false);
										// Add the Activity Description Title in the first Column
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = Properties.AppResources.Document_ActivityTable_RowTitle_Assumptions;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										// Add the Activity Description value in the second Column
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [1], parIsFirstRow: false);
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = recActivity.ActivityAssumptions;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										objActivityTable.Append(objTableRow);

										// Create the Activity Optionality row for the table
										objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: false);
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [0], parIsFirstRow: false);
										// Add the Activity Description Title in the first Column
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = Properties.AppResources.Document_ActivityTable_RowTitle_Optionality;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										// Add the Activity Description value in the second Column
										objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns [1], parIsFirstRow: false);
										objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
										tableText = recActivity.ActivityOptionalityValue;
										objRun = oxmlDocument.Construct_RunText(tableText);
										objParagraph.Append(objRun);
										objTableCell.Append(objParagraph);
										objTableRow.Append(objTableCell);
										objActivityTable.Append(objTableRow);
												
										// Insert the Activities Description Table
										objBody.Append(objActivityTable);
										Console.WriteLine("\t Generated the Table with Activities for {0} - {1}", 
											recActivity.Id, recActivity.Title);
										} // if (this.Activity_Description_Table)

									} // try
                                        catch (DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Activity ID " + node.NodeID
										+ " doesn't exist in SharePoint and it couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 7);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Activity " + node.NodeID + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									break;
									}

								catch(Exception exc)
									{
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
								} // if (this.Activities)
							break;	
							}
						case enumNodeTypes.ESL:  // Service Level associated with Deliverable pertaining to Service Element
							{
							if(this.Service_Level_Heading)
								Console.WriteLine("Service Level goes here");

							break;
							}
						}
					} // foreach(Hierarchy node in this.SelectedNodes)

				//------------------------------------------------------
				// Insert the Deliverable, Report, Meeting (DRM) Section
				if(this.DRM_Section)
					{
					//--------------------------------------------------
					// Insert the Deliverables, Reports and Meetings Section
					objParagraph = oxmlDocument.Insert_Section();
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_DRM_Section_Text,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					if(this.Deliverables)
						{
						objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Deliverables_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Deliverable_";
						// Insert the individual Deliverables in the section
						foreach(KeyValuePair<int, string> deliverableItem in dictDeliverables.OrderBy(key => key.Value))
							{
							if(this.Deliverable_Heading)
								{
								try
									{
									// Obtain the Deliverable info from SharePoint
									var dsDeliverables = datacontexSDDP.Deliverables
										.Expand(p => p.GlossaryAndAcronyms);

									var rsDeliverables =
										from dsDeliverable in dsDeliverables
										where dsDeliverable.Id == deliverableItem.Key
										select dsDeliverable;

									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);
									
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark+recDeliverable.Id);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Deliverable Description
									if(this.Deliverable_Description)
										{
										if(recDeliverable.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recDeliverable.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.ISDDescription != null)
										} //if(this.Deliverable_Description)

									// Check if the user specified to include the Deliverable Inputs
									if(this.Deliverable_Inputs)
										{
										if(recDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Inputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Inputs != null)
										} //if(this.Deliverable_Inputs)

									// Check if the user specified to include the Deliverable Outputs
									if(this.Deliverable_Outputs)
										{
										if(recDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Outputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Outputs != null)
										} //if(this.Deliverable_Outputs)

									// Check if the user specified to include the Deliverable DD's Obligations
									if(this.DDs_Deliverable_Obligations)
										{
										if(recDeliverable.SPObligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.SPObligations,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.SPObligations != null)
										} //if(this.DDS_Deliverable_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Deliverable_Responsibilities)
										{
										if(recDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.ClientResponsibilities,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Deliverable_Responsibilities)

									// Check if the user specified to include the Deliverable Exclusions
									if(this.Deliverable_Exclusions)
										{
										if(recDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Exclusions,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Exclusions != null)
										} //if(this.Deliverable_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Deliverable_Governance_Controls)
										{
										if(recDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.GovernanceControls,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.GovernanceControls != null)
										} //if(this.Deliverable_GovernanceControls)

									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
									if(recDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										// Check if the user selected Acronyms and Glossy of Terms are requied
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											if(this.Acronyms || this.Glossary_of_Terms)
												{
												foreach(var entry in recDeliverable.GlossaryAndAcronyms)
													{
													TermAndAcronym objTermAndAcronym = new TermAndAcronym();
													objTermAndAcronym.ID = entry.Id;
													objTermAndAcronym.Acronym = entry.Acronym;
													objTermAndAcronym.Term = entry.Title;
													objTermAndAcronym.Meaning = entry.Definition;
													this.TermAndAcronymList.Add(objTermAndAcronym);
													Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Id, entry.Title);
													}
                                                            } // if(this.Acronyms || this.Glossary_of_Terms)
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
                                             } //Try
								catch(DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + deliverableItem.Key
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + deliverableItem.Key + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + deliverableItem.Key 
										+ " contains an error in one of its Enahnce Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}

								catch(Exception exc)
									{
									this.LogError("Content Error in Deliverable " + deliverableItem.Key +
										" Please review all content for this deliverable and correct it.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Content Error in Deliverable " + deliverableItem.Key +
										" Please review all content for this deliverable and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}

								} // if(this.DeliverableHeading
							} // foreach (KeyValuePair<int, String>.....
						} //if(this.Deliverables)

					if(this.Reports)
						{
						objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Reports_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Report_";
						// Insert the individual Reports in the section
						foreach(KeyValuePair<int, string> reportItem in dictReports.OrderBy(key => key.Value))
							{
							if(this.Deliverable_Heading)
								{
								try
									{
									// Obtain the Deliverable info from SharePoint
									var dsDeliverables = datacontexSDDP.Deliverables
										.Expand(p => p.GlossaryAndAcronyms);

									var rsDeliverables =
										from dsDeliverable in dsDeliverables
										where dsDeliverable.Id == reportItem.Key
										select dsDeliverable;

									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);

									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + recDeliverable.Id);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Report Description
									if(this.Report_Description)
										{
										if(recDeliverable.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recDeliverable.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.ISDDescription != null)
										} //if(this.Deliverable_Description)

									// Check if the user specified to include the Report Inputs
									if(this.Report_Inputs)
										{
										if(recDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Inputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Inputs != null)
										} //if(this.Deliverable_Inputs)

									// Check if the user specified to include the Report Outputs
									if(this.Report_Outputs)
										{
										if(recDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Outputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Outputs != null)
										} //if(this.Deliverable_Outputs)

									// Check if the user specified to include the Report DD's Obligations
									if(this.DDs_Report_Obligations)
										{
										if(recDeliverable.SPObligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.SPObligations,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.SPObligations != null)
										} //if(this.DDS_Deliverable_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Report_Responsibilities)
										{
										if(recDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.ClientResponsibilities,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Deliverable_Responsibilities)

									// Check if the user specified to include the Report Exclusions
									if(this.Report_Exclusions)
										{
										if(recDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Exclusions,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Exclusions != null)
										} //if(this.Deliverable_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Deliverable_Governance_Controls)
										{
										if(recDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.GovernanceControls,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.GovernanceControls != null)
										} //if(this.Deliverable_GovernanceControls)

									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
									if(recDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										// Check if the user selected Acronyms and Glossy of Terms are requied
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											if(this.Acronyms || this.Glossary_of_Terms)
												{
												foreach(var entry in recDeliverable.GlossaryAndAcronyms)
													{
													TermAndAcronym objTermAndAcronym = new TermAndAcronym();
													objTermAndAcronym.ID = entry.Id;
													objTermAndAcronym.Acronym = entry.Acronym;
													objTermAndAcronym.Term = entry.Title;
													objTermAndAcronym.Meaning = entry.Definition;
													this.TermAndAcronymList.Add(objTermAndAcronym);
													Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Id, entry.Title);
													}
												} // if(this.Acronyms || this.Glossary_of_Terms)
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
									} //Try
								catch(DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + reportItem.Key
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + reportItem.Key + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + reportItem.Key
										+ " contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}

								catch(Exception exc)
									{
									this.LogError("Content Error in Deliverable " + reportItem.Key +
										" Please review all content for this deliverable and correct it.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Content Error in Deliverable " + reportItem.Key +
										" Please review all content for this deliverable and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}
								} // if(this.DeliverableHeading
							}
						} //if(this.Reports)

					if(this.Meetings)
						{
						objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Meetings_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						string deliverableBookMark = "Meeting_";
						// Insert the individual Meetings in the section
						foreach(KeyValuePair<int, string> meetingItem in dictMeetings.OrderBy(key => key.Value))
							{
							if(this.Meeting_Heading)
								{
								try
									{
									// Obtain the Meeting info from SharePoint
									var dsDeliverables = datacontexSDDP.Deliverables
										.Expand(p => p.GlossaryAndAcronyms);

									var rsDeliverables =
										from dsDeliverable in dsDeliverables
										where dsDeliverable.Id == meetingItem.Key
										select dsDeliverable;

									var recDeliverable = rsDeliverables.FirstOrDefault();
									Console.WriteLine("\t\t + {0} - {1}", recDeliverable.Id, recDeliverable.Title);

									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 3, parBookMark: deliverableBookMark + recDeliverable.Id);
									objRun = oxmlDocument.Construct_RunText(parText2Write: recDeliverable.ISDHeading);
									// Check if a hyperlink must be inserted
									if(documentCollection_HyperlinkURL != "")
										{
										hyperlinkCounter += 1;
										Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
											parMainDocumentPart: ref objMainDocumentPart,
											parImageRelationshipId: hyperlinkImageRelationshipID,
											parClickLinkURL: Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI + recDeliverable.Id,
											parHyperlinkID: hyperlinkCounter);
										objRun.Append(objDrawing);
										}
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);

									// Check if the user specified to include the Meeting Description
									if(this.Meeting_Description)
										{
										if(recDeliverable.ISDDescription != null)
											{
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 3,
												parHTML2Decode: recDeliverable.ISDDescription,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.ISDDescription != null)
										} //if(this.Meeting_Description)

									// Check if the user specified to include the Meeting Inputs
									if(this.Meeting_Inputs)
										{
										if(recDeliverable.Inputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableInputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Inputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Inputs != null)
										} //if(this.Meeting_Inputs)

									// Check if the user specified to include the Meeting Outputs
									if(this.Meeting_Outputs)
										{
										if(recDeliverable.Outputs != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableOutputs_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Outputs,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Outputs != null)
										} //if(this.Meeting_Outputs)

									// Check if the user specified to include the Meeting DD's Obligations
									if(this.DDs_Meeting_Obligations)
										{
										if(recDeliverable.SPObligations != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableDDsObligations_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.SPObligations,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.SPObligations != null)
										} //if(this.DDS_Report_Oblidations)

									// Check if the user specified to include the Client Responsibilities
									if(this.Clients_Meeting_Responsibilities)
										{
										if(recDeliverable.ClientResponsibilities != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableClientResponsibilities_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";
											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.ClientResponsibilities,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Client_Responsibilities != null)
										} //if(this.Clients_Report_Responsibilities)

									// Check if the user specified to include the Deliverable Exclusions
									if(this.Deliverable_Exclusions)
										{
										if(recDeliverable.Exclusions != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableExclusions_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);
											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.Exclusions,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.Exclusions != null)
										} //if(this.Report_Exclusions)

									// Check if the user specified to include the Governance Controls
									if(this.Meeting_Governance_Controls)
										{
										if(recDeliverable.GovernanceControls != null)
											{
											// Insert the Heading
											objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 4);
											objRun = oxmlDocument.Construct_RunText(
												parText2Write: Properties.AppResources.Document_DeliverableGovernanceControls_Heading_Text);
											objParagraph.Append(objRun);
											objBody.Append(objParagraph);

											currentListURI = Properties.AppResources.SharePointURL +
												Properties.AppResources.List_DeliverablesURI +
												currentHyperlinkViewEditURI +
												recDeliverable.Id;
											if(this.ColorCodingLayer1)
												currentContentLayer = "Layer1";
											else
												currentContentLayer = "None";

											// Insert the contents
											objHTMLdecoder.DecodeHTML(
												parMainDocumentPart: ref objMainDocumentPart,
												parDocumentLevel: 4,
												parHTML2Decode: recDeliverable.GovernanceControls,
												parContentLayer: currentContentLayer,
												parHyperlinkImageRelationshipID: hyperlinkImageRelationshipID,
												parHyperlinkURL: currentListURI,
												parTableCaptionCounter: ref tableCaptionCounter,
												parImageCaptionCounter: ref imageCaptionCounter,
												parHyperlinkID: ref hyperlinkCounter,
												parPageHeightTwips: this.PageHight,
												parPageWidthTwips: this.PageWith);
											} // if(recDeliverable.GovernanceControls != null)
										} //if(this.Deliverable_GovernanceControls)

									// Check if there are any Glossary Terms or Acronyms associated with the Deliverable.
									if(recDeliverable.GlossaryAndAcronyms.Count > 0)
										{
										// Check if the user selected Acronyms and Glossy of Terms are requied
										if(this.Acronyms_Glossary_of_Terms_Section)
											{
											if(this.Acronyms || this.Glossary_of_Terms)
												{
												foreach(var entry in recDeliverable.GlossaryAndAcronyms)
													{
													TermAndAcronym objTermAndAcronym = new TermAndAcronym();
													objTermAndAcronym.ID = entry.Id;
													objTermAndAcronym.Acronym = entry.Acronym;
													objTermAndAcronym.Term = entry.Title;
													objTermAndAcronym.Meaning = entry.Definition;
													this.TermAndAcronymList.Add(objTermAndAcronym);
													Console.WriteLine("\t\t\t + Term & Acronym added: {0} - {1}", entry.Id, entry.Title);
													}
												} // if(this.Acronyms || this.Glossary_of_Terms)
											} // if(this.Acronyms_Glossary_of_Terms_Section)
										} //if(recDeliverable.GlossaryAndAcronyms.Count > 0)
									} //Try
								catch(DataServiceClientException)
									{
									// If the entry is not found - write an error in the document and record an error in the error log.
									this.LogError("Error: The Deliverable ID " + meetingItem.Key
										+ " doesn't exist in SharePoint and couldn't be retrieved.");
									objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Error: Deliverable " + meetingItem.Key + " is missing.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}
								catch(InvalidTableFormatException exc)
									{
									Console.WriteLine("Exception occurred: {0}", exc.Message);
									// A Table content error occurred, record it in the error log.
									this.LogError("Error: The Deliverable ID: " + meetingItem.Key
										+ " contains an error in one of its Enhance Rich Text columns. Please review the content (especially tables).");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 2);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "A content error occurred at this position and valid content could " +
										"not be interpreted and inserted here. Please review the content in the SharePoint system and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									}

								catch(Exception exc)
									{
									this.LogError("Content Error in Deliverable " + meetingItem.Key +
										" Please review all content for this deliverable and correct it.");
									objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 5);
									objRun = oxmlDocument.Construct_RunText(
										parText2Write: "Content Error in Deliverable " + meetingItem.Key +
										" Please review all content for this deliverable and correct it.",
										parIsNewSection: false,
										parIsError: true);
									objParagraph.Append(objRun);
									objBody.Append(objParagraph);
									Console.WriteLine("Exception occurred: {0} - {1}", exc.HResult, exc.Message);
									}

								} // if(this.MeetingHeading
							}	// foreach.....
						} //if(this.Meetings)
					} //if(this.DRM_Section)

				//-------------------------------------------------------
				// Insert the Service Levels Section
				if(this.Service_Level_Section)
					{
					// Insert the Service Levels Section
					if(this.Service_Level_Section)
						{
						objParagraph = oxmlDocument.Insert_Section();
						objRun = oxmlDocument.Construct_RunText(
							parText2Write: Properties.AppResources.Document_DRM_Section_Text,
							parIsNewSection: true);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);

						objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
						objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_ServiceLevels_Heading_Text);
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						}

					if(this.Service_Level_Heading)
						{


						if(this.Service_Level_Commitments_Table)
							{

							} //if(this.Service_Level_Commitments_Table)
						} //if(this.Service_Level_Heading)
					} //if(this.Service_Level_Section)

Glossary_and_Acronyms:
				//--------------------------------------------------
				// Insert the Glossary of Terms and Acronym Section
				if(this.Acronyms_Glossary_of_Terms_Section)
					{
					objParagraph = oxmlDocument.Insert_Section();
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_GlossaryAndAcronymSection_HeadingText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					}
				//-------------------------------------------------
				// Insert the Acronyms
				if(this.Acronyms)
					{
					objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_Acronyms_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL, 
							parHyperlinkID: hyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					Console.WriteLine("The Acronyms List before sort...");
					foreach(TermAndAcronym item in this.TermAndAcronymList)
						{
						Console.WriteLine("\t\t + {0}", item.ID);
						};

					if(this.TermAndAcronymList.Count < 1)
						{
						objParagraph = oxmlDocument.Construct_Paragraph(1);
						objRun = oxmlDocument.Construct_RunText("No acronyms were defined.");
						objParagraph.Append(objRun);
						objBody.Append(objParagraph);
						goto Glossary_of_Terms;
						}
					List<TermAndAcronym> listTermAndAcronyms = this.TermAndAcronymList;
					// Populate the Accronyms and Terms
					string result = TermAndAcronym.PopulateTerms(ref listTermAndAcronyms);
					if(result.Contains("Error"))
						{
						objParagraph = oxmlDocument.Construct_Error(result);
						objBody.Append(objParagraph);
						goto Glossary_of_Terms;
						}
					
					this.TermAndAcronymList = listTermAndAcronyms;
					if(this.TermAndAcronymList.Count > 0)
						{
						// Sort the list by Acronym
						this.TermAndAcronymList.Sort(delegate (TermAndAcronym x, TermAndAcronym y)
							{
								if(x.Acronym == null && y.Acronym == null)
									return 0;
								else if(x.Acronym == null)
									return -1;
								else if(y.Acronym == null)
									return 1;
								else
									return x.Acronym.CompareTo(y.Acronym);
								});
						Console.WriteLine("After Sort by Acronym...");
						foreach(TermAndAcronym item in this.TermAndAcronymList)
							{
							Console.WriteLine("\t\t + {0} - {1} - {2}", item.ID, item.Term, item.Acronym);
							}

						// Construct a Table object instance
						Table objTable = new Table();
						objTable = oxmlDocument.ConstructTable(
							parPageWidth: this.PageWith,
							parFirstRow: true, 
							parNoVerticalBand: true, 
							parNoHorizontalBand: false);
						TableRow objTableRow = new TableRow();
						TableCell objTableCell = new TableCell();
						string tableText = "";
						TableGrid objTableGrid = new TableGrid();
						List<UInt32> lstTableColumns = new List<UInt32>();
						lstTableColumns.Add(this.PageWith * 20 / 100);
						lstTableColumns.Add(this.PageWith * 80 / 100);
						//columnWidth = lstTableColumns [0].ToString();
						objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
						// Append the TableGrid object instance to the Table object instance
						objTable.Append(objTableGrid);
						// Create a TableRow object instance
						objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: true);
						objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
						objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[0] , parIsFirstRow: true);
						// Create a Pargaraph for the text to go into the TableCell
						objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
						tableText = "Acronym";
						objRun = oxmlDocument.Construct_RunText(tableText);
						objParagraph.Append(objRun);
						objTableCell.Append(objParagraph);
						objTableRow.Append(objTableCell);
						objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
						objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[1] , parIsFirstRow: true);
						// Create another Pargaraph for the text to go into the TableCell
						objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
						tableText = "Term";
						objRun = oxmlDocument.Construct_RunText(tableText);
						objParagraph.Append(objRun);
						objTableCell.Append(objParagraph);
						objTableRow.Append(objTableCell);
						objTable.Append(objTableRow);
						Console.WriteLine("\t Generate Table with Acronyms");
						foreach(TermAndAcronym item in this.TermAndAcronymList)
							{
							if(item.Acronym != null)
								{
								Console.WriteLine("\t\t + {0} - ({1}) - {2}", item.Term, item.ID, item.Acronym);
								// Create a TableRow object instance
								objTableRow = oxmlDocument.ConstructTableRow();
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
								objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[0]);
								// Create a Pargaraph for the text to go into the TableCell
								objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
								objRun = oxmlDocument.Construct_RunText(item.Acronym);
								objParagraph.Append(objRun);
								objTableCell.Append(objParagraph);
								objTableRow.Append(objTableCell);
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
								objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[1]);
								// Create another Pargaraph for the text to go into the TableCell
								objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
								objRun = oxmlDocument.Construct_RunText(item.Term);
								objParagraph.Append(objRun);
								objTableCell.Append(objParagraph);
								objTableRow.Append(objTableCell);
								objTable.Append(objTableRow);
								}
							}    // end of ForEach Loop
						objBody.Append(objTable);
						}     //if(this.TermAndAcronymList.Count > 0)
					} // if (this.Acronyms)

Glossary_of_Terms:	//----------------------------------------------------
				// If the user selected to have a Glossary of Terms
				if(this.Glossary_of_Terms)
					{
					objParagraph = oxmlDocument.Insert_Heading(parHeadingLevel: 1);
					objRun = oxmlDocument.Construct_RunText(parText2Write: Properties.AppResources.Document_GlossaryOfTerms_HeadingText);
					// Check if a hyperlink must be inserted
					if(documentCollection_HyperlinkURL != "")
						{
						hyperlinkCounter += 1;
						DocumentFormat.OpenXml.Wordprocessing.Drawing objDrawing = oxmlDocument.ConstructClickLinkHyperlink(
							parMainDocumentPart: ref objMainDocumentPart,
							parImageRelationshipId: hyperlinkImageRelationshipID,
							parClickLinkURL: documentCollection_HyperlinkURL,
							parHyperlinkID: hyperlinkCounter);
						objRun.Append(objDrawing);
						}
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);

					if(this.TermAndAcronymList.Count > 0)
						{
						List<TermAndAcronym> listTermAndAcronyms = this.TermAndAcronymList;
						string result = TermAndAcronym.PopulateTerms(ref listTermAndAcronyms);
						if(result.Contains("Error"))
							{
							objParagraph = oxmlDocument.Construct_Error(result);
							objBody.Append(objParagraph);
							goto Document_Acceptance_Section;
							}
						else
							{
							this.TermAndAcronymList = listTermAndAcronyms;
						// Sort the list by Term
						// https://msdn.microsoft.com/en-US/library/w56d4y5z(v=vs.110).aspx
						this.TermAndAcronymList.Sort(delegate (TermAndAcronym x, TermAndAcronym y)
								{
									if(x.Term == null && y.Term == null)
										return 0;
									else if(x.Term == null)
										return -1;
									else if(y.Term == null)
										return 1;
									else
										return x.Term.CompareTo(y.Term);
									});
							Console.WriteLine("After Sorting Temrs...");
							foreach(TermAndAcronym item in this.TermAndAcronymList)
							{
							Console.WriteLine("\t\t + {0} - {1} - {2}", item.ID, item.Term, item.Acronym);
							}

						// Construct a Table object instance
						Table objTable = new Table();
						objTable = oxmlDocument.ConstructTable(
							parPageWidth: this.PageWith, 
							parFirstRow: true, 
							parNoVerticalBand: true, 
							parNoHorizontalBand: false);
						TableRow objTableRow = new TableRow();
						TableCell objTableCell = new TableCell();
						TableGrid objTableGrid = new TableGrid();
						List<UInt32> lstTableColumns = new List<UInt32>();
						lstTableColumns.Add(this.PageWith * 30 / 100);
						lstTableColumns.Add(this.PageWith * 70 / 100);
						objTableGrid = oxmlDocument.ConstructTableGrid(lstTableColumns);
						// Append the TableGrid object instance to the Table object instance
						objTable.Append(objTableGrid);
						// Create a TableRow object instance
						objTableRow = oxmlDocument.ConstructTableRow(parIsFirstRow: true);
						objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
						objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[0], parIsFirstRow: true);
						// Create a Pargaraph for the text to go into the TableCell
						objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
						objRun = oxmlDocument.Construct_RunText("Term");
						objParagraph.Append(objRun);
						objTableCell.Append(objParagraph);
						objTableRow.Append(objTableCell);
						objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
						objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[1], parIsFirstRow: true);
						// Create another Pargaraph for the text to go into the TableCell
						objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
						objRun = oxmlDocument.Construct_RunText("Meaning");
						objParagraph.Append(objRun);
						objTableCell.Append(objParagraph);
						objTableRow.Append(objTableCell);
						objTable.Append(objTableRow);
						Console.WriteLine("\t Generate Table with Terms");
						foreach(TermAndAcronym item in this.TermAndAcronymList)
							{
							if(item.Term != null)
								{
								Console.WriteLine("\t\t + {0} - ({1})", item.Term, item.ID);
								// Create a TableRow object instance
								objTableRow = oxmlDocument.ConstructTableRow();
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
								objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[0]);
								// Create a Pargaraph for the text to go into the TableCell
								objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
								objRun = oxmlDocument.Construct_RunText(item.Term);
								objParagraph.Append(objRun);
								objTableCell.Append(objParagraph);
								objTableRow.Append(objTableCell);
								objParagraph = oxmlDocument.Construct_Paragraph(parBodyTextLevel: 1, parIsTableParagraph: true);
								objTableCell = oxmlDocument.ConstructTableCell(lstTableColumns[1]);
								// Create another Pargaraph for the text to go into the TableCell
								objParagraph = oxmlDocument.Construct_Paragraph(1, parIsTableParagraph: true);
								objRun = oxmlDocument.Construct_RunText(item.Meaning);
								objParagraph.Append(objRun);
								objTableCell.Append(objParagraph);
								objTableRow.Append(objTableCell);
								objTable.Append(objTableRow);
								}
							}    // end of ForEach Loop
						objBody.Append(objTable);
						}     // No errors
					}    // this.TermAndAcronymList.Count > 0)
				}	// if(this.Glossary_of_Terms)

Document_Acceptance_Section:
				// Generate the Document Acceptance Section if it was selected
				if(this.Document_Acceptance_Section)
					{
					objParagraph = oxmlDocument.Insert_Section();
					objRun = oxmlDocument.Construct_RunText(
						parText2Write: Properties.AppResources.Document_AcceptanceText,
						parIsNewSection: true);
					objParagraph.Append(objRun);
					objBody.Append(objParagraph);
					if(this.DocumentAcceptanceRichText != null)
						{
						objHTMLdecoder.DecodeHTML(
							parMainDocumentPart: ref objMainDocumentPart,
							parDocumentLevel: 1,
							parPageWidthTwips: this.PageWith,
							parPageHeightTwips: this.PageHight,
							parHTML2Decode: this.DocumentAcceptanceRichText,
							parTableCaptionCounter: ref tableCaptionCounter,
							parImageCaptionCounter: ref imageCaptionCounter,
							parHyperlinkID: ref hyperlinkCounter);
						}
					}

				Console.WriteLine("\t\t Document generated, now saving and closing the document.");
				// Save and close the Document
				objWPdocument.Close();

				//TODO: add code to validate the created xml document. 
				// https://msdn.microsoft.com/en-us/library/bb497334%28v=office.12%29.aspx

				Console.WriteLine(
					"Generation started...: {0} \nGeneration completed: {1} \n Durarion..........: {2}",
					timeStarted,
					DateTime.Now,
					(DateTime.Now - timeStarted));
				} // end Try

			catch(OpenXmlPackageException exc)
				{
				//TODO: add code to catch exception.
				}
			catch(ArgumentNullException exc)
				{
				//TODO: add code to catch exception.
				}

			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}

		} // end of Services_Framework_Document_DRM_Sections class

	/// <summary>
	/// This class represent the Statement of Work (SoW) with sperate DRM (Deliverable Report Meeting) sections
	/// It inherits from the DRM Sections Class.
	/// </summary>
	class Contract_SoW_Service_Description : External_DRM_Sections
		{
		/// <summary>
		/// this option takes the values passed into the method as a list of integers
		/// which represents the options the user selected and transposing the values by
		/// setting the properties of the object.
		/// </summary>
		/// <param name="parOptions">The input must represent a List<int> object.</int></param>
		/// <returns></returns>
		public void TransposeDocumentOptions(ref List<int> parOptions)
			{
			int errors = 0;
			if(parOptions != null)
				{
				if(parOptions.Count > 0)
					{
					foreach(int option in parOptions)
						{
						switch(option)
							{
							case 195:
								this.Introductory_Section = true;
								break;
							case 196:
								this.Introduction = true;
								break;
							case 197:
								this.Service_Portfolio_Section = true;
								break;
							case 198:
								this.Service_Portfolio_Description = true;
								break;
							case 199:
								this.Service_Family_Heading = true;
								break;
							case 200:
								this.Service_Family_Description = true;
								break;
							case 201:
								this.Service_Product_Heading = true;
								break;
							case 202:
								this.Service_Product_Description = true;
								break;
							case 203:
								this.Service_Feature_Heading = true;
								break;
							case 204:
								this.Service_Feature_Description = true;
								break;
							case 205:
								this.Deliverables_Reports_Meetings = true;
								break;
							case 206:
								this.DRM_Heading = true;
								break;
							case 207:
								this.DRM_Summary = true;
								break;
							case 208:
								this.Service_Levels = true;
								break;
							case 209:
								this.Service_Level_Heading = true;
								break;
							case 210:
								this.Service_Level_Commitments_Table = true;
								break;
							case 211:
								this.DRM_Section = true;
								break;
							case 212:
								this.Deliverables = true;
								break;
							case 213:
								this.Deliverable_Heading = true;
								break;
							case 214:
								this.Deliverable_Description = true;
								break;
							case 215:
								this.DDs_Deliverable_Obligations = true;
								break;
							case 216:
								this.Clients_Deliverable_Responsibilities = true;
								break;
							case 217:
								this.Deliverable_Exclusions = true;
								break;
							case 218:
								this.Deliverable_Governance_Controls = true;
								break;
							case 219:
								this.Reports = true;
								break;
							case 220:
								this.Report_Heading = true;
								break;
							case 221:
								this.Report_Description = true;
								break;
							case 222:
								this.DDs_Report_Obligations = true;
								break;
							case 223:
								this.Clients_Report_Responsibilities = true;
								break;
							case 224:
								this.Report_Exclusions = true;
								break;
							case 225:
								this.Report_Governance_Controls = true;
								break;
							case 226:
								this.Meetings = true;
								break;
							case 227:
								this.Meeting_Heading = true;
								break;
							case 228:
								this.Meeting_Description = true;
								break;
							case 229:
								this.DDs_Meeting_Obligations = true;
								break;
							case 230:
								this.Clients_Meeting_Responsibilities = true;
								break;
							case 231:
								this.Meeting_Exclusions = true;
								break;
							case 232:
								this.Meeting_Governance_Controls = true;
								break;
							case 233:
								this.Acronyms_Glossary_of_Terms_Section = true;
								break;
							case 234:
								this.Acronyms = true;
								break;
							case 235:
								this.Glossary_of_Terms = true;
								break;
							default:
								// just ignore
								break;
							}
						} // foreach(int option in parOptions)
					}
				else
					{
					this.LogError("There are no selected options - (Application Error)");
					errors += 1;
					}
				}
			else
				{
				this.LogError("The selected options are null - (Application Error)");
				errors += 1;
				}
			}

		public bool Generate()
			{
			Console.WriteLine("\t\t Begin to generate {0}", this.DocumentType);
			//TODO: Code to be added for Contract_SoW_Service_Description's Generate method.
			Console.WriteLine("\t\t Complete the generation of {0}", this.DocumentType);
			return true;
			}
		} // end of SowD_Document_DRM_Sections class

	class TermAndAcronym 
		{
		private string _term;
		public string Term
			{
			get{return this._term;}
			set{this._term = value;}
			}
		private string _meaning;
		public string Meaning
			{
			get{return this._meaning;}
			set{this._meaning = value;}
			}
		private string _acronym;
		public string Acronym
			{
			get{return this._acronym;}
			set{this._acronym = value;}
			}
		private int _id;
		public int ID
			{
			get{return this._id;}
			set{this._id = value;}
			}


		public static string PopulateTerms(ref List<TermAndAcronym> parTermsAndAcronyms)
			{
			// Initially all the terms will be added by inserting only the ID of the entry that resides in the
			// Glossary and Acronyms List in SharePoint, then at a later stage this method is used to poulate the Term, Acronym and the Meanings.

			DesignAndDeliveryPortfolioDataContext datacontexSDDP = new DesignAndDeliveryPortfolioDataContext(new
				Uri(DocGenerator.Properties.AppResources.SharePointSiteURL + Properties.AppResources.SharePointRESTuri)); // "/_vti_bin/listdata.svc"));
			datacontexSDDP.Credentials = CredentialCache.DefaultCredentials;
			datacontexSDDP.MergeOption = System.Data.Services.Client.MergeOption.NoTracking;
			try
				{
				Console.WriteLine("\t There are {0} Terms and Acronyms to process", parTermsAndAcronyms.Count);
				if(parTermsAndAcronyms.Count > 0)
					{
					var GlossaryAndAcronymsList = datacontexSDDP.GlossaryAndAcronyms;
					//var GlossaryAndAcronyms = from GaAL in GlossaryAndAcronymsList select GaAL;

					foreach(TermAndAcronym item in parTermsAndAcronyms)
						{
						Console.WriteLine("\t {0} was read...", item.ID);
						if(item.Term == null)
							{
							var foundEntries =
								from entry in GlossaryAndAcronymsList
								where entry.Id == item.ID
								select entry;
							foreach(var GAentry in foundEntries)
								{
								//Console.WriteLine("\t\t {0} was not found...", item.ID);
								Console.WriteLine("\t\t Found entry {0} - ({1}) = {2}", GAentry.Id, GAentry.Title, GAentry.Acronym);
								item.Term = GAentry.Title;
								item.Acronym = GAentry.Acronym;
								item.Meaning = GAentry.Definition;
								break;
								}
							}
						}
					}
				return "Success";
				}
			catch(Exception ex)
				{
				Console.WriteLine("Exception: [{0}] occurred and was caught. \n{1}", ex.HResult.ToString(), ex.Message);

				if(ex.HResult == -2146330330)
					return "Error: Cannot access site: " + Properties.AppResources.SharePointSiteURL + " Ensure the computer is connected to the Dimension Data Domain network";
				else if(ex.HResult == -2146233033)
					return "Error: Input string missing to connect to " + Properties.AppResources.SharePointSiteURL + " Ensure the computer is connected to the Dimension Data Domain network";
				else
					return "Error: Unexpected error occurred. " + ex.HResult.ToString() + " - " + ex.Message;
				}
			}    // end of class
		}
	}