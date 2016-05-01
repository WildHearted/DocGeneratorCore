using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{

	public class GlossaryAcronym
		{
		public string Term{get; set;}
		public string Meaning{get; set;}
		public string Acronym{get; set;}
		public int ID{get; set;}
		} // end of Class GlossaryAndAcronym

	public class ServicePortfolio
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string PortfolioType{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string SOWheading{get; set;}
		public string SOWdescription{get; set;}
		public List<ServiceFamily> listServiceFamilies{get; set;}
		}

	public class ServiceFamily
		{
		public int ID{get; set;}
		public int? ServicePortfolioID{get; set;}
		public string Title{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string SOWheading{get; set;}
		public string SOWdescription{get; set;}

		} // end of class ServicePFamily

	///##################################################
	/// <summary>
	/// Service Product object represent an entry in the Service Products SharePoint List
	/// </summary>
	public class ServiceProduct
		{
		public int ID{get; set;}public int? ServiceFamilyID{get; set;}
		public string Title{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string KeyDDbenefits{get; set;}
		public string KeyClientBenefits{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string SOWheading{get; set;}
		public string SOWdescription{get; set;}
		public double? PlannedElements{get; set;}
		public double? PlannedFeatures{get; set;}
		public double? PlannedDeliverables{get; set;}
		public double? PlannedServiceLevels{get; set;}
		public double? PlannedMeetings{get; set;}
		public double? PlannedReports{get; set;}
		public double? PlannedActivities{get; set;}
		public double? PlannedActivityEffortDrivers{get; set;}		
		} // end of class ServiceProduct

	///############################################
	/// <summary>
	/// This object represents an entry in the Service Elements SharePoint List
	/// </summary>
	public class ServiceElement
		{
		public int ID{get; set;}
		public int? ServiceProductID{get; set;}
		public string Title{get; set;}
		public double? SortOrder{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string Objectives{get; set;}
		public string KeyClientAdvantages{get; set;}
		public string KeyClientBenefits{get; set;}
		public string KeyDDbenefits{get; set;}
		public string KeyPerformanceIndicators{get; set;}
		public string CriticalSuccessFactors{get; set;}
		public string ProcessLink{get; set;}
		public string ContentLayerValue{get; set;}
		public int? ContentPredecessorElementID{get; set;}
		public string ContentStatus{get; set;}

		} // end Class ServiceElement

	///##############################################################
	///#### Service Feature Object
	///##############################################################
	/// <summary>
	/// This object represents an entry in the Service Features SharePoint List.
	/// </summary>
	public class ServiceFeature
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public int? ServiceProductID{get; set;}
		public double? SortOrder{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string SOWheading{get; set;}
		public string SOWdescription{get; set;}
		public string ContentLayerValue{get; set;}
		public int? ContentPredecessorFeatureID{get; set;}
		public string ContentStatus{get; set;}

		} // end Class ServiceFeature

	/// #############################################
	/// ### Deliverables Object
	/// #############################################
	/// <summary>
	/// This object represent an entry in the Deliverables SharePoint List.
	/// </summary>
	public class Deliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string ISDsummary{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string CSDsummary{get; set;}
		public string SoWheading{get; set;}
		public string SoWdescription{get; set;}
		public string SoWsummary{get; set;}
		public string DeliverableType{get; set;}
		public string Inputs{get; set;}
		public string Outputs{get; set;}
		public string DDobligations{get; set;}
		public string ClientResponsibilities{get; set;}
		public string Exclusions{get; set;}
		public string GovernanceControls{get; set;}
		public double? SortOrder{get; set;}
		public string TransitionDescription{get; set;}
		public List<string> SupportingSystems{get; set;}
		public string WhatHasChanged{get; set;}
		public string ContentLayerValue{get; set;}
		public string ContentStatus{get; set;}
		public Dictionary<int, string> GlossaryAndAcronyms{get; set;}
		public int? ContentPredecessorDeliverableID{get; set;}
		public List<int?> RACIaccountables{get; set;}
		public List<int?> RACIresponsibles{get; set;}
		public List<int?> RACIinformeds{get; set;}
		public List<int?> RACIconsulteds{get; set;}
		
		} // end Class Deliverables

	// ####################################################
	// ### Deliverable Service Levels class
	// ####################################################
	/// <summary>
	/// 
	/// </summary>
	public class DeliverableServiceLevel
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string ContentStatus{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public ServiceLevel AssociatedServiceLevel{get; set;}
		public int? AssociatedServiceLevelID{get; set;}
		public ServiceProduct AssociatedServiceProduct{get; set;}
		public int? AssociatedServiceProductID{get; set;}
		public string AdditionalConditions{get; set;}

		}// end of DeliverableServiceLevels class

	// ####################################################
	// ### Deliverable Activities class
	// ####################################################
	/// <summary>
	/// 
	/// </summary>
	public class DeliverableActivity{public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public Activity AssociatedActivity{get; set;}
		public int? AssociatedActivityID{get; set;}
		
		}// end of DeliverableActivities class

	//##########################################################
	/// <summary>
	/// This object represents an entry in the DeliverableTechnologies SharePoint List
	/// Each entry in the list is a DeliverableTechnology object.
	/// </summary>
	public class DeliverableTechnology
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Considerations{get; set;}
		public TechnologyProduct TechnologyProduct{get; set;}
		public Deliverable Deliviverable{get; set;}
		public string RoadmapStatus{get; set;}

		} // end of TechnologyProduct class

	//##########################################################
	//### FeatureDeliverable class
	//#########################################################
	/// <summary>
	/// The FeatureDeliverable object is the junction table or the cross-reference table between Service Features and Deliverables.
	/// </summary>
	public class FeatureDeliverable{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public ServiceFeature AssociatedFeature{get; set;}
		public int? AssociatedFeatureID{get; set;}
		
		} // end of FeatureDeliverable class

	//##########################################################
	//### ElementDeliverable class
	//#########################################################
	/// <summary>
	/// The ElementDeliverable objects is the junction table or the cross-reference table between Service Elements and Deliverables.
	/// </summary>
	public class ElementDeliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public ServiceElement AssociatedElement{get; set;}
		public int? AssociatedElementID{get; set;}
		
		} // end of ElementDeliverable class

	// ###################################
	// ### Mapping Object
	// ###################################

	/// <summary>
	/// The Mapping object represents an entry in the Mappings List in SharePoint.
	/// </summary>
	public class Mapping
		{
		public int? ID{get; set;}
		public string Title{get; set;}
		public string ClientName{get; set;}
		
		} // end Class Mapping

	//###############################################
	/// <summary>
	/// The MappingServiceTower object represents an entry in the Mapping Service Towers List in SharePoint.
	/// </summary>
	public class MappingServiceTower
		{
		public int ID{get; set;}
		public string Title{get; set;}
		
		} // end Class Mapping Service Towers

	//##########################################
	/// <summary>
	/// The MappingRequirement object represents an entry in the MappingRequirements List.
	/// </summary>
	public class MappingRequirement
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public int? MappingServiceTowerID{get; set;}
		public double? SortOrder{get; set;}
		public string RequirementText{get; set;}
		public string RequirementServiceLevel{get; set;}
		public string SourceReference{get; set;}
		public string ComplianceStatus{get; set;}
		public string ComplianceComments{get; set;}
		
		} // end Class Mapping Requirements

	//############################################
	/// <summary>
	/// The Mapping Deliverable is the class used to for the Mapping Deliverables SharePoint List.
	/// </summary>
	//############################################
	public class MappingDeliverable
		{
		public int ID { get; set; }
		public int? MappingRequirementID{get; set;}
		public string Title { get; set; }
		/// <summary>
		/// Represents the translated value in the Deliverable Choice column of the MappingDeliverable List. TRUE if "New" else FALSE
		/// </summary>
		public bool NewDeliverable { get; set; }
		public string ComplianceComments { get; set; }
		public String NewRequirement { get; set; }
		public int? MappedDeliverableID { get; set; }
		}

	//#############################################
	/// <summary>
	/// The MappingAssumption represents an entry of the Mapping Assumptions List in SharePoint
	/// </summary>
	public class MappingAssumption
		{
		public int ID{get; set;}
		public int? MappingRequirementID{get; set;}
		public string Title{get; set;}
		public string Description{get; set;}
		
		}
	//##################################################
	/// <summary>
	/// Mapping Risk Object
	/// </summary>
	public class MappingRisk
		{
		public int ID{get; set;}
		public int? MappingRequirementID{get; set;}
		public string Title{get; set;}
		public string Statement{get; set;}
		public string Mitigation{get; set;}
		public double? ExposureValue{get; set;}
		public string Status{get; set;}
		public string Exposure{get; set;}
		} // End of Class MappingRisk
	
	/// <summary>
	/// The Mapping Service Level is the class used to for the Mapping Service Levels SharePoint List.
	/// </summary>
	public class MappingServiceLevel
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string RequirementText{get; set;}
		public bool? NewServiceLevel{get; set;}
		public string ServiceLevelText{get; set;}
		public int? MappedServiceLevelID{get; set;}
		public int? MappedDeliverableID{get; set;}
		}

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Service Levels SharePoint List
	/// </summary>
	public class ServiceLevel
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string SOWheading{get; set;}
		public string SOWdescription{get; set;}
		public string ContentStatus{get; set;}
		public string Measurement{get; set;}
		public string MeasurementInterval{get; set;}
		public string ReportingInterval{get; set;}
		public string CalcualtionMethod{get; set;}
		public string CalculationFormula{get; set;}
		public string ServiceHours{get; set;}
		public List<ServiceLevelTarget> PerfomanceThresholds{get; set;}
		public List<ServiceLevelTarget> PerformanceTargets{get; set;}
		public string BasicConditions{get; set;}
		public string AdditionalConditions{get; set;}
		
		} // end of Service Levels class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Activities SharePoint List
	/// </summary>
	public class ServiceLevelTarget
		{
		public int ID{get; set;}
		public string Type{get; set;}
		public string Title{get; set;}
		public string ContentStatus{get; set;}
		}
	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Activities SharePoint List
	/// </summary>
	public class Activity
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public double? SortOrder{get; set;}
		public string Optionality{get; set;}
		public string ISDheading{get; set;}
		public string ISDdescription{get; set;}
		public string CSDheading{get; set;}
		public string CSDdescription{get; set;}
		public string SOWheading{get; set;}
		public string SOWdescription{get; set;}
		public string ContentStatus{get; set;}
		public string Input{get; set;}
		public string Output{get; set;}
		public string Catagory{get; set;}
		public string Assumptions{get; set;}
		public string OLAvariations{get; set;}
		public string OLA{get; set;}
		public List<JobRole> RACI_Responsible{get; set;}
		public List<JobRole> RACI_Accountable{get; set;}
		public List<JobRole> RACI_Consulted{get; set;}
		public List<JobRole> RACI_Informed{get; set;}
		
		} // end of Activitiy class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Job Framewotk Alignment SharePoint List
	/// But each entry is essentially a JobRole, therefore the class is named JobRole
	/// </summary>
	public class JobRole
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string DeliveryDomain{get; set;}
		public string SpecificRegion{get; set;}
		public string RelevantBusinessUnit{get; set;}
		public string OtherJobTitles{get; set;}
		public string JobFrameworkLink{get; set;}
		
		} // end of JobRole class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Technology Categories SharePoint List
	/// Each entry in the list is a Technology Category object.
	/// </summary>
	public class TechnologyCategory
		{
		public int ID{get; set;}
		public string Title{get; set;}
		
		} // end of TechnologyCategory class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Technology Vendors SharePoint List
	/// Each entry in the list is a Technology Vendor object.
	/// </summary>
	public class TechnologyVendor
		{
		public int ID{get; set;}
		public string Title{get; set;}
		
		} // end of TechnologyVendor class

	//##########################################################
	/// <summary>
	/// This object represents an entry in the Technology Products SharePoint List
	/// Each entry in the list is a Technology Product object.
	/// </summary>
	public class TechnologyProduct
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Prerequisites{get; set;}
		public TechnologyCategory Category{get; set;}
		public TechnologyVendor Vendor{get; set;}

		
		} // end of TechnologyProduct class

	public class CompleteDataSet
		{
		public Dictionary<int, JobRole> dsJobroles{get; set;}
		public Dictionary<int, GlossaryAcronym> dsGlossaryAcronyms{get; set;}
		public Dictionary<int, ServicePortfolio> dsPortfolios{get; set;}
		public Dictionary<int, ServiceFamily> dsFamilies{get; set;}
		public Dictionary<int, ServiceProduct> dsProducts{get; set;}
		public Dictionary<int, ServiceElement> dsElements{get; set;}
		public Dictionary<int, ServiceFeature> dsFeatures{get; set;}
		public Dictionary<int, Deliverable> dsDeliverables{get; set;}
		public Dictionary<int, ElementDeliverable> dsElementDeliverables{get; set;}
		public Dictionary<int, FeatureDeliverable> dsFeatureDeliverables{get; set;}
		public Dictionary<int, Activity> dsActivities{get; set;}
		public Dictionary<int, DeliverableActivity> dsDeliverableActivities{get; set;}
		public Dictionary<int, TechnologyProduct> dsTechnologyProducts{get; set;}
		public Dictionary<int, DeliverableTechnology> dsDeliverableTechnologies{get; set;}
		public Dictionary<int, ServiceLevel> dsServiceLevels{get; set;}
		public Dictionary<int, DeliverableServiceLevel> dsDeliverableServiceLevels{get; set;}
		public Dictionary<int?, Mapping> dsMappings{get; set;}
		public Dictionary<int, MappingServiceTower> dsMappingServiceTowers{get; set;}
		public Dictionary<int, MappingRequirement> dsMappingRequirements{get; set;}
		public Dictionary<int, MappingAssumption> dsMappingAssumptions{get; set;}
		public Dictionary<int, MappingDeliverable> dsMappingDeliverables{get; set;}
		public Dictionary<int, MappingRisk> dsMappingRisks{get; set;}
		public Dictionary<int, MappingServiceLevel> dsMappingServiceLevels{get; set;}
		public bool PopulateBaseObjects(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP)
			{
			int intLastReadID = 0;
			bool boolFetchMore = false;
			DateTime startTime = DateTime.Now;
			DateTime setStart = DateTime.Now;
			// Please Note: 
			// SharePoint's REST API has a limit which returns only 1000 entries at a time
			// therefore a paging principle is implemented to return all the entries in the List.

			try
				{
				Console.Write("\nPopulating the complete DataSet...");

				// -------------------------
				// Populate GlossaryAcronyms
				Console.Write("\n\t + Glossary & Acronyms...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsGlossaryAcronyms = new Dictionary<int, GlossaryAcronym>();
				do
					{
					var rsGlossaryAcronyms =
						from dsGlossaryAcronym in parDatacontexSDDP.GlossaryAndAcronyms
						where dsGlossaryAcronym.Id > intLastReadID
						select dsGlossaryAcronym;

					boolFetchMore = false;

					foreach(GlossaryAndAcronymsItem record in rsGlossaryAcronyms)
						{
						GlossaryAcronym objGlossaryAcronym = new GlossaryAcronym();
						intLastReadID = record.Id;
						boolFetchMore = true;
						objGlossaryAcronym.ID = record.Id;
						objGlossaryAcronym.Term = record.Title;
						objGlossaryAcronym.Acronym = record.Acronym;
						objGlossaryAcronym.Meaning = record.Definition;
						this.dsGlossaryAcronyms.Add(key: record.Id, value: objGlossaryAcronym);
						}

					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsGlossaryAcronyms.Count, DateTime.Now - setStart);

				// Populate JobRoles
				Console.Write("\n\t + JobRoles...\t\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsJobroles = new Dictionary<int, JobRole>();
				var dsJobFrameworks = parDatacontexSDDP.JobFrameworkAlignment
					.Expand(jf => jf.JobDeliveryDomain);
				do
					{
					var rsJobFrameworks =
						from dsJobFramework in dsJobFrameworks
						where dsJobFramework.Id > intLastReadID
						select dsJobFramework;

					boolFetchMore = false;
					foreach(JobFrameworkAlignmentItem record in rsJobFrameworks)
						{
						JobRole objJobRole = new JobRole();
						objJobRole.ID = record.Id;
						intLastReadID = record.Id;
						boolFetchMore = true;
						objJobRole.Title = record.Title;
						objJobRole.OtherJobTitles = record.RelatedRoleTitle;
						if(record.JobDeliveryDomain.Title != null)
							objJobRole.DeliveryDomain = record.JobDeliveryDomain.Title;
						if(record.RelevantBusinessUnitValue != null)
							objJobRole.RelevantBusinessUnit = record.RelevantBusinessUnitValue;
						if(record.SpecificRegionValue != null)
							objJobRole.SpecificRegion = record.SpecificRegionValue;
						this.dsJobroles.Add(key: record.Id, value: objJobRole);
						}
					} while(boolFetchMore);
				Console.Write("\t\t\t {0} \t {1}", this.dsJobroles.Count, DateTime.Now - setStart);

				// -------------------------
				// Populate TechnologyProdcuts
				Console.Write("\n\t + TechnologyProducts...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsTechnologyProducts = new Dictionary<int, TechnologyProduct>();

				var dsTechnologyProducts = parDatacontexSDDP.TechnologyProducts
					.Expand(tp => tp.TechnologyCategory)
					.Expand(tp => tp.TechnologyVendor);

				do
					{

					var rsTechnologyProducts =
						from dsTechProduct in dsTechnologyProducts
						where dsTechProduct.Id > intLastReadID
						select dsTechProduct;

					boolFetchMore = false;

					foreach(TechnologyProductsItem record in rsTechnologyProducts)
						{
						TechnologyProduct objTechProduct = new TechnologyProduct();
						objTechProduct.ID = record.Id;
						intLastReadID = record.Id;
						boolFetchMore = true;
						objTechProduct.Title = record.Title;
						TechnologyVendor objTechVendor = new TechnologyVendor();
						objTechVendor.ID = record.TechnologyVendor.Id;
						objTechVendor.Title = record.TechnologyVendor.Title;
						objTechProduct.Vendor = objTechVendor;
						TechnologyCategory objTechCategory = new TechnologyCategory();
						objTechCategory.ID = record.TechnologyCategory.Id;
						objTechCategory.Title = record.TechnologyCategory.Title;
						objTechProduct.Category = objTechCategory;
						objTechProduct.Prerequisites = record.TechnologyPrerequisites;
						this.dsTechnologyProducts.Add(key: record.Id, value: objTechProduct);
						}
					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsTechnologyProducts.Count, DateTime.Now - setStart);

				//--------------------------------
				// Populate the Service Portfolios
				Console.Write("\n\t + ServicePortfolios...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsPortfolios = new Dictionary<int, ServicePortfolio>();
				do
					{

					var rsPortfolios = from dsPortfolio in parDatacontexSDDP.ServicePortfolios
								    where dsPortfolio.Id > intLastReadID
								    select dsPortfolio;

					boolFetchMore = false;

					foreach(var recPortfolio in rsPortfolios)
						{
						ServicePortfolio objPortfolio = new ServicePortfolio();
						objPortfolio.ID = recPortfolio.Id;
						intLastReadID = recPortfolio.Id;
						boolFetchMore = true;
						objPortfolio.Title = recPortfolio.Title;
						objPortfolio.PortfolioType = recPortfolio.PortfolioTypeValue;
						objPortfolio.ISDheading = recPortfolio.ISDHeading;
						objPortfolio.ISDdescription = recPortfolio.ISDDescription;
						objPortfolio.CSDheading = recPortfolio.ContractHeading;
						objPortfolio.CSDdescription = recPortfolio.CSDDescription;
						objPortfolio.SOWheading = recPortfolio.ContractHeading;
						objPortfolio.SOWdescription = recPortfolio.ContractDescription;
						this.dsPortfolios.Add(key: recPortfolio.Id, value: objPortfolio);
						}
					} while(boolFetchMore);
				Console.Write("\t\t {0} \t {1}", this.dsPortfolios.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Families
				Console.Write("\n\t + ServiceFamilies...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsFamilies = new Dictionary<int, ServiceFamily>();
				do
					{

					var rsFamilies = from dsFamily in parDatacontexSDDP.ServiceFamilies
								  where dsFamily.Id > intLastReadID
								  select dsFamily;

					boolFetchMore = false;

					foreach(var recFamily in rsFamilies)
						{
						ServiceFamily objFamily = new ServiceFamily();
						objFamily.ID = recFamily.Id;
						intLastReadID = recFamily.Id;
						boolFetchMore = true;
						objFamily.Title = recFamily.Title;
						objFamily.ServicePortfolioID = recFamily.Service_PortfolioId;
						objFamily.ISDheading = recFamily.ISDHeading;
						objFamily.ISDdescription = recFamily.ISDDescription;
						objFamily.CSDheading = recFamily.ContractHeading;
						objFamily.CSDdescription = recFamily.CSDDescription;
						objFamily.SOWheading = recFamily.ContractHeading;
						objFamily.SOWdescription = recFamily.ContractDescription;
						this.dsFamilies.Add(key: recFamily.Id, value: objFamily);
						}
					} while(boolFetchMore);
				Console.Write("\t\t {0} \t {1}", this.dsFamilies.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Products
				Console.Write("\n\t + ServiceProducts...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsProducts = new Dictionary<int, ServiceProduct>();
				do
					{
					var rsProducts = from dsProduct in parDatacontexSDDP.ServiceProducts
								  where dsProduct.Id > intLastReadID
								  select dsProduct;

					boolFetchMore = false;

					foreach(var recProduct in rsProducts)
						{
						ServiceProduct objProduct = new ServiceProduct();
						objProduct.ID = recProduct.Id;
						intLastReadID = recProduct.Id;
						boolFetchMore = true;
						objProduct.Title = recProduct.Title;
						objProduct.ServiceFamilyID = recProduct.Service_PortfolioId;
						objProduct.ISDheading = recProduct.ISDHeading;
						objProduct.ISDdescription = recProduct.ISDDescription;
						objProduct.CSDheading = recProduct.ContractHeading;
						objProduct.CSDdescription = recProduct.CSDDescription;
						objProduct.SOWheading = recProduct.ContractHeading;
						objProduct.SOWdescription = recProduct.ContractDescription;
						objProduct.KeyClientBenefits = recProduct.KeyClientBenefits;
						objProduct.KeyDDbenefits = recProduct.KeyDDBenefits;
						objProduct.PlannedActivities = recProduct.PlannedActivities;
						objProduct.PlannedActivityEffortDrivers = recProduct.PlannedActivityEffortDrivers;
						objProduct.PlannedDeliverables = recProduct.PlannedDeliverables;
						objProduct.PlannedElements = recProduct.PlannedElements;
						objProduct.PlannedFeatures = recProduct.PlannedFeatures;
						objProduct.PlannedMeetings = recProduct.PlannedMeetings;
						objProduct.PlannedReports = recProduct.PlannedReports;
						objProduct.PlannedServiceLevels = recProduct.PlannedServiceLevels;
						this.dsProducts.Add(key: recProduct.Id, value: objProduct);
						}
					} while(boolFetchMore);
				Console.Write("\t\t {0} \t {1}", this.dsProducts.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Element 
				Console.Write("\n\t + ServiceElements...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsElements = new Dictionary<int, ServiceElement>();
				do
					{
					var rsElements = from dsElement in parDatacontexSDDP.ServiceElements
								  where dsElement.Id > intLastReadID
								  select dsElement;

					boolFetchMore = false;

					foreach(var recElement in rsElements)
						{
						ServiceElement objElement = new ServiceElement();
						objElement.ID = recElement.Id;
						intLastReadID = recElement.Id;
						boolFetchMore = true;
						objElement.Title = recElement.Title;
						objElement.ServiceProductID = recElement.Service_ProductId;
						objElement.SortOrder = recElement.SortOrder;
						objElement.ISDheading = recElement.ISDHeading;
						objElement.ISDdescription = recElement.ISDDescription;
						objElement.KeyClientAdvantages = recElement.KeyClientAdvantages;
						objElement.KeyClientBenefits = recElement.KeyClientBenefits;
						objElement.KeyDDbenefits = recElement.KeyDDBenefits;
						objElement.CriticalSuccessFactors = recElement.CriticalSuccessFactors;
						objElement.ProcessLink = recElement.ProcessLink;
						objElement.KeyPerformanceIndicators = recElement.KeyPerformanceIndicators;
						objElement.ContentLayerValue = recElement.ContentLayerValue;
						objElement.ContentPredecessorElementID = recElement.ContentPredecessorElementId;
						objElement.ContentStatus = recElement.ContentStatusValue;
						this.dsElements.Add(key: recElement.Id, value: objElement);
						}
					} while(boolFetchMore);
				Console.Write("\t\t {0} \t {1}", this.dsElements.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Feature 
				Console.Write("\n\t + ServiceFeatures...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsFeatures = new Dictionary<int, ServiceFeature>();
				do
					{

					var rsFeatures = from dsFeature in parDatacontexSDDP.ServiceFeatures
								  where dsFeature.Id > intLastReadID
								  select dsFeature;

					boolFetchMore = false;

					foreach(var recFeature in rsFeatures)
						{
						ServiceFeature objFeature = new ServiceFeature();

						intLastReadID = recFeature.Id;
						boolFetchMore = true;
						objFeature.ID = recFeature.Id;
						objFeature.Title = recFeature.Title;
						objFeature.ServiceProductID = recFeature.Service_ProductId;
						objFeature.SortOrder = recFeature.SortOrder;
						objFeature.CSDheading = recFeature.ContractHeading;
						objFeature.CSDdescription = recFeature.CSDDescription;
						objFeature.SOWheading = recFeature.ContractHeading;
						objFeature.SOWdescription = recFeature.ContractDescription;
						objFeature.ContentLayerValue = recFeature.ContentLayerValue;
						objFeature.ContentPredecessorFeatureID = recFeature.ContentPredecessorFeatureId;
						objFeature.ContentStatus = recFeature.ContentStatusValue;
						this.dsFeatures.Add(key: recFeature.Id, value: objFeature);
						}
					} while(boolFetchMore);
				Console.Write("\t\t {0} \t {1}", this.dsFeatures.Count, DateTime.Now - setStart);

				//-----------------------
				// Populate Deliverables
				Console.Write("\n\t + Deliverables...\t\t");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsDeliverables = new Dictionary<int, Deliverable>();
				var dsDeliverables = parDatacontexSDDP.Deliverables
					.Expand(dlv => dlv.SupportingSystems)
					.Expand(dlv => dlv.GlossaryAndAcronyms)
					.Expand(dlv => dlv.Responsible_RACI)
					.Expand(dlv => dlv.Accountable_RACI)
					.Expand(dlv => dlv.Consulted_RACI)
					.Expand(dlv => dlv.Informed_RACI);
				do
					{
					var rsDeliverables =
						from dsDeliverable in dsDeliverables
						where dsDeliverable.Id > intLastReadID
						select dsDeliverable;

					boolFetchMore = false;

					foreach(DeliverablesItem recDeliverable in rsDeliverables)
						{
						Deliverable objDeliverable = new Deliverable();
						intLastReadID = recDeliverable.Id;
						boolFetchMore = true;
						objDeliverable.ID = recDeliverable.Id;
						objDeliverable.Title = recDeliverable.Title;
						objDeliverable.DeliverableType = recDeliverable.DeliverableTypeValue;
						objDeliverable.SortOrder = recDeliverable.SortOrder;
						objDeliverable.ISDheading = recDeliverable.ISDHeading;
						objDeliverable.ISDsummary = recDeliverable.ISDSummary;
						objDeliverable.ISDdescription = recDeliverable.ISDDescription;
						objDeliverable.CSDheading = recDeliverable.CSDHeading;
						objDeliverable.CSDsummary = recDeliverable.CSDSummary;
						objDeliverable.CSDdescription = recDeliverable.CSDDescription;
						objDeliverable.SoWheading = recDeliverable.ContractHeading;
						objDeliverable.SoWsummary = recDeliverable.ContractSummary;
						objDeliverable.SoWdescription = recDeliverable.ContractDescription;
						objDeliverable.TransitionDescription = recDeliverable.TransitionDescription;
						objDeliverable.Inputs = recDeliverable.Inputs;
						objDeliverable.Outputs = recDeliverable.Outputs;
						objDeliverable.DDobligations = recDeliverable.SPObligations;
						objDeliverable.ClientResponsibilities = recDeliverable.ClientResponsibilities;
						objDeliverable.Exclusions = recDeliverable.Exclusions;
						objDeliverable.GovernanceControls = recDeliverable.GovernanceControls;
						objDeliverable.WhatHasChanged = recDeliverable.WhatHasChanged;
						objDeliverable.ContentStatus = recDeliverable.ContentStatusValue;
						objDeliverable.ContentLayerValue = recDeliverable.ContentLayerValue;
						objDeliverable.ContentPredecessorDeliverableID = recDeliverable.ContentPredecessor_DeliverableId;
						// Add the Glossary and Acronym terms to the Deliverable object
						if(recDeliverable.GlossaryAndAcronyms.Count > 0)
							{
							foreach(GlossaryAndAcronymsItem recGlossAcronym in recDeliverable.GlossaryAndAcronyms)
								{
								if(objDeliverable.GlossaryAndAcronyms == null)
									{
									objDeliverable.GlossaryAndAcronyms = new Dictionary<int, string>();
									}
								if(objDeliverable.GlossaryAndAcronyms.ContainsKey(recGlossAcronym.Id) == false)
									objDeliverable.GlossaryAndAcronyms.Add(recGlossAcronym.Id, recGlossAcronym.Title);
								}
							}
						// Add the Supporting systems
						if(recDeliverable.SupportingSystems != null)
							{
							objDeliverable.SupportingSystems = new List<string>();
							foreach(var recSupportingSystem in recDeliverable.SupportingSystems)
								{
								objDeliverable.SupportingSystems.Add(recSupportingSystem.Value);
								}
							}

						//Populate the RACI dictionaries
						// --- RACIresponsibles
						if(recDeliverable.Responsible_RACI.Count > 0)
							{
							objDeliverable.RACIresponsibles = new List<int?>();
							foreach(var recJobRole in recDeliverable.Responsible_RACI)
								{
								objDeliverable.RACIresponsibles.Add(recJobRole.Id);
								}
							}

						// --- RACIaccountables
						if(recDeliverable.Accountable_RACI != null)
							{
							objDeliverable.RACIaccountables = new List<int?>();
							if(recDeliverable.Accountable_RACI != null)
								{
								objDeliverable.RACIaccountables.Add(recDeliverable.Accountable_RACIId);
								}
							}
						// --- RACIconsulteds
						if(recDeliverable.Consulted_RACI.Count > 0)
							{
							objDeliverable.RACIconsulteds = new List<int?>();
							foreach(var recJobRole in recDeliverable.Consulted_RACI)
								{
								objDeliverable.RACIconsulteds.Add(recJobRole.Id);
								}
							}
						// --- RACIinformeds
						if(recDeliverable.Informed_RACI.Count > 0)
							{
							objDeliverable.RACIinformeds = new List<int?>();
							foreach(var recJobRole in recDeliverable.Informed_RACI)
								{
								JobRole objJobRole = new JobRole();
								objJobRole.ID = recJobRole.Id;
								objJobRole.Title = recJobRole.Title;
								objDeliverable.RACIinformeds.Add(recJobRole.Id);
								}
							}
						this.dsDeliverables.Add(key: recDeliverable.Id, value: objDeliverable);
						}
					} while(boolFetchMore);
				Console.Write("\t\t {0} \t {1}", this.dsDeliverables.Count, DateTime.Now - setStart);

				//--------------------------------------
				// Populate Service Element Deliverables
				Console.Write("\n\t + ElementDeliverables...\t");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsElementDeliverables = new Dictionary<int, ElementDeliverable>();
				do
					{
					var rsElementDeliverable = from dsElementDeliverable in parDatacontexSDDP.ElementDeliverables
										  where dsElementDeliverable.Id > intLastReadID
										  select dsElementDeliverable;

					boolFetchMore = false;

					foreach(var recElementDeliverable in rsElementDeliverable)
						{
						ElementDeliverable objElementDeliverable = new ElementDeliverable();
						intLastReadID = recElementDeliverable.Id;
						boolFetchMore = true;
						objElementDeliverable.ID = recElementDeliverable.Id;
						objElementDeliverable.Title = recElementDeliverable.Title;
						objElementDeliverable.AssociatedDeliverableID = recElementDeliverable.Deliverable_Id;
						objElementDeliverable.AssociatedElementID = recElementDeliverable.Service_ElementId;
						objElementDeliverable.Optionality = recElementDeliverable.OptionalityValue;
						this.dsElementDeliverables.Add(key: recElementDeliverable.Id, value: objElementDeliverable);
						}
					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsElementDeliverables.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate Service Feature Deliverables
				Console.Write("\n\t + FeatureDeliverables...\t");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsFeatureDeliverables = new Dictionary<int, FeatureDeliverable>();
				do
					{
					var rsFeatureDeliverable = from dsFeatureDeliverable in parDatacontexSDDP.FeatureDeliverables
										  where dsFeatureDeliverable.Id > intLastReadID
										  select dsFeatureDeliverable;

					boolFetchMore = false;

					foreach(var recFeatureDeliverable in rsFeatureDeliverable)
						{
						FeatureDeliverable objFeatureDeliverable = new FeatureDeliverable();
						intLastReadID = recFeatureDeliverable.Id;
						boolFetchMore = true;
						objFeatureDeliverable.ID = recFeatureDeliverable.Id;
						objFeatureDeliverable.Title = recFeatureDeliverable.Title;
						objFeatureDeliverable.AssociatedDeliverableID = recFeatureDeliverable.Deliverable_Id;
						objFeatureDeliverable.AssociatedFeatureID = recFeatureDeliverable.Service_FeatureId;
						objFeatureDeliverable.Optionality = recFeatureDeliverable.OptionalityValue;
						this.dsFeatureDeliverables.Add(key: recFeatureDeliverable.Id, value: objFeatureDeliverable);
						}
					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsFeatureDeliverables.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate DeliverableTechnologies
				Console.Write("\n\t + DeliverableTechnologies...\t");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsDeliverableTechnologies = new Dictionary<int, DeliverableTechnology>();

				do
					{
					var rsDeliverableTechnologies = from dsDeliverableTechnology in parDatacontexSDDP.DeliverableTechnologies
											  where dsDeliverableTechnology.Id > intLastReadID
											  select dsDeliverableTechnology;

					boolFetchMore = false;

					foreach(var recDeliverableTechnology in rsDeliverableTechnologies)
						{
						DeliverableTechnology objDeliverableTechnology = new DeliverableTechnology();
						intLastReadID = recDeliverableTechnology.Id;
						boolFetchMore = true;
						objDeliverableTechnology.ID = recDeliverableTechnology.Id;
						objDeliverableTechnology.Title = recDeliverableTechnology.Title;
						objDeliverableTechnology.Considerations = recDeliverableTechnology.TechnologyConsiderations;
						objDeliverableTechnology.RoadmapStatus = recDeliverableTechnology.TechnologyRoadmapStatusValue;
						objDeliverableTechnology.Deliviverable = this.dsDeliverables
							.Where(d => d.Key == recDeliverableTechnology.Deliverable_Id).FirstOrDefault().Value;
						objDeliverableTechnology.TechnologyProduct = this.dsTechnologyProducts
							.Where(t => t.Key == recDeliverableTechnology.TechnologyProductsId).FirstOrDefault().Value;
						this.dsDeliverableTechnologies.Add(key: recDeliverableTechnology.Id, value: objDeliverableTechnology);
						}
					} while(boolFetchMore);
				Console.Write(" {0} \t {1}", this.dsDeliverableTechnologies.Count, DateTime.Now - setStart);

				// -------------------------
				// Populate Activities
				Console.Write("\n\t + Activities...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsActivities = new Dictionary<int, Activity>();
				var datasetActivities = parDatacontexSDDP.Activities
					.Expand(ac => ac.Activity_Category)
					.Expand(ac => ac.OLA_);

				do
					{
					var rsActivities =
						from dsActivities in datasetActivities
						where dsActivities.Id > intLastReadID
						select dsActivities;

					boolFetchMore = false;

					foreach(ActivitiesItem record in rsActivities)
						{
						Activity objActivity = new Activity();
						intLastReadID = record.Id;
						boolFetchMore = true;
						objActivity.ID = record.Id;
						objActivity.Title = record.Title;
						objActivity.SortOrder = record.SortOrder;
						objActivity.Catagory = record.Activity_Category.Title;
						objActivity.Assumptions = record.ActivityAssumptions;
						objActivity.ContentStatus = record.ContentStatusValue;
						objActivity.ISDheading = record.ISDHeading;
						objActivity.ISDdescription = record.ISDDescription;
						objActivity.Input = record.ActivityInput;
						objActivity.Output = record.ActivityOutput;
						objActivity.CSDheading = record.CSDHeading;
						objActivity.CSDdescription = record.CSDDescription;
						objActivity.SOWheading = record.CSDDescription;
						if(record.OLA_ != null)
							objActivity.OLA = record.OLA_.Title;
						objActivity.OLAvariations = record.OLAVariations;
						objActivity.Optionality = record.ActivityOptionalityValue;
						if(record.Accountable_RACI != null)
							{
							objActivity.RACI_Accountable = new List<JobRole>();
							objActivity.RACI_Accountable.Add(this.dsJobroles
								.Where(j => j.Key == record.Accountable_RACIId).FirstOrDefault().Value);
							}
						if(record.Responsible_RACI != null && record.Responsible_RACI.Count() > 0)
							{
							objActivity.RACI_Responsible = new List<JobRole>();
							foreach(var entryJobRole in record.Responsible_RACI)
								{
								objActivity.RACI_Responsible.Add(this.dsJobroles
								.Where(j => j.Key == entryJobRole.Id).FirstOrDefault().Value);
								}
							}
						if(record.Consulted_RACI != null && record.Consulted_RACI.Count() > 0)
							{
							objActivity.RACI_Consulted = new List<JobRole>();
							foreach(var entryJobRole in record.Consulted_RACI)
								{
								objActivity.RACI_Consulted.Add(this.dsJobroles
								.Where(j => j.Key == entryJobRole.Id).FirstOrDefault().Value);
								}
							}
						if(record.Informed_RACI != null && record.Informed_RACI.Count() > 0)
							{
							objActivity.RACI_Informed = new List<JobRole>();
							foreach(var entryJobRole in record.Informed_RACI)
								{
								objActivity.RACI_Informed.Add(this.dsJobroles
								.Where(j => j.Key == entryJobRole.Id).FirstOrDefault().Value);
								}
							}
						this.dsActivities.Add(key: record.Id, value: objActivity);
						}
					} while(boolFetchMore);
				Console.Write("\t\t\t\t {0} \t {1}", this.dsActivities.Count, DateTime.Now - setStart);


				//---------------------------------------
				// Populate DeliverableActivities
				//---------------------------------------
				Console.Write("\n\t + DeliverableActivities...\t");
				intLastReadID = 0;
				setStart = DateTime.Now;
				this.dsDeliverableActivities = new Dictionary<int, DeliverableActivity>();
				do
					{
					var rsDeliverableActivities = from dsDeliverableActivity in parDatacontexSDDP.DeliverableActivities
											where dsDeliverableActivity.Id > intLastReadID
											select dsDeliverableActivity;

					boolFetchMore = false;

					foreach(var recDeliverableActivity in rsDeliverableActivities)
						{
						DeliverableActivity objDeliverableActivity = new DeliverableActivity();
						intLastReadID = recDeliverableActivity.Id;
						boolFetchMore = true;
						objDeliverableActivity.ID = recDeliverableActivity.Id;
						objDeliverableActivity.Title = recDeliverableActivity.Title;
						objDeliverableActivity.Optionality = recDeliverableActivity.OptionalityValue;
						objDeliverableActivity.AssociatedActivityID = recDeliverableActivity.Activity_Id;
						objDeliverableActivity.AssociatedDeliverableID = recDeliverableActivity.Deliverable_Id;
						objDeliverableActivity.AssociatedDeliverable = this.dsDeliverables
							.Where(d => d.Key == recDeliverableActivity.Deliverable_Id).FirstOrDefault().Value;
						objDeliverableActivity.AssociatedActivity = this.dsActivities
							.Where(a => a.Key == recDeliverableActivity.Activity_Id).FirstOrDefault().Value;
						this.dsDeliverableActivities.Add(key: recDeliverableActivity.Id, value: objDeliverableActivity);
						}
					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsDeliverableActivities.Count, DateTime.Now - setStart);

				// -------------------------
				// Populate ServiceLevels
				// -------------------------
				Console.Write("\n\t + ServiceLevels...\t");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsServiceLevels = new Dictionary<int, ServiceLevel>();
				var datasetServiceLevels = parDatacontexSDDP.ServiceLevels
					.Expand(sl => sl.Service_Hour);

				do
					{
					var rsServiceLevels =
						from dsServiceLevel in datasetServiceLevels
						where dsServiceLevel.Id > intLastReadID
						select dsServiceLevel;

					boolFetchMore = false;

					foreach(ServiceLevelsItem record in rsServiceLevels)
						{
						ServiceLevel objServiceLevel = new ServiceLevel();
						intLastReadID = record.Id;
						boolFetchMore = true;
						objServiceLevel.ID = record.Id;
						objServiceLevel.Title = record.Title;
						objServiceLevel.ISDheading = record.ISDHeading;
						objServiceLevel.ISDdescription = record.ISDDescription;
						objServiceLevel.CSDheading = record.CSDHeading;
						objServiceLevel.CSDdescription = record.CSDDescription;
						objServiceLevel.BasicConditions = record.BasicServiceLevelConditions;
						objServiceLevel.CalcualtionMethod = record.CalculationMethod;
						objServiceLevel.CalculationFormula = record.CalculationFormula;
						objServiceLevel.ContentStatus = record.ContentStatusValue;
						objServiceLevel.Measurement = record.ServiceLevelMeasurement;
						objServiceLevel.MeasurementInterval = record.MeasurementIntervalValue;
						objServiceLevel.SOWheading = record.ContractHeading;
						objServiceLevel.SOWdescription = record.ContractDescription;
						objServiceLevel.ReportingInterval = record.ReportingIntervalValue;
						if(record.Service_HourId != null)
							objServiceLevel.ServiceHours = record.Service_Hour.Title;
						objServiceLevel.BasicConditions = record.BasicServiceLevelConditions;
						// ---------------------------------------------
						// Load the Service Level Performance Thresholds
						// ---------------------------------------------
						var dsThresholds =
							from dsThreshold in parDatacontexSDDP.ServiceLevelTargets
							where dsThreshold.Service_LevelId == record.Id && dsThreshold.ThresholdOrTargetValue == "Threshold"
							orderby dsThreshold.Title
							select dsThreshold;

						if(dsThresholds.Count() > 0)
							{
							objServiceLevel.PerfomanceThresholds = new List<ServiceLevelTarget>();
							foreach(var thresholdItem in dsThresholds)
								{
								ServiceLevelTarget objSLthreshold = new ServiceLevelTarget();
								objSLthreshold.ID = thresholdItem.Id;
								objSLthreshold.Title = thresholdItem.Title.Substring(thresholdItem.Title.IndexOf(": ", 0) + 2,
									thresholdItem.Title.Length - thresholdItem.Title.IndexOf(": ", 0) - 2);
								objSLthreshold.Type = thresholdItem.ThresholdOrTargetValue;
								objSLthreshold.ContentStatus = thresholdItem.ContentStatusValue;
								objServiceLevel.PerfomanceThresholds.Add(objSLthreshold);
								}
							}
						//--------------------------------------------
						// Load the Service Level Performance Targets
						//--------------------------------------------
						var dsTargets =
							from dsThreshold in parDatacontexSDDP.ServiceLevelTargets
							where dsThreshold.Service_LevelId == record.Id && dsThreshold.ThresholdOrTargetValue == "Target"
							orderby dsThreshold.Title
							select dsThreshold;

						if(dsTargets.Count() > 0)
							{
							objServiceLevel.PerformanceTargets = new List<ServiceLevelTarget>();
							foreach(var targetEntry in dsTargets)
								{
								ServiceLevelTarget objSLtarget = new ServiceLevelTarget();
								objSLtarget.ID = targetEntry.Id;
								objSLtarget.Title = targetEntry.Title.Substring(targetEntry.Title.IndexOf(": ", 0) + 2,
									(targetEntry.Title.Length - targetEntry.Title.IndexOf(": ", 0) - 2));
								objSLtarget.Type = targetEntry.ThresholdOrTargetValue;
								objSLtarget.ContentStatus = targetEntry.ContentStatusValue;
								objServiceLevel.PerformanceTargets.Add(objSLtarget);
								}
							}
						this.dsServiceLevels.Add(key: record.Id, value: objServiceLevel);
						}
					} while(boolFetchMore);
				Console.Write("\t\t\t {0} \t {1}", this.dsServiceLevels.Count, DateTime.Now - startTime);

				//---------------------------------------
				// Populate DeliverableServiceLevels
				Console.Write("\n\t + DeliverableServiceLevels...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsDeliverableServiceLevels = new Dictionary<int, DeliverableServiceLevel>();
				do
					{
					var rsDeliverableServiceLevels = from dsDeliverableServiceLevel in parDatacontexSDDP.DeliverableServiceLevels
											   where dsDeliverableServiceLevel.Id > intLastReadID
											   select dsDeliverableServiceLevel;

					boolFetchMore = false;

					foreach(var record in rsDeliverableServiceLevels)
						{
						DeliverableServiceLevel objDeliverableServiceLevel = new DeliverableServiceLevel();
						intLastReadID = record.Id;
						boolFetchMore = true;
						objDeliverableServiceLevel.ID = record.Id;
						objDeliverableServiceLevel.Title = record.Title;
						objDeliverableServiceLevel.Optionality = record.OptionalityValue;
						objDeliverableServiceLevel.ContentStatus = record.ContentStatusValue;
						objDeliverableServiceLevel.AdditionalConditions = record.AdditionalConditions;
						objDeliverableServiceLevel.AssociatedDeliverableID = record.Service_LevelId;
						objDeliverableServiceLevel.AssociatedServiceLevelID = record.Service_LevelId;
						objDeliverableServiceLevel.AssociatedServiceProductID = record.Service_ProductId;
						objDeliverableServiceLevel.AssociatedDeliverable = this.dsDeliverables
							.Where(d => d.Key == record.Deliverable_Id).FirstOrDefault().Value;
						objDeliverableServiceLevel.AssociatedServiceLevel = this.dsServiceLevels
							.Where(a => a.Key == record.Service_LevelId).FirstOrDefault().Value;
						objDeliverableServiceLevel.AssociatedServiceProduct = this.dsProducts
							.Where(p => p.Key == record.Service_ProductId).FirstOrDefault().Value;
						this.dsDeliverableServiceLevels.Add(key: record.Id, value: objDeliverableServiceLevel);
						}
					} while(boolFetchMore);
				Console.WriteLine("\t {0} \t {1}", this.dsDeliverableServiceLevels.Count, DateTime.Now - setStart);

				Console.WriteLine("\n\tPopulating the complete DataSet took ended at {0} and took {1}.", DateTime.Now, DateTime.Now - startTime);
				return true;
				}
			catch(DataServiceClientException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nStatusCode: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.StatusCode, exc.StackTrace);
				return false;
				}
			catch(DataServiceQueryException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nResponse: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				return false;
				}
			catch(DataServiceTransportException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nResponse:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				return false;
				}
			catch(System.Net.Sockets.SocketException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nTargetSite:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
				return false;
				}
			catch(Exception exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1}\nSource:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Source, exc.StackTrace);
				return false;
				}
			}

		public bool PopulateMappingObjects(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parMapping)
			{
			int intLastReadID = 0;
			bool boolFetchMore = false;
			DateTime startTime = DateTime.Now;
			DateTime setStart = DateTime.Now;
			// Please Note: 
			// SharePoint's REST API has a limit which returns only 1000 entries at a time
			// therefore a paging mechanism is implemented to return all the entries in the List.

			try
				{
				Console.Write("\nPopulating the complete Mappings DataSet...");
				//-------------------------------------------------------------
				// Populate Mapping
				Console.Write("\n\t + Mappings...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsMappings = new Dictionary<int?, Mapping>();
				var datasetMappings = parDatacontexSDDP.Mappings
					.Expand(m => m.Client_);

				var rsMappings =	
					from dsMapping in datasetMappings
					where dsMapping.Id == parMapping 
					select dsMapping;

				var recordM = rsMappings.First();

				if(recordM != null)
					{
					Mapping objMapping = new Mapping();
					objMapping.ID = recordM.Id;
					objMapping.Title = recordM.Title;
					objMapping.ClientName = recordM.Client_.DocGenClientName;
					this.dsMappings.Add(recordM.Id, objMapping);
					}

				Console.Write("\t\t\t\t {0} \t {1}", this.dsMappings.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate Mapping Service Towers
				Console.Write("\n\t + MappingServiceTowers...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsMappingServiceTowers = new Dictionary<int, MappingServiceTower>();
				do
					{
					var rsMappingServiceTowers = from dsMappingServiceTowers in parDatacontexSDDP.MappingServiceTowers
											   where dsMappingServiceTowers.Mapping_Id == parMapping
											   && dsMappingServiceTowers.Id > intLastReadID
											   select dsMappingServiceTowers;

					boolFetchMore = false;

					foreach(var recordMST in rsMappingServiceTowers)
						{
						MappingServiceTower objMappingServiceTower = new MappingServiceTower();
						intLastReadID = recordMST.Id;
						boolFetchMore = true;
						objMappingServiceTower.ID = recordMST.Id;
						objMappingServiceTower.Title = recordMST.Title;
						this.dsMappingServiceTowers.Add(recordMST.Id, objMappingServiceTower);
						}
					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsMappingServiceTowers.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate Mapping Requirements
				Console.Write("\n\t + MappingRequirements...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsMappingRequirements = new Dictionary<int, MappingRequirement>();
				do
					{
					var rsMappingRequirements = 
						from dsMappingRequirements in parDatacontexSDDP.MappingRequirements
						where dsMappingRequirements.Mapping_Id == parMapping
							&& dsMappingRequirements.Id > intLastReadID
						select dsMappingRequirements;

					boolFetchMore = false;

					foreach(var recordMR in rsMappingRequirements)
						{
						MappingRequirement objMappingRequirement = new MappingRequirement();
						intLastReadID = recordMR.Id;
						boolFetchMore = true;
						objMappingRequirement.ID = recordMR.Id;
						objMappingRequirement.Title = recordMR.Title;
						objMappingRequirement.MappingServiceTowerID = recordMR.Mapping_TowerId;
						objMappingRequirement.ComplianceComments = recordMR.ComplianceComments;
						objMappingRequirement.ComplianceStatus = recordMR.ComplianceStatusValue;
						objMappingRequirement.RequirementServiceLevel = recordMR.RequirementServiceLevel;
						objMappingRequirement.RequirementText = recordMR.RequirementText;
						objMappingRequirement.SourceReference = recordMR.SourceReference;
						objMappingRequirement.SortOrder = recordMR.SortOrder;
						this.dsMappingRequirements.Add(key: recordMR.Id, value: objMappingRequirement);
						}
					} while(boolFetchMore);
				Console.Write("\t {0} \t {1}", this.dsMappingRequirements.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate Mapping Assumptions, Risks, Deliverables
				Console.Write("\n\t + Mapping Assumptions, Risks, Deliverables...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsMappingAssumptions = new Dictionary<int, MappingAssumption>();
				this.dsMappingRisks = new Dictionary<int, MappingRisk>();
				this.dsMappingDeliverables = new Dictionary<int, MappingDeliverable>();
				this.dsMappingServiceLevels = new Dictionary<int, MappingServiceLevel>();

				// Populate the Mapping Requirements
				if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
					{
					foreach(var itemRequirement in this.dsMappingRequirements)
						{
						var rsMappingAssumptions =
							from dsMappingAssumptions in parDatacontexSDDP.MappingAssumptions
							where dsMappingAssumptions.Mapping_RequirementId == itemRequirement.Key
							select dsMappingAssumptions;

						// Populate the Mapping Assumptions
						foreach(var recordMA in rsMappingAssumptions)
							{
							MappingAssumption objMappingAssumption = new MappingAssumption();
							objMappingAssumption.ID = recordMA.Id;
							objMappingAssumption.MappingRequirementID = recordMA.Mapping_RequirementId;
							objMappingAssumption.Title = recordMA.Title;
							objMappingAssumption.Description = recordMA.AssumptionDescription;
							this.dsMappingAssumptions.Add(key: recordMA.Id, value: objMappingAssumption);
							} //foreach(var recordMA in rsMappingAssumptions)

						// Populate the Mapping Risks
						var rsMappingRisks =
							from dsMappingRisks in parDatacontexSDDP.MappingRisks
							where dsMappingRisks.Mapping_RequirementId == itemRequirement.Key
							select dsMappingRisks;

						foreach(var recordRisk in rsMappingRisks)
							{
							MappingRisk objMappingRisk = new MappingRisk();
							objMappingRisk.ID = recordRisk.Id;
							objMappingRisk.MappingRequirementID = recordRisk.Mapping_RequirementId;
							objMappingRisk.Title = recordRisk.Title;
							objMappingRisk.Statement = recordRisk.RiskStatement;
							objMappingRisk.Status = recordRisk.RiskStatusValue;
							objMappingRisk.Mitigation = recordRisk.RiskMitigation;
							objMappingRisk.Exposure = recordRisk.RiskExposureValue;
							objMappingRisk.ExposureValue = recordRisk.RiskExposureValue0;
							this.dsMappingRisks.Add(key: recordRisk.Id, value: objMappingRisk);
							} //foreach(var recordRisk in rsMappingRisks)

						// Populate the MapingRisks
						var rsMappingDeliverables =
							from dsMappingDeliverable in parDatacontexSDDP.MappingDeliverables
							where dsMappingDeliverable.Mapping_RequirementId == itemRequirement.Key
							select dsMappingDeliverable;

						foreach(var recordMappingDeliv in rsMappingDeliverables)
							{
							MappingDeliverable objMappingDeliverable = new MappingDeliverable();
							objMappingDeliverable.ID = recordMappingDeliv.Id;
							objMappingDeliverable.MappingRequirementID = recordMappingDeliv.Mapping_RequirementId;
							objMappingDeliverable.Title = recordMappingDeliv.Title;
							if(recordMappingDeliv.DeliverableChoiceValue == "New")
								objMappingDeliverable.NewDeliverable = true;
							else
								objMappingDeliverable.NewDeliverable = false;
							objMappingDeliverable.MappedDeliverableID = recordMappingDeliv.Mapped_DeliverableId;
							objMappingDeliverable.NewRequirement = recordMappingDeliv.DeliverableRequirement;
							objMappingDeliverable.ComplianceComments = recordMappingDeliv.ComplianceComments;
							this.dsMappingDeliverables.Add(key: recordMappingDeliv.Id, value: objMappingDeliverable);

							// Populate the Mapping Service Levels
							var rsMappingServiceLevels =
							from dsMappingServiceLevel in parDatacontexSDDP.MappingServiceLevels
							where dsMappingServiceLevel.Mapping_DeliverableId == recordMappingDeliv.Id
							select dsMappingServiceLevel;

							foreach(var recordMSL in rsMappingServiceLevels)
								{
								MappingServiceLevel objMappingServiceLevel = new MappingServiceLevel();
								objMappingServiceLevel.ID = recordMSL.Id;
								objMappingServiceLevel.Title = recordMSL.Title;
								objMappingServiceLevel.MappedDeliverableID = recordMSL.Mapping_DeliverableId;
								objMappingServiceLevel.NewServiceLevel = recordMSL.NewServiceLevel;
								objMappingServiceLevel.MappedServiceLevelID = recordMSL.Service_LevelId;
								objMappingServiceLevel.RequirementText = recordMSL.ServiceLevelRequirement;
								this.dsMappingServiceLevels.Add(key: recordMSL.Id, value: objMappingServiceLevel);
								} //foreach(var recordMA in rsMappingAssumptions)
							} //foreach(var recordDeliv in rsMappingDeliverable)
						} // foreach(var itemRequirement in this.dsMappingRequirements)
					} // if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)

				Console.Write("\n\t\t + MappingAssumptions:\t\t{0}", this.dsMappingAssumptions.Count);
				Console.Write("\n\t\t + MappingRisks:\t\t\t{0}", this.dsMappingRisks.Count);
				Console.Write("\n\t\t + MappingDeliverables:\t\t{0}", this.dsMappingDeliverables.Count);
				Console.Write("\n\t\t + MappingServiceLevels:\t{0}", this.dsMappingServiceLevels.Count);

				Console.Write("\n\t = it Took {0}", DateTime.Now - setStart);

				Console.WriteLine("\n\tPopulating the Mappings DataSet ended at {0} and took {1}.", DateTime.Now, DateTime.Now - startTime);
				return true;
				}
			catch(DataServiceClientException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1} - StatusCode:{2}\n{3}.", exc.HResult, exc.Message, exc.StatusCode, exc.StackTrace);
				return false;
				}
			catch(DataServiceQueryException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1} - StatusCode:{2}\n{3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				return false;
				}
			catch(DataServiceTransportException exc)
				{
				Console.WriteLine("\n*** Exception ERROR ***\n{0} - {1} \n{3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				return false;
				}
			}
		}
	}

