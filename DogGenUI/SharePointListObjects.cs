using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using DocGeneratorCore.SDDPServiceReference;

namespace DocGeneratorCore
	{
	//+ GlossaryAcronym
	public class GlossaryAcronym
		{
		public int ID{get; set;}
		public string Term{get; set;}
		public string Meaning{get; set;}
		public string Acronym{get; set;}
		public DateTime? Modified{get; set;}
		} // end of Class GlossaryAndAcronym

	//+ ServicePortfolio
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
		public DateTime LastRefreshedOn{get; set;}
		}

	//+ ServiceFamily
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
		public DateTime LastRefreshedOn{get; set;}
		} // end of class ServicePFamily

	//+ ServiceProduct
	/// <summary>
	/// Service Product object represent an entry in the Service Products SharePoint List
	/// </summary>
	public class ServiceProduct
		{
		public int ID{get; set;}
		public int? ServiceFamilyID{get; set;}
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
		public DateTime LastRefreshedOn{get; set;}
		} // end of class ServiceProduct

	//+ServiceElement
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
		public DateTime LastRefreshedOn{get; set;}
		} // end Class ServiceElement

	//+ ServiceFeature
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
		public DateTime LastRefreshedOn{get; set;}
		} // end Class ServiceFeature

	//+ Deliverables
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
		public DateTime LastRefreshedOn{get; set;}
		} // end Class Deliverables

	//+ DeliverableServiceLevels
	/// <summary>
	///
	/// </summary>
	public class DeliverableServiceLevel
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string ContentStatus{get; set;}
		public string Optionality{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public int? AssociatedServiceLevelID{get; set;}
		public int? AssociatedServiceProductID{get; set;}
		public string AdditionalConditions{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		}// end of DeliverableServiceLevels class

	//+ DeliverableActivities
	/// <summary>
	///
	/// </summary>
	public class DeliverableActivity
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public int? AssociatedActivityID{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		}// end of DeliverableActivities class

	//+ DeliverableTechnology
	/// <summary>
	/// This object represents an entry in the DeliverableTechnologies SharePoint List
	/// Each entry in the list is a DeliverableTechnology object.
	/// </summary>
	public class DeliverableTechnology
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Considerations{get; set;}
		public int? TechnologyProductID{get; set;}
		public int? DeliviverableID{get; set;}
		public string RoadmapStatus{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end of TechnologyProduct class

	//+ FeatureDeliverable
	/// <summary>
	/// The FeatureDeliverable object is the junction table or the cross-reference table between Service Features and Deliverables.
	/// </summary>
	public class FeatureDeliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public int? AssociatedFeatureID{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end of FeatureDeliverable class

	//+ ElementDeliverable
	/// <summary>
	/// The ElementDeliverable objects is the junction table or the cross-reference table between Service Elements and Deliverables.
	/// </summary>
	public class ElementDeliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public int? AssociatedElementID{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end of ElementDeliverable class

	//+ Mapping
	/// <summary>
	/// The Mapping object represents an entry in the Mappings List in SharePoint.
	/// </summary>
	public class Mapping
		{
		public int? ID{get; set;}
		public string Title{get; set;}
		public string ClientName{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end Class Mapping

	//+ MappingServiceTower
	/// <summary>
	/// The MappingServiceTower object represents an entry in the Mapping Service Towers List in SharePoint.
	/// </summary>
	public class MappingServiceTower
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end Class Mapping Service Towers

	//+ MappingRequirement
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
		public DateTime LastRefreshedOn{get; set;}
		} // end Class Mapping Requirements

	//+ MappingDeliverable
	/// <summary>
	/// The Mapping Deliverable is the class used to for the Mapping Deliverables SharePoint List.
	/// </summary>
	//############################################
	public class MappingDeliverable
		{
		public int ID{get; set;}
		public int? MappingRequirementID{get; set;}
		public string Title{get; set;}
		/// <summary>
		/// Represents the translated value in the Deliverable Choice column of the MappingDeliverable List. TRUE if "New" else FALSE
		/// </summary>
		public bool NewDeliverable{get; set;}
		public string ComplianceComments{get; set;}
		public String NewRequirement{get; set;}
		public int? MappedDeliverableID{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		}

	//+ MappingAssumption
	/// <summary>
	/// The MappingAssumption represents an entry of the Mapping Assumptions List in SharePoint
	/// </summary>
	public class MappingAssumption
		{
		public int ID{get; set;}
		public int? MappingRequirementID{get; set;}
		public string Title{get; set;}
		public string Description{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		}

	//+ MaapingRisk
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
		public DateTime LastRefreshedOn{get; set;}
		} // End of Class MappingRisk


	//+ MappingServiceLevel
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
		public DateTime LastRefreshedOn{get; set;}
		}

	//+ ServiceLevel
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
		public DateTime LastRefreshedOn{get; set;}
		} // end of Service Levels class

	//+ SeviceLevelTarget
	/// <summary>
	/// This object repsents an entry in the Activities SharePoint List
	/// </summary>
	public class ServiceLevelTarget
		{
		public int ID{get; set;}
		public string Type{get; set;}
		public string Title{get; set;}
		public string ContentStatus{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		}

	//+ Activity
	/// <summary>
	/// This object repsents an entry in the Activities SharePoint List
	/// </summary>
	public class Activity
		{
		public int ID {get; set;}

		public string Title {get; set;}

		public double? SortOrder {get; set;}

		public string Optionality {get; set;}

		public string ISDheading {get; set;}

		public string ISDdescription {get; set;}

		public string CSDheading {get; set;}

		public string CSDdescription {get; set;}

		public string SOWheading {get; set;}

		public string SOWdescription {get; set;}

		public string ContentStatus {get; set;}

		public string Input {get; set;}

		public string Output {get; set;}

		public string Catagory {get; set;}

		public string Assumptions {get; set;}

		public string OLAvariations {get; set;}

		public string OLA {get; set;}

		public List<int> RACI_ResponsibleID {get; set;}

		public List<int?> RACI_AccountableID {get; set;}

		public List<int> RACI_ConsultedID {get; set;}

		public List<int> RACI_InformedID {get; set;}

		public DateTime LastRefreshedOn {get; set;}

		public string OwningEntity { get; set; }
		} // end of Activitiy class

	//+ JobRole
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
		public DateTime LastRefreshedOn{get; set;}
		} // end of JobRole class

	//+ TechnologyCategory
	/// <summary>
	/// This object repsents an entry in the Technology Categories SharePoint List
	/// Each entry in the list is a Technology Category object.
	/// </summary>
	public class TechnologyCategory
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end of TechnologyCategory class

	//+ technologyVendor
	/// <summary>
	/// This object repsents an entry in the Technology Vendors SharePoint List
	/// Each entry in the list is a Technology Vendor object.
	/// </summary>
	public class TechnologyVendor
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		} // end of TechnologyVendor class

	//+ TechnologyProduct
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
		public DateTime LastRefreshedOn{get; set;}
		} // end of TechnologyProduct class

	public class CompleteDataSet
		{
		
		public Dictionary<int, GlossaryAcronym> dsGlossaryAcronyms{get; set;}
		public Dictionary<int, JobRole> dsJobroles{get; set;}
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
		public DesignAndDeliveryPortfolioDataContext SDDPdatacontext{get; set;}
		public DateTime LastRefreshedOn{get; set;}
		public DateTime RefreshingDateTimeStamp{get; set;}
		public bool IsDataSetComplete{get; set;}
		public string SharePointSiteURL { get; set; }
		public string SharePointSiteSubURL { get; set; }

		//- These variables are the **Thread Controller objects** which handle the locking of the data threads in the following methods:
		//- **PopulateBaseDataset** and **PopulateMappingDataset**
		private static readonly Object lockThread1 = new Object();

		private static readonly Object lockThread2 = new Object();
		private static readonly Object lockThread3 = new Object();
		private static readonly Object lockThread4 = new Object();
		private static readonly Object lockThread5 = new Object();
		private static readonly Object lockThread6 = new Object();
		private static readonly Object lockThread7 = new Object();
		private static readonly Object lockThreadSynchro = new Object();

		//- Specify the CountdownEvent which is used to **WAIT** until all the DatasetPopulation threads complete
		//- after which it set the **IsDataSetComplete** to True;
		public static CountdownEvent threadCountDown = new CountdownEvent(6);

		//===G
		//++ PopulateBaseDataObjects
		/// <summary>
		/// 	This method populate the complete Dataset from SharePoint into Memory stored in the object's DataSet property
		/// Any failure (exception will result in an incomplete data set indicted by the IsDataSetComplete = false.
		/// </summary>
		public void PopulateBaseDataObjects()
			{
			try
				{
				this.IsDataSetComplete = false;
				Stopwatch objStopWatchCompleteDataSet;

				//- Control ** Thread Processing **
				switch(Thread.CurrentThread.Name)
					{
				case ("Data1"):
					goto Thread1start;
				case ("Data2"):
					goto Thread2start;
				case ("Data3"):
					goto Thread3start;
				case ("Data4"):
					goto Thread4start;
				case ("Data5"):
					goto Thread5start;
				case ("Data6"):
					goto Thread6start;
				case ("Data7"):
					goto Thread7start;
				default:
						{
						//- Any other thread is considered the **Sychronisation** thread that will wait until
						//- all the DataThreads completed and then return to the caller...
						threadCountDown.Reset(7);
						goto ThreadSynchroStart;
						}
					}

//+ Please Note:
//- -------------------------------------------------------------------------------------
//- SharePoint's REST API has a limit which returns only 1000 entries at a time
//- therefore a paging principle is implemented to return all the entries in the List.
//- -------------------------------------------------------------------------------------
//---g
Thread1start:
				lock(lockThread1)
					{
					Console.Write("\n### + Thread 1 Start... ###");
					Stopwatch stopwatchThread1 = Stopwatch.StartNew();
					//- -----------------------------------
					// Populate **GlossaryAcronyms**
					//- ----------------------------------
					int intEntriesCounter1 = 0;
					Stopwatch objStopWatch1 = Stopwatch.StartNew();
					int intLastReadID1 = 0;
					bool bFetchMore1 = true;

					DateTime dtLastRefreshOn1 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsGlossaryAcronyms == null)
						this.dsGlossaryAcronyms = new Dictionary<int, GlossaryAcronym>();
					else
						dtLastRefreshOn1 = this.LastRefreshedOn;

					while(bFetchMore1)
						{
						var rsGlossaryAcronyms =
							from dsGlossaryAcronym in this.SDDPdatacontext.GlossaryAndAcronyms
							where dsGlossaryAcronym.Id > intLastReadID1
							&& dsGlossaryAcronym.Modified > dtLastRefreshOn1
							select dsGlossaryAcronym;

						intEntriesCounter1 = 0;

						foreach(GlossaryAndAcronymsItem record in rsGlossaryAcronyms)
							{
							intEntriesCounter1 += 1;
							GlossaryAcronym objGlossaryAcronym;
							if(this.dsGlossaryAcronyms.TryGetValue(key: record.Id, value: out objGlossaryAcronym))
								dsGlossaryAcronyms.Remove(key: record.Id);
							else
								objGlossaryAcronym = new GlossaryAcronym();

							intLastReadID1 = record.Id;
							objGlossaryAcronym.ID = record.Id;
							objGlossaryAcronym.Term = record.Title;
							objGlossaryAcronym.Acronym = record.Acronym;
							objGlossaryAcronym.Meaning = record.Definition;
							objGlossaryAcronym.Modified = record.Modified;

							dsGlossaryAcronyms.Add(key: record.Id, value: objGlossaryAcronym);
							}
						if(intEntriesCounter1 < 1000)
							break;
						}
					objStopWatch1.Stop();
					Console.Write("\n\t + T1 - Glossary & Acronyms...\t\t {0} \t {1}", this.dsGlossaryAcronyms.Count, objStopWatch1.Elapsed);
					//- --------------------------
					// Populate **JobRoles**
					//- --------------------------
					intLastReadID1 = 0;
					objStopWatch1.Restart();
					bFetchMore1 = true;

					var dsJobFrameworks = this.SDDPdatacontext.JobFrameworkAlignment
						.Expand(jf => jf.JobDeliveryDomain);

					dtLastRefreshOn1 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsJobroles == null)
						this.dsJobroles = new Dictionary<int, JobRole>();
					else
						dtLastRefreshOn1 = this.LastRefreshedOn;

					while(bFetchMore1)
						{
						var rsJobFrameworks =
							from dsJobFramework in dsJobFrameworks
							where dsJobFramework.Id > intLastReadID1
							&& dsJobFramework.Modified > dtLastRefreshOn1
							select dsJobFramework;

						intEntriesCounter1 = 0;

						foreach(JobFrameworkAlignmentItem record in rsJobFrameworks)
							{
							intEntriesCounter1 += 1;
							JobRole objJobRole;
							if(this.dsJobroles.TryGetValue(key: record.Id, value: out objJobRole))
								dsGlossaryAcronyms.Remove(key: record.Id);
							else
								objJobRole = new JobRole();

							intLastReadID1 = record.Id;
							objJobRole.ID = record.Id;
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
						if(intEntriesCounter1 < 1000)
							break;
						}
					objStopWatch1.Stop();
					Console.Write("\n\t + T1 - JobRoles...\t\t\t\t\t {0} \t {1}", this.dsJobroles.Count.ToString("D3"), objStopWatch1.Elapsed);

					//- --------------------------------------
					// Populate ** TechnologyProdcuts **
					//- --------------------------------------
					intLastReadID1 = 0;
					objStopWatch1.Restart();
					bFetchMore1 = true;

					var dsTechnologyProducts = this.SDDPdatacontext.TechnologyProducts
						.Expand(tp => tp.TechnologyCategory)
						.Expand(tp => tp.TechnologyVendor);

					dtLastRefreshOn1 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsTechnologyProducts == null)
						this.dsTechnologyProducts = new Dictionary<int, TechnologyProduct>();
					else
						dtLastRefreshOn1 = this.LastRefreshedOn;

					while(bFetchMore1)
						{
						var rsTechnologyProducts =
							from dsTechProduct in dsTechnologyProducts
							where dsTechProduct.Id > intLastReadID1
							&& dsTechProduct.Modified > dtLastRefreshOn1
							select dsTechProduct;

						intEntriesCounter1 = 0;

						foreach(TechnologyProductsItem record in rsTechnologyProducts)
							{
							intEntriesCounter1 += 1;
							TechnologyProduct objTechProduct;
							if(this.dsTechnologyProducts.TryGetValue(key: record.Id, value: out objTechProduct))
								this.dsTechnologyProducts.Remove(key: record.Id);
							else
								objTechProduct = new TechnologyProduct();

							objTechProduct.ID = record.Id;
							intLastReadID1 = record.Id;
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
						if(intEntriesCounter1 < 1000)
							break;
						}
					objStopWatch1.Stop();
					Console.Write("\n\t + T1 - TechnologyProducts...\t\t {0} \t {1}", this.dsTechnologyProducts.Count.ToString("D3"), objStopWatch1.Elapsed);
					stopwatchThread1.Stop();
					Console.Write("\t\t### - Thread 1 Ended... duration: {0}", stopwatchThread1.Elapsed);
					//- **Signal** the CountDownEvent that thread 1 ended.
					threadCountDown.Signal();
					//- The thread exits the method.
					return;
					} //- end Thread1 Lock

//---g
Thread2start:
				lock(lockThread2)
					{
					Console.Write("\n### + Thread 2 Start... ###");
					Stopwatch stopwatchThread2 = Stopwatch.StartNew();
					//- ---------------------------------------------------
					// Populate the ** Service Portfolios **
					//- ----------------------------------------------------
					int intEntriesCounter2 = 0;
					int intLastReadID2 = 0;
					bool bFetechmore2 = true;
					Stopwatch objStopWatch2 = Stopwatch.StartNew();

					DateTime dtLastRefreshOn2 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsPortfolios == null)
						this.dsPortfolios = new Dictionary<int, ServicePortfolio>();
					else
						dtLastRefreshOn2 = this.LastRefreshedOn;

					while(bFetechmore2)
						{
						var rsPortfolios =
							from dsPortfolio in this.SDDPdatacontext.ServicePortfolios
							where dsPortfolio.Id > intLastReadID2
							&& dsPortfolio.Modified > dtLastRefreshOn2
							select dsPortfolio;

						intEntriesCounter2 = 0;

						foreach(var recordPortfolio in rsPortfolios)
							{
							intEntriesCounter2 += 1;
							ServicePortfolio objPortfolio;
							if(this.dsPortfolios.TryGetValue(key: recordPortfolio.Id, value: out objPortfolio))
								this.dsTechnologyProducts.Remove(key: recordPortfolio.Id);
							else
								objPortfolio = new ServicePortfolio();

							objPortfolio.ID = recordPortfolio.Id;
							intLastReadID2 = recordPortfolio.Id;
							objPortfolio.Title = recordPortfolio.Title;
							objPortfolio.PortfolioType = recordPortfolio.PortfolioTypeValue;
							objPortfolio.ISDheading = recordPortfolio.ISDHeading;
							objPortfolio.ISDdescription = recordPortfolio.ISDDescription;
							objPortfolio.CSDheading = recordPortfolio.ContractHeading;
							objPortfolio.CSDdescription = recordPortfolio.CSDDescription;
							objPortfolio.SOWheading = recordPortfolio.ContractHeading;
							objPortfolio.SOWdescription = recordPortfolio.ContractDescription;

							this.dsPortfolios.Add(key: recordPortfolio.Id, value: objPortfolio);
							}
						if(intEntriesCounter2 < 1000)
							break;
						}
					objStopWatch2.Stop();
					Console.Write("\n\t + T2 - ServicePortfolios\t\t\t {0} \t {1}", this.dsPortfolios.Count.ToString("D3"), objStopWatch2.Elapsed);

					//- ------------------------------------
					// Populate ** Service Families **
					//- ------------------------------------
					intLastReadID2 = 0;
					objStopWatch2.Restart();
					bFetechmore2 = true;
					dtLastRefreshOn2 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsFamilies == null)
						this.dsFamilies = new Dictionary<int, ServiceFamily>();
					else
						dtLastRefreshOn2 = this.LastRefreshedOn;

					while(bFetechmore2)
						{
						var rsFamilies = from dsFamily in this.SDDPdatacontext.ServiceFamilies
									  where dsFamily.Id > intLastReadID2 && dsFamily.Modified > dtLastRefreshOn2
									  select dsFamily;

						intEntriesCounter2 = 0;

						foreach(var recordFamily in rsFamilies)
							{
							intEntriesCounter2 += 1;
							ServiceFamily objFamily;
							if(this.dsFamilies.TryGetValue(key: recordFamily.Id, value: out objFamily))
								this.dsFamilies.Remove(key: recordFamily.Id);
							else
								objFamily = new ServiceFamily();

							objFamily.ID = recordFamily.Id;
							intLastReadID2 = recordFamily.Id;
							objFamily.Title = recordFamily.Title;
							objFamily.ServicePortfolioID = recordFamily.Service_PortfolioId;
							objFamily.ISDheading = recordFamily.ISDHeading;
							objFamily.ISDdescription = recordFamily.ISDDescription;
							objFamily.CSDheading = recordFamily.ContractHeading;
							objFamily.CSDdescription = recordFamily.CSDDescription;
							objFamily.SOWheading = recordFamily.ContractHeading;
							objFamily.SOWdescription = recordFamily.ContractDescription;

							this.dsFamilies.Add(key: recordFamily.Id, value: objFamily);
							}
						if(intEntriesCounter2 < 1000)
							break;
						}
					objStopWatch2.Stop();
					Console.Write("\n\t + T2 - ServiceFamilies...\t\t\t {0} \t {1}", this.dsFamilies.Count.ToString("D3"), objStopWatch2.Elapsed);

					//- -------------------------------------
					// Populate ** Service Products **
					//- -------------------------------------
					intLastReadID2 = 0;
					objStopWatch2.Restart();
					bFetechmore2 = true;
					dtLastRefreshOn2 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsProducts == null)
						this.dsProducts = new Dictionary<int, ServiceProduct>();
					else
						dtLastRefreshOn2 = this.LastRefreshedOn;

					while(bFetechmore2)
						{
						var rsProducts =
							from dsProduct in this.SDDPdatacontext.ServiceProducts
							where dsProduct.Id > intLastReadID2
							&& dsProduct.Modified > dtLastRefreshOn2
							select dsProduct;

						intEntriesCounter2 = 0;

						foreach(var recordProduct in rsProducts)
							{
							intEntriesCounter2 += 1;
							ServiceProduct objProduct;
							if(this.dsProducts.TryGetValue(key: recordProduct.Id, value: out objProduct))
								this.dsProducts.Remove(key: recordProduct.Id);
							else
								objProduct = new ServiceProduct();

							objProduct.ID = recordProduct.Id;
							intLastReadID2 = recordProduct.Id;
							objProduct.Title = recordProduct.Title;
							objProduct.ServiceFamilyID = recordProduct.Service_PortfolioId;
							objProduct.ISDheading = recordProduct.ISDHeading;
							objProduct.ISDdescription = recordProduct.ISDDescription;
							objProduct.CSDheading = recordProduct.ContractHeading;
							objProduct.CSDdescription = recordProduct.CSDDescription;
							objProduct.SOWheading = recordProduct.ContractHeading;
							objProduct.SOWdescription = recordProduct.ContractDescription;
							objProduct.KeyClientBenefits = recordProduct.KeyClientBenefits;
							objProduct.KeyDDbenefits = recordProduct.KeyDDBenefits;
							objProduct.PlannedActivities = recordProduct.PlannedActivities;
							objProduct.PlannedActivityEffortDrivers = recordProduct.PlannedActivityEffortDrivers;
							objProduct.PlannedDeliverables = recordProduct.PlannedDeliverables;
							objProduct.PlannedElements = recordProduct.PlannedElements;
							objProduct.PlannedFeatures = recordProduct.PlannedFeatures;
							objProduct.PlannedMeetings = recordProduct.PlannedMeetings;
							objProduct.PlannedReports = recordProduct.PlannedReports;
							objProduct.PlannedServiceLevels = recordProduct.PlannedServiceLevels;

							this.dsProducts.Add(key: recordProduct.Id, value: objProduct);
							}
						if(intEntriesCounter2 < 1000)
							break;
						}
					objStopWatch2.Stop();
					Console.Write("\n\t + T2 - ServiceProducts...\t\t\t {0} \t {1}", this.dsProducts.Count.ToString("D3"), objStopWatch2.Elapsed);

					stopwatchThread2.Stop();
					Console.Write("\t\t### - Thread 2 Ended... duration: {0}", stopwatchThread2.Elapsed);
					//- **Signal** the CountDownEvent that thread 2 ended.
					threadCountDown.Signal();
					//- the tread exits the method
					return;
					} // end Lock(Thread2)

//---g
Thread7start:
				lock(lockThread7)
					{
					Console.Write("\n### + Thread 7 Start... ###");
					Stopwatch stopwatchThread7 = Stopwatch.StartNew();
					//- ---------------------------------------------------
					// Populate the ** Service Portfolios **
					//- ----------------------------------------------------
					int intEntriesCounter7 = 0;
					int intLastReadID7 = 0;
					bool bFetechmore7 = true;
					Stopwatch objStopWatch7 = Stopwatch.StartNew();

					//-- --------------------------------------------
					// Populate **Service Element**
					//-- --------------------------------------------
					intLastReadID7 = 0;
					objStopWatch7.Restart();
					bFetechmore7 = true;
					DateTime dtLastRefreshOn7 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsElements == null)
						this.dsElements = new Dictionary<int, ServiceElement>();
					else
						dtLastRefreshOn7 = this.LastRefreshedOn;

					while(bFetechmore7)
						{
						var rsElements = from dsElement in this.SDDPdatacontext.ServiceElements
									  where dsElement.Id > intLastReadID7
									  && dsElement.Modified > dtLastRefreshOn7
									  select dsElement;

						intEntriesCounter7 = 0;

						foreach(var recElement in rsElements)
							{
							intEntriesCounter7 += 1;
							ServiceElement objElement;
							if(this.dsElements.TryGetValue(key: recElement.Id, value: out objElement))
								this.dsElements.Remove(key: recElement.Id);
							else
								objElement = new ServiceElement();

							objElement.ID = recElement.Id;
							intLastReadID7 = recElement.Id;
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
						if(intEntriesCounter7 < 1000)
							break;
						}
					objStopWatch7.Stop();
					Console.Write("\n\t + T7 - ServiceElements...\t\t\t {0} \t {1}", this.dsElements.Count.ToString("D3"), objStopWatch7.Elapsed);

					//- ----------------------------------
					// Populate **Service Feature**
					//- -----------------------------------
					intLastReadID7 = 0;
					objStopWatch7.Restart();
					intEntriesCounter7 = 0;
					bFetechmore7 = true;
					dtLastRefreshOn7 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsFeatures == null)
						this.dsFeatures = new Dictionary<int, ServiceFeature>();
					else
						dtLastRefreshOn7 = this.LastRefreshedOn;

					while(bFetechmore7)
						{
						var rsFeatures = from dsFeature in this.SDDPdatacontext.ServiceFeatures
									  where dsFeature.Id > intLastReadID7
									  && dsFeature.Modified > dtLastRefreshOn7
									  select dsFeature;

						intEntriesCounter7 = 0;

						foreach(var recFeature in rsFeatures)
							{
							intEntriesCounter7 += 1;
							ServiceFeature objFeature;
							if(this.dsFeatures.TryGetValue(key: recFeature.Id, value: out objFeature))
								this.dsFeatures.Remove(key: recFeature.Id);
							else
								objFeature = new ServiceFeature();

							intLastReadID7 = recFeature.Id;
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
						if(intEntriesCounter7 < 1000)
							break;
						}
					objStopWatch7.Stop();
					Console.Write("\n\t + T7 - ServiceFeatures...\t\t\t {0} \t {1}", this.dsFeatures.Count.ToString("D3"), objStopWatch7.Elapsed);
					stopwatchThread7.Stop();
					Console.Write("\t\t### - Thread 7 Ended... duration: {0}", stopwatchThread7.Elapsed);
					//- **Signal** the CountDownEvent that thread 2 ended.
					threadCountDown.Signal();
					//- the tread exits the method
					return;
					} // end Lock(Thread7)

Thread3start:
//---g
				lock(lockThread3)
					{
					Console.Write("\n### + Thread 3 Start... ###");
					Stopwatch stopwatchThread3 = Stopwatch.StartNew();
					//- -----------------------------------
					// Populate ** Deliverables **
					//- -----------------------------------
					Stopwatch objStopWatch3 = Stopwatch.StartNew();
					int intLastReadID3 = 0;
					bool bFetchMore3 = true;

					DateTime dtLasRefreshOn3 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsDeliverables == null)
						this.dsDeliverables = new Dictionary<int, Deliverable>();
					else
						dtLasRefreshOn3 = this.LastRefreshedOn;

					var dsDeliverables = this.SDDPdatacontext.Deliverables
						.Expand(dlv => dlv.SupportingSystems)
						.Expand(dlv => dlv.GlossaryAndAcronyms)
						.Expand(dlv => dlv.Responsible_RACI)
						.Expand(dlv => dlv.Accountable_RACI)
						.Expand(dlv => dlv.Consulted_RACI)
						.Expand(dlv => dlv.Informed_RACI);

					while(bFetchMore3)
						{
						var rsDeliverables =
							from dsDeliverable in dsDeliverables
							where dsDeliverable.Id > intLastReadID3
							&& dsDeliverable.Modified > dtLasRefreshOn3
							select dsDeliverable;

						int intEntriesCounter3 = 0;

						foreach(DeliverablesItem recDeliverable in rsDeliverables)
							{
							intEntriesCounter3 += 1;
							Deliverable objDeliverable;
							if(this.dsDeliverables.TryGetValue(key: recDeliverable.Id, value: out objDeliverable))
								this.dsDeliverables.Remove(key: recDeliverable.Id);
							else
								objDeliverable = new Deliverable();

							intLastReadID3 = recDeliverable.Id;
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
						if(intEntriesCounter3 < 1000)
							break;
						}

					objStopWatch3.Stop();
					Console.Write("\n\t + T3 - Deliverables...\t\t\t\t {0} \t {1}", this.dsDeliverables.Count.ToString("D3"), objStopWatch3.Elapsed);
					stopwatchThread3.Stop();
					Console.Write("\t\t### - Thread 3 Ended... duration: {0}", stopwatchThread3.Elapsed);
					//- **Signal** the CountDownEvent that thread 3 ended.
					threadCountDown.Signal();
					//- the tread exits the method
					return;
					} // end Lock(objThread3)

Thread4start:
//---g
				lock(lockThread4)
					{
					Console.Write("\n### + Thread 4 Start... ###");
					Stopwatch stopwatchThread4 = Stopwatch.StartNew();
					//- ---------------------------------------------------------
					// Populate ** Element Deliverables **
					//- ---------------------------------------------------------
					Stopwatch objStopWatch4 = Stopwatch.StartNew();
					int intLastReadID4 = 0;
					int intEntriesCounter4 = 0;
					bool bFetchMore4 = true;

					DateTime dtLasRefreshOn4 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsElementDeliverables == null)
						this.dsElementDeliverables = new Dictionary<int, ElementDeliverable>();
					else
						dtLasRefreshOn4 = this.LastRefreshedOn;

					while(bFetchMore4)
						{
						var rsElementDeliverable =
							from dsElementDeliverable in this.SDDPdatacontext.ElementDeliverables
							where dsElementDeliverable.Id > intLastReadID4
							&& dsElementDeliverable.Modified > dtLasRefreshOn4
							select dsElementDeliverable;

						intEntriesCounter4 = 0;

						foreach(var recElementDeliverable in rsElementDeliverable)
							{
							ElementDeliverable objElementDeliverable;
							if(this.dsElementDeliverables.TryGetValue(key: recElementDeliverable.Id, value: out objElementDeliverable))
								this.dsElementDeliverables.Remove(key: recElementDeliverable.Id);
							else
								objElementDeliverable = new ElementDeliverable();

							intEntriesCounter4 += 1;
							intLastReadID4 = recElementDeliverable.Id;
							objElementDeliverable.ID = recElementDeliverable.Id;
							objElementDeliverable.Title = recElementDeliverable.Title;
							objElementDeliverable.AssociatedDeliverableID = recElementDeliverable.Deliverable_Id;
							objElementDeliverable.AssociatedElementID = recElementDeliverable.Service_ElementId;
							objElementDeliverable.Optionality = recElementDeliverable.OptionalityValue;

							this.dsElementDeliverables.Add(key: recElementDeliverable.Id, value: objElementDeliverable);
							}
						if(intEntriesCounter4 < 1000)
							break;
						}
					objStopWatch4.Stop();
					Console.Write("\n\t + T4 - ElementDeliverables...\t\t {0} \t {1}", this.dsElementDeliverables.Count.ToString("D3"), objStopWatch4.Elapsed);

					//- -------------------------------------------------
					// Populate ** Feature Deliverables **
					//- --------------------------------------------------
					objStopWatch4 = Stopwatch.StartNew();
					intLastReadID4 = 0;
					bFetchMore4 = true;

					dtLasRefreshOn4 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsFeatureDeliverables == null)
						this.dsFeatureDeliverables = new Dictionary<int, FeatureDeliverable>();
					else
						dtLasRefreshOn4 = this.LastRefreshedOn;

					while(bFetchMore4)
						{
						var rsFeatureDeliverable =
							from dsFeatureDeliverable in this.SDDPdatacontext.FeatureDeliverables
							where dsFeatureDeliverable.Id > intLastReadID4
							&& dsFeatureDeliverable.Modified > dtLasRefreshOn4
							select dsFeatureDeliverable;

						intEntriesCounter4 = 0;

						foreach(var recFeatureDeliverable in rsFeatureDeliverable)
							{
							FeatureDeliverable objFeatureDeliverable;
							if(this.dsFeatureDeliverables.TryGetValue(key: recFeatureDeliverable.Id, value: out objFeatureDeliverable))
								this.dsFeatureDeliverables.Remove(key: recFeatureDeliverable.Id);
							else
								objFeatureDeliverable = new FeatureDeliverable();

							intEntriesCounter4 += 1;
							intLastReadID4 = recFeatureDeliverable.Id;
							objFeatureDeliverable.ID = recFeatureDeliverable.Id;
							objFeatureDeliverable.Title = recFeatureDeliverable.Title;
							objFeatureDeliverable.AssociatedDeliverableID = recFeatureDeliverable.Deliverable_Id;
							objFeatureDeliverable.AssociatedFeatureID = recFeatureDeliverable.Service_FeatureId;
							objFeatureDeliverable.Optionality = recFeatureDeliverable.OptionalityValue;

							this.dsFeatureDeliverables.Add(key: recFeatureDeliverable.Id, value: objFeatureDeliverable);
							}
						if(intEntriesCounter4 < 1000)
							break;
						}
					objStopWatch4.Stop();
					Console.Write("\n\t + T4 - FeatureDeliverables...\t\t {0} \t {1}", this.dsFeatureDeliverables.Count.ToString("D3"), objStopWatch4.Elapsed);

					//- -----------------------------------------------------
					// Populate ** DeliverableTechnologies **
					//- -----------------------------------------------------
					objStopWatch4 = Stopwatch.StartNew();
					intLastReadID4 = 0;
					bFetchMore4 = true;

					dtLasRefreshOn4 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsDeliverableTechnologies == null)
						this.dsDeliverableTechnologies = new Dictionary<int, DeliverableTechnology>();
					else
						dtLasRefreshOn4 = this.LastRefreshedOn;

					while(bFetchMore4)
						{
						var rsDeliverableTechnologies =
							from dsDeliverableTechnology in this.SDDPdatacontext.DeliverableTechnologies
							where dsDeliverableTechnology.Id > intLastReadID4
							&& dsDeliverableTechnology.Modified > dtLasRefreshOn4
							select dsDeliverableTechnology;

						intEntriesCounter4 = 0;

						foreach(var recDeliverableTechnology in rsDeliverableTechnologies)
							{
							DeliverableTechnology objDeliverableTechnology;
							if(this.dsDeliverableTechnologies.TryGetValue(key: recDeliverableTechnology.Id, value: out objDeliverableTechnology))
								this.dsDeliverableTechnologies.Remove(key: recDeliverableTechnology.Id);
							else
								objDeliverableTechnology = new DeliverableTechnology();

							intEntriesCounter4 += 1;
							intLastReadID4 = recDeliverableTechnology.Id;
							objDeliverableTechnology.ID = recDeliverableTechnology.Id;
							objDeliverableTechnology.Title = recDeliverableTechnology.Title;
							objDeliverableTechnology.Considerations = recDeliverableTechnology.TechnologyConsiderations;
							objDeliverableTechnology.RoadmapStatus = recDeliverableTechnology.TechnologyRoadmapStatusValue;
							objDeliverableTechnology.DeliviverableID = recDeliverableTechnology.Deliverable_Id;
							objDeliverableTechnology.TechnologyProductID = recDeliverableTechnology.TechnologyProductsId;

							this.dsDeliverableTechnologies.Add(key: recDeliverableTechnology.Id, value: objDeliverableTechnology);
							}
						if(intEntriesCounter4 < 1000)
							break;
						}
					objStopWatch4.Stop();
					Console.Write("\n\t + T4 - DeliverableTechnologies...\t {0} \t {1}", this.dsDeliverableTechnologies.Count.ToString("D3"), objStopWatch4.Elapsed);
					stopwatchThread4.Stop();
					Console.Write("\t\t### - Thread 4 Ended... duration: {0}", stopwatchThread4.Elapsed);
					//- **Signal** the CountDownEvent that thread 1 ended.
					threadCountDown.Signal();
					//- the tread exits the method
					return;
					} // end Lock(objThread4)

Thread5start:
//---g
				lock(lockThread5)
					{
					Console.Write("\n### + Thread 5 Start... ###");
					Stopwatch stopwatchThread5 = Stopwatch.StartNew();
					//- ------------------------------------
					// Populate ** Activities **
					//- ------------------------------------
					Stopwatch objStopWatch5 = Stopwatch.StartNew();
					int intLastReadID5 = 0;
					int intEntriesCounter5 = 0;
					bool bFetchMore5 = true;

					DateTime dtLasRefreshOn5 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsActivities == null)
						this.dsActivities = new Dictionary<int, Activity>();
					else
						dtLasRefreshOn5 = this.LastRefreshedOn;

					var dsActivities = this.SDDPdatacontext.Activities
						.Expand(ac => ac.Activity_Category)
						.Expand(ac => ac.OLA_);

					while(bFetchMore5)
						{
						var rsActivities =
							from dsActivity in dsActivities
							where dsActivity.Id > intLastReadID5
							&& dsActivity.Modified > dtLasRefreshOn5
							select dsActivity;

						intEntriesCounter5 = 0;

						foreach(ActivitiesItem record in rsActivities)
							{
							Activity objActivity;
							if(this.dsActivities.TryGetValue(key: record.Id, value: out objActivity))
								this.dsActivities.Remove(key: record.Id);
							else
								objActivity = new Activity();

							intEntriesCounter5 += 1;
							intLastReadID5 = record.Id;
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
							objActivity.OwningEntity = record.OwningEntityValue;
							if(record.OLA_ != null)
								objActivity.OLA = record.OLA_.Title;
							objActivity.OLAvariations = record.OLAVariations;
							objActivity.Optionality = record.ActivityOptionalityValue;
							if(record.Accountable_RACI != null)
								{
								objActivity.RACI_AccountableID = new List<int?>();
								objActivity.RACI_AccountableID.Add(record.Accountable_RACIId);
								}
							if(record.Responsible_RACI != null && record.Responsible_RACI.Count() > 0)
								{
								objActivity.RACI_ResponsibleID = new List<int>();
								foreach(var entryJobRole in record.Responsible_RACI)
									{
									objActivity.RACI_ResponsibleID.Add(entryJobRole.Id);
									}
								}
							if(record.Consulted_RACI != null && record.Consulted_RACI.Count() > 0)
								{
								objActivity.RACI_ConsultedID = new List<int>();
								foreach(var entryJobRole in record.Consulted_RACI)
									{
									objActivity.RACI_ConsultedID.Add(record.Id);
									}
								}
							if(record.Informed_RACI != null && record.Informed_RACI.Count() > 0)
								{
								objActivity.RACI_InformedID = new List<int>();
								foreach(var entryJobRole in record.Informed_RACI)
									{
									objActivity.RACI_InformedID.Add(record.Id);
									}
								}

							this.dsActivities.Add(key: record.Id, value: objActivity);
							}
						if(intEntriesCounter5 < 1000)
							break;
						}
					objStopWatch5.Stop();
					Console.Write("\n\t + T5 - Activities...\t\t\t\t {0} \t {1}", this.dsActivities.Count.ToString("D3"), objStopWatch5.Elapsed);

					//- ---------------------------------------------
					// Populate ** DeliverableActivities **
					//- ---------------------------------------------
					objStopWatch5 = Stopwatch.StartNew();
					intLastReadID5 = 0;
					bFetchMore5 = true;

					dtLasRefreshOn5 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsDeliverableActivities == null)
						this.dsDeliverableActivities = new Dictionary<int, DeliverableActivity>();
					else
						dtLasRefreshOn5 = this.LastRefreshedOn;

					while(bFetchMore5)
						{
						var rsDeliverableActivities =
							from dsDeliverableActivity in this.SDDPdatacontext.DeliverableActivities
							where dsDeliverableActivity.Id > intLastReadID5
							&& dsDeliverableActivity.Modified > dtLasRefreshOn5
							select dsDeliverableActivity;

						intEntriesCounter5 = 0;

						foreach(var recDeliverableActivity in rsDeliverableActivities)
							{
							DeliverableActivity objDeliverableActivity;
							if(this.dsDeliverableActivities.TryGetValue(key: recDeliverableActivity.Id, value: out objDeliverableActivity))
								this.dsDeliverableActivities.Remove(key: recDeliverableActivity.Id);
							else
								objDeliverableActivity = new DeliverableActivity();

							intLastReadID5 = recDeliverableActivity.Id;
							intEntriesCounter5 += 1;
							objDeliverableActivity.ID = recDeliverableActivity.Id;
							objDeliverableActivity.Title = recDeliverableActivity.Title;
							objDeliverableActivity.Optionality = recDeliverableActivity.OptionalityValue;
							objDeliverableActivity.AssociatedActivityID = recDeliverableActivity.Activity_Id;
							objDeliverableActivity.AssociatedDeliverableID = recDeliverableActivity.Deliverable_Id;

							this.dsDeliverableActivities.Add(key: recDeliverableActivity.Id, value: objDeliverableActivity);
							}
						if(intEntriesCounter5 < 1000)
							break;
						}
					objStopWatch5.Stop();
					Console.Write("\n\t + T5 - DeliverableActivities\t\t {0} \t {1}", this.dsDeliverableActivities.Count.ToString("D3"), objStopWatch5.Elapsed);
					stopwatchThread5.Stop();
					Console.Write("\t\t### - Thread 5 Ended... duration: {0}", stopwatchThread5.Elapsed);
					//- **Signal** the CountDownEvent that thread 5 ended.
					threadCountDown.Signal();
					//- the tread exits the method
					return;
					} // end lock(objThreadLock5)

Thread6start:
//---g
				lock(lockThread6)
					{
					Console.Write("\n### + Thread 6 Start... ###");
					Stopwatch stopwatchThread6 = Stopwatch.StartNew();
					//- ---------------------------------
					// Populate ** ServiceLevels **
					//- ---------------------------------
					Stopwatch stopwatch6 = Stopwatch.StartNew();
					int intLastReadID6 = 0;
					int intEntriesCounter6 = 0;
					bool bFetchMore6 = true;

					DateTime dtLasRefreshOn6 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsServiceLevels == null)
						this.dsServiceLevels = new Dictionary<int, ServiceLevel>();
					else
						dtLasRefreshOn6 = this.LastRefreshedOn;

					var datasetServiceLevels = this.SDDPdatacontext.ServiceLevels
						.Expand(sl => sl.Service_Hour);

					while(bFetchMore6)
						{
						var rsServiceLevels =
							from dsServiceLevel in datasetServiceLevels
							where dsServiceLevel.Id > intLastReadID6
							&& dsServiceLevel.Modified > dtLasRefreshOn6
							select dsServiceLevel;

						intEntriesCounter6 = 0;

						foreach(ServiceLevelsItem record in rsServiceLevels)
							{
							ServiceLevel objServiceLevel;
							if(this.dsServiceLevels.TryGetValue(key: record.Id, value: out objServiceLevel))
								this.dsServiceLevels.Remove(key: record.Id);
							else
								objServiceLevel = new ServiceLevel();

							intEntriesCounter6 += 1;
							intLastReadID6 = record.Id;
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
								from dsThreshold in this.SDDPdatacontext.ServiceLevelTargets
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
							// --------------------------------------------
							// Load the Service Level Performance Targets
							// --------------------------------------------
							var dsTargets =
								from dsThreshold in this.SDDPdatacontext.ServiceLevelTargets
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

						if(intEntriesCounter6 < 1000)
							break;
						}
					stopwatch6.Stop();
					Console.Write("\n\t + T6 - ServiceLevels...\t\t\t {0} \t {1}", this.dsServiceLevels.Count.ToString("D3"), stopwatch6.Elapsed);

					// ---------------------------------------
					// Populate DeliverableServiceLevels
					stopwatch6 = Stopwatch.StartNew();
					intLastReadID6 = 0;
					bFetchMore6 = true;

					dtLasRefreshOn6 = new DateTime(2000, 1, 1, 0, 0, 0);
					if(this.dsDeliverableServiceLevels == null)
						this.dsDeliverableServiceLevels = new Dictionary<int, DeliverableServiceLevel>();
					else
						dtLasRefreshOn6 = this.LastRefreshedOn;

					while(bFetchMore6)
						{
						var rsDeliverableServiceLevels =
							from dsDeliverableServiceLevel in this.SDDPdatacontext.DeliverableServiceLevels
							where dsDeliverableServiceLevel.Id > intLastReadID6
							&& dsDeliverableServiceLevel.Modified > dtLasRefreshOn6
							select dsDeliverableServiceLevel;

						intEntriesCounter6 = 0;

						foreach(var record in rsDeliverableServiceLevels)
							{
							DeliverableServiceLevel objDeliverableServiceLevel;
							if(this.dsDeliverableServiceLevels.TryGetValue(key: record.Id, value: out objDeliverableServiceLevel))
								this.dsDeliverableServiceLevels.Remove(key: record.Id);
							else
								objDeliverableServiceLevel = new DeliverableServiceLevel();

							intLastReadID6 = record.Id;
							intEntriesCounter6 += 1;
							objDeliverableServiceLevel.ID = record.Id;
							objDeliverableServiceLevel.Title = record.Title;
							objDeliverableServiceLevel.Optionality = record.OptionalityValue;
							objDeliverableServiceLevel.ContentStatus = record.ContentStatusValue;
							objDeliverableServiceLevel.AdditionalConditions = record.AdditionalConditions;
							objDeliverableServiceLevel.AssociatedDeliverableID = record.Deliverable_Id;
							objDeliverableServiceLevel.AssociatedServiceLevelID = record.Service_LevelId;
							objDeliverableServiceLevel.AssociatedServiceProductID = record.Service_ProductId;

							this.dsDeliverableServiceLevels.Add(key: record.Id, value: objDeliverableServiceLevel);
							}
						if(intEntriesCounter6 < 1000)
							break;
						}
					stopwatch6.Stop();
					Console.Write("\n\t + T6 - DeliverableServiceLevels...\t {0} \t {1}", this.dsDeliverableServiceLevels.Count.ToString("D3"), stopwatch6.Elapsed);
					stopwatchThread6.Stop();
					Console.Write("\t\t### - Thread 6 Ended... duration: {0}", stopwatchThread6.Elapsed);
					//- **Signal** the CountDownEvent that thread 6 ended.
					threadCountDown.Signal();
					//- the tread exits the method
					return;
					} // end lock(objThreadLock6)

//---g
ThreadSynchroStart:
				lock(lockThreadSynchro)
					{
					//- -----------------------------------------------------------------------------------------------------------------
					// **Monitor** the DataPopulation and wait for each thread to complete, before setting the:
					// - **RefreshingDateTimeStamp**
					// - **IsDataSetComplete**
					//- ------------------------------------------------------------------------------------------------------------------
					objStopWatchCompleteDataSet = Stopwatch.StartNew();
					threadCountDown.Wait();
					this.LastRefreshedOn = this.RefreshingDateTimeStamp;
					this.IsDataSetComplete = true;
					} // end lock(objThreadSychro)
				objStopWatchCompleteDataSet.Stop();
				Console.Write("\n\nPopulating the complete DataSet took {0}", objStopWatchCompleteDataSet.Elapsed);
				//- The main thread does not terminate, it returns to continue with execution...
				return;
				}
			catch(DataServiceClientException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nStatusCode: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.StatusCode, exc.StackTrace);
				this.IsDataSetComplete = false;
				}
			catch(DataServiceQueryException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nResponse: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				this.IsDataSetComplete = false;
				}
			catch(DataServiceTransportException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nResponse:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				this.IsDataSetComplete = false;
				}
			catch(System.Net.Sockets.SocketException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nTargetSite:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
				this.IsDataSetComplete = false;
				}
			catch(Exception exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nSource:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Source, exc.StackTrace);
				this.IsDataSetComplete = false;
				}
			}

		//++ PopulateMappingDataSet
		public bool PopulateMappingDataset(DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
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

				Console.Write("\t\t\t\t {0} \t {1}", this.dsMappings.Count.ToString("D3"), DateTime.Now - setStart);

				//+ Populate Mapping Service Towers
				Console.Write("\n\t + MappingServiceTowers...");
				setStart = DateTime.Now;
				intLastReadID = 0;
				this.dsMappingServiceTowers = new Dictionary<int, MappingServiceTower>();
				do
					{
					var rsMappingServiceTowers = 
						from dsMappingServiceTowers in parDatacontexSDDP.MappingServiceTowers
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
				Console.Write("\t {0} \t {1}", this.dsMappingServiceTowers.Count.ToString("D3"), DateTime.Now - setStart);

				//+ Populate Mapping Requirements
				Console.Write("\n\t + MappingRequirements...");
				setStart = DateTime.Now;
				this.dsMappingRequirements = new Dictionary<int, MappingRequirement>();
				// Populate the Mapping Requirements
				if(this.dsMappingServiceTowers != null && this.dsMappingServiceTowers.Count > 0)
					{
					foreach(var itemServiceTower in this.dsMappingServiceTowers)
						{
						var rsMappingRequirements =
							from dsMappingRequirements in parDatacontexSDDP.MappingRequirements
							where dsMappingRequirements.Mapping_TowerId == itemServiceTower.Key
							select dsMappingRequirements;

						foreach(var recordMR in rsMappingRequirements)
							{
							MappingRequirement objMappingRequirement = new MappingRequirement();
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
						}
					}
				Console.Write("\t {0} \t {1}", this.dsMappingRequirements.Count.ToString("D3"), DateTime.Now - setStart);

				//+ Populate Mapping Assumptions
				Console.Write("\n\t + MappingAssumptions...");
				setStart = DateTime.Now;
				this.dsMappingAssumptions = new Dictionary<int, MappingAssumption>();

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
						} // foreach(var itemRequirement in this.dsMappingRequirements)
					} // if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
				Console.Write("\t {0} \t {1}", this.dsMappingAssumptions.Count.ToString("D3"), DateTime.Now - setStart);

				//+ Populate Mapping Risks...
				Console.Write("\n\t + MappingRisks...");
				setStart = DateTime.Now;
				this.dsMappingRisks = new Dictionary<int, MappingRisk>();
				// Populate the Mapping Requirements
				if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
					{
					foreach(var itemRequirement in this.dsMappingRequirements)
						{
						var rsMappingAssumptions =
							from dsMappingAssumptions in parDatacontexSDDP.MappingAssumptions
							where dsMappingAssumptions.Mapping_RequirementId == itemRequirement.Key
							select dsMappingAssumptions;

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
						} // foreach(var itemRequirement in this.dsMappingRequirements)
					} // if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
				Console.Write("\t\t\t {0} \t {1}", this.dsMappingRisks.Count.ToString("D3"), DateTime.Now - setStart);

				// ---------------------------------------------
				//+ Populate Mapping Deliverables...
				Console.Write("\n\t + Mapping Deliverables...");
				setStart = DateTime.Now;
				this.dsMappingDeliverables = new Dictionary<int, MappingDeliverable>();

				// Populate the Mapping Deliverables
				if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
					{
					foreach(var itemRequirement in this.dsMappingRequirements)
						{
						// Populate the Maping Deliverables..
						var rsMappingDeliverables =
							from dsMappingDeliverable in parDatacontexSDDP.MappingDeliverables
							where dsMappingDeliverable.Mapping_RequirementId == itemRequirement.Key
							select dsMappingDeliverable;

						foreach(var recordMappingDeliverable in rsMappingDeliverables)
							{
							MappingDeliverable objMappingDeliverable = new MappingDeliverable();
							objMappingDeliverable.ID = recordMappingDeliverable.Id;
							objMappingDeliverable.MappingRequirementID = recordMappingDeliverable.Mapping_RequirementId;
							objMappingDeliverable.Title = recordMappingDeliverable.Title;
							if(recordMappingDeliverable.DeliverableChoiceValue == "New")
								objMappingDeliverable.NewDeliverable = true;
							else
								objMappingDeliverable.NewDeliverable = false;
							objMappingDeliverable.MappedDeliverableID = recordMappingDeliverable.Mapped_DeliverableId;
							objMappingDeliverable.NewRequirement = recordMappingDeliverable.DeliverableRequirement;
							objMappingDeliverable.ComplianceComments = recordMappingDeliverable.ComplianceComments;
							this.dsMappingDeliverables.Add(key: recordMappingDeliverable.Id, value: objMappingDeliverable);
							} //foreach(var recordDeliv in rsMappingDeliverable)
						} // foreach(var itemRequirement in this.dsMappingRequirements)
					} // if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
				Console.Write("\t {0} \t {1}", this.dsMappingDeliverables.Count.ToString("D3"), DateTime.Now - setStart);

				//+ Populate Mapping Service Levels
				Console.Write("\n\t + MappingServiceLevels");
				setStart = DateTime.Now;
				this.dsMappingServiceLevels = new Dictionary<int, MappingServiceLevel>();

				// Populate the Mapping Service Levels
				if(this.dsMappingDeliverables != null && this.dsMappingServiceLevels.Count > 0)
					{
					foreach(var itemMappingDeliverable in this.dsMappingDeliverables)
						{
						// Populate the Mapping Service Levels
						var rsMappingServiceLevels =
						from dsMappingServiceLevel in parDatacontexSDDP.MappingServiceLevels
						where dsMappingServiceLevel.Mapping_DeliverableId == itemMappingDeliverable.Key
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
							} // foreach(var recordMSL in rsMappingServiceLevels)
						} // foreach(var itemMappingDeliverable in this.dsMappingDeliverables)
					} // if(this.dsMappingRequirements != null && this.dsMappingRequirements.Count > 0)
				Console.Write("\t\t {0} \t {1}", this.dsMappingServiceLevels.Count.ToString("D3"), DateTime.Now - setStart);

				Console.Write("\n\n\tPopulating the Mappings DataSet ended at {0} and took {1}.", DateTime.Now, DateTime.Now - startTime);
				this.IsDataSetComplete = true;
				return true;
				}
			catch(DataServiceClientException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} - StatusCode:{2}\n{3}.", exc.HResult, exc.Message, exc.StatusCode, exc.StackTrace);
				this.IsDataSetComplete = true;
				return false;
				}
			catch(DataServiceQueryException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} - StatusCode:{2}\n{3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				return false;
				}
			catch(DataServiceTransportException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} \n{3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				return false;
				}
			}
		}
	}