using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{

	class ServicePortfolio
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

		// ----------------------------
		// Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parPopulateNextLevel = false,
			bool parPopulateAllLevels = false)
			{
			try
				{
				// Access the Service Portfolios List
				var rsPortfolios =
					from dsPortfolio in parDatacontexSDDP.ServicePortfolios
					where dsPortfolio.Id == parID
					select dsPortfolio;

				ServicePortfoliosItem recPortfolio = rsPortfolios.FirstOrDefault();

				if(recPortfolio == null) // Service Portfolio was not found
					{
					this.ID = 0;
					}
				else
					{
					this.ID = recPortfolio.Id;
					this.Title = recPortfolio.Title;
					this.ISDheading = recPortfolio.ISDHeading;
					this.ISDdescription = recPortfolio.ISDDescription;
					this.CSDheading = recPortfolio.CSDHeading;
					this.CSDdescription = recPortfolio.CSDDescription;
					this.SOWheading = recPortfolio.ContractHeading;
					this.SOWdescription = recPortfolio.ContractDescription;
					}
				} // try
			
			catch(DataServiceQueryException exc)
				{
				throw new DataServiceQueryException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return true;
			}
		}

	class ServiceFamily
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

		// ----------------------------
		// Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP, int? parID)
			{
			try
				{
				// Access the Service Families List
				var rsFamilies =
					from dsFamilies in parDatacontexSDDP.ServiceFamilies
					where dsFamilies.Id == parID
					select dsFamilies;

				var recFamily = rsFamilies.FirstOrDefault();
				if(recFamily == null) // Service Family was not found
					{
					// throw new DataEntryNotFoundException("Service Family content for ID:" + parID + " could not be found in SharePoint.");
					this.ID = 0;
					}
				else
					{
					this.ID = recFamily.Id;
					this.ServicePortfolioID = recFamily.Service_PortfolioId;
					this.Title = recFamily.Title;
					this.ISDheading = recFamily.ISDHeading;
					this.ISDdescription = recFamily.ISDDescription;
					this.CSDheading = recFamily.CSDHeading;
					this.CSDdescription = recFamily.CSDDescription;
					this.SOWheading = recFamily.ContractHeading;
					this.SOWdescription = recFamily.ContractDescription;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			}

		} // end of class ServicePFamily

	///##################################################
	/// <summary>
	/// Service Product object represent an entry in the Service Products SharePoint List
	/// </summary>
	class ServiceProduct
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

		// ----------------------------
		// Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				// Access the Service Products List
				var rsProducts =
					from dsProduct in parDatacontexSDDP.ServiceProducts
					where dsProduct.Id == parID
					select dsProduct;

				var recProduct = rsProducts.FirstOrDefault();
				if(recProduct == null) // Service Product was not found
					{
					//throw new DataEntryNotFoundException("Service Product content for ID:" + parID + " could not be found in SharePoint.");
					this.ID = 0;
					}
				else
					{
					this.ID = recProduct.Id;
					this.ServiceFamilyID = recProduct.Service_FamilyId;
					this.Title = recProduct.Title;
					this.ISDheading = recProduct.ISDHeading;
					this.ISDdescription = recProduct.ISDDescription;
					this.KeyClientBenefits = recProduct.KeyClientBenefits;
					this.KeyDDbenefits = recProduct.KeyDDBenefits;
					this.CSDheading = recProduct.CSDHeading;
					this.CSDdescription = recProduct.CSDDescription;
					this.SOWheading = recProduct.ContractHeading;
					this.SOWdescription = recProduct.ContractDescription;
					this.PlannedActivities = recProduct.PlannedActivities;
					this.PlannedActivityEffortDrivers = recProduct.PlannedActivityEffortDrivers;
					this.PlannedDeliverables = recProduct.PlannedDeliverables;
					this.PlannedElements = recProduct.PlannedElements;
					this.PlannedFeatures = recProduct.PlannedFeatures;
					this.PlannedMeetings = recProduct.PlannedMeetings;
					this.PlannedReports = recProduct.PlannedReports;
					this.PlannedServiceLevels = recProduct.PlannedServiceLevels;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			}

		} // end of class ServiceProduct

	///############################################
	/// <summary>
	/// This object represents an entry in the Service Elements SharePoint List
	/// </summary>
	class ServiceElement
		{
		public int ID{get; set;}
		public int? ServiceProductID{get; set;}
		public string Title{get; set;}
		public double? SortOrder{get; set;}public string ISDheading{get; set;}
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
		public ServiceElement Layer1up{get; set;}
		// ----------------------------
		// Service Element Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID, bool parGetLayer1up = false)
			{
			try
				{
				// Access the Service Elements List
				var rsElements =
					from dsElement in parDatacontexSDDP.ServiceElements
					where dsElement.Id == parID
					select dsElement;

				var recElement = rsElements.FirstOrDefault();
				if(recElement == null) // Service Element was not found
					{
					//throw new DataEntryNotFoundException("Service Element content for ID:" + parID + " could not be found in SharePoint.");
					this.ID = 0;
					return false;
					}
				else
					{
					this.ID = recElement.Id;
					this.ServiceProductID = recElement.Service_ProductId;
					this.Title = recElement.Title;
					this.SortOrder = recElement.SortOrder;
					this.ISDheading = recElement.ISDHeading;
					this.ISDdescription = recElement.ISDDescription;
					this.Objectives = recElement.Objective;
					this.KeyClientAdvantages = recElement.KeyClientAdvantages;
					this.KeyClientBenefits = recElement.KeyClientBenefits;
					this.KeyDDbenefits = recElement.KeyDDBenefits;
					this.KeyPerformanceIndicators = recElement.KeyPerformanceIndicators;
					this.CriticalSuccessFactors = recElement.CriticalSuccessFactors;
					this.ProcessLink = recElement.ProcessLink;
					this.ContentStatus = recElement.ContentStatusValue;
					//this.ContentLayerValue = this.ContentLayerValue;
					this.ContentLayerValue = recElement.ContentLayerValue;
					this.ContentPredecessorElementID = recElement.ContentPredecessorElementId;
					if(parGetLayer1up == true && recElement.ContentPredecessorElementId != null)
						{
						ServiceElement objServiceElementlayer1up = new ServiceElement();
						try
							{
							objServiceElementlayer1up.PopulateObject(
								parDatacontexSDDP: parDatacontexSDDP,
								parID: recElement.ContentPredecessorElementId,
								parGetLayer1up: true);

							this.Layer1up = objServiceElementlayer1up;
							}
						catch(DataEntryNotFoundException)
							{
							this.Layer1up = null;
							}
						}
					else
						{
						this.Layer1up = null;
						}
					} //if(recElement != null) // Service Element was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			} // end Populate method

		///----------------------------------------------
		/// Obtain a List of Service Element Objects 
		/// ---------------------------------------------
		/// <summary>
		/// Obtain a list containing all the Service Element objects associated with the value in the parServiceProductID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP">an SDDP data connection.</param>
		/// <param name="parServiceProductID">The Service Porduct ID for which the list must be populated.</param>
		/// <param name="parGetContentLayers">ehrn TRUE, the content layers are also Populated, else no content layers are fetched. The optional parameter value is TRUE, </param>
		/// <returns></returns>
		public static List<ServiceElement> ObtainListOfObjects(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parServiceProductID,
			bool parGetContentLayers = true)
			{
			List<ServiceElement> listServiceElements = new List<ServiceElement>();

			try
				{
				// Access the ServiceElements List
				var rsServiceElements =
					from dsServiceElements in parDatacontextSDDP.ServiceElements
					where dsServiceElements.Service_ProductId == parServiceProductID
					orderby dsServiceElements.SortOrder
					select new
						{
						dsServiceElements.Id,
						dsServiceElements.Title,
						dsServiceElements.SortOrder,
						dsServiceElements.ISDHeading,
						dsServiceElements.ISDDescription,
						dsServiceElements.KeyClientAdvantages,
						dsServiceElements.KeyClientBenefits,
						dsServiceElements.KeyDDBenefits,
						dsServiceElements.KeyPerformanceIndicators,
						dsServiceElements.CriticalSuccessFactors,
						dsServiceElements.Objective,
						dsServiceElements.ContentLayerValue,
						dsServiceElements.ContentPredecessorElementId,
						dsServiceElements.ProcessLink
						};

				if(rsServiceElements.Count() == 0) // MappingTowers was not found
					{
					//throw new DataEntryNotFoundException("No Mapping Tower entries for Mapping ID:" +
					//	parMappingID + " could be found in SharePoint.");
					return listServiceElements;
					}

				foreach(var record in rsServiceElements)
					{
					ServiceElement objServiceElement = new ServiceElement();
					objServiceElement.ID = record.Id;
					objServiceElement.Title = record.Title;
					objServiceElement.SortOrder = record.SortOrder;
					objServiceElement.ISDheading = record.ISDHeading;
					objServiceElement.ISDdescription = record.ISDDescription;
					objServiceElement.KeyClientAdvantages = record.KeyClientAdvantages;
					objServiceElement.KeyClientBenefits = record.KeyClientBenefits;
					objServiceElement.KeyDDbenefits = record.KeyDDBenefits;
					objServiceElement.KeyPerformanceIndicators = record.KeyPerformanceIndicators;
					objServiceElement.CriticalSuccessFactors = record.CriticalSuccessFactors;
					objServiceElement.Objectives = record.Objective;
					objServiceElement.ContentLayerValue = record.ContentLayerValue;
					objServiceElement.ContentPredecessorElementID = record.ContentPredecessorElementId;
					objServiceElement.ProcessLink = record.ProcessLink;
					if(objServiceElement.ContentPredecessorElementID != null
						&& parGetContentLayers == true)
						{
						ServiceElement objLayer1up = new ServiceElement();
						objLayer1up.PopulateObject(parDatacontextSDDP, objServiceElement.ContentPredecessorElementID, parGetContentLayers);
						objServiceElement.Layer1up = objLayer1up;
						}
					else
						{
						objServiceElement.Layer1up = null;
						}
					listServiceElements.Add(objServiceElement);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listServiceElements;
			} // end if ObtainListOfObjects

		} // end Class ServiceElement

	///##############################################################
	///#### Service Feature Object
	///##############################################################
	/// <summary>
	/// This object represents an entry in the Service Features SharePoint List.
	/// </summary>
	class ServiceFeature
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
		public ServiceFeature Layer1up{get; set;}
		public string ContentStatus{get; set;}

		// ----------------------------
		// Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parGetLayer1up = false)
			{
			try
				{
				// Access the Service Features List
				var rsFeatures =
					from dsFeature in parDatacontexSDDP.ServiceFeatures
					where dsFeature.Id == parID
					select new
						{
						dsFeature.Id,
						dsFeature.Title,
						dsFeature.SortOrder,
						dsFeature.CSDHeading,
						dsFeature.CSDDescription,
						dsFeature.ContractHeading,
						dsFeature.ContractDescription,
						dsFeature.ContentLayerValue,
						dsFeature.ContentPredecessorFeatureId,
						dsFeature.ContentStatusValue
						};

				var recFeature = rsFeatures.FirstOrDefault();
				if(recFeature == null) // Service Feature was not found
					{
					//throw new DataEntryNotFoundException("Service Feature content for ID:" +
					//	parID + " could not be found in SharePoint.");
					this.ID = 0;
					return false;
					}
				else
					{
					this.ID = recFeature.Id;
					this.Title = recFeature.Title;
					this.SortOrder = recFeature.SortOrder;
					this.CSDheading = recFeature.CSDHeading;
					this.CSDdescription = recFeature.CSDDescription;
					this.SOWheading = recFeature.ContractHeading;
					this.SOWdescription = recFeature.ContractDescription;
					this.ContentStatus = recFeature.ContentStatusValue;

					//this.ContentLayerValue = this.ContentLayerValue;
					this.ContentLayerValue = recFeature.ContentLayerValue;
					this.ContentPredecessorFeatureID = recFeature.ContentPredecessorFeatureId;
					if(parGetLayer1up == true
					&& recFeature.ContentPredecessorFeatureId != null)
						{
						ServiceFeature objServiceFeatureLayer1up = new ServiceFeature();
						try
							{
							objServiceFeatureLayer1up.PopulateObject(
								parDatacontexSDDP: parDatacontexSDDP,
								parID: recFeature.ContentPredecessorFeatureId,
								parGetLayer1up: true);

							this.Layer1up = objServiceFeatureLayer1up;
							}
						catch(DataEntryNotFoundException)
							{
							this.Layer1up = null;
							}
						}
					else
						{
						this.Layer1up = null;
						}
					} //if(recFeature != null) // Service Feature was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			} // end Populate method

		///----------------------------------------------
		/// <summary>
		/// Obtain a list containing all the Service Feature objects associated with the value in the parServiceProductID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parServiceProductID"></param>
		/// <param name="parGetContentLayers">Optional parameter which determines whether related content layers are also obtained. Default is TRUE</param>
		/// <returns></returns>
		public static List<ServiceFeature> ObtainListOfObjects(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parServiceProductID,
			bool parGetContentLayers = true)
			{
			List<ServiceFeature> listServiceFeature = new List<ServiceFeature>();

			try
				{
				// Access the ServiceElements List
				var rsServiceFeatures =
					from dsServiceServiceFeature in parDatacontextSDDP.ServiceFeatures
					where dsServiceServiceFeature.Service_ProductId == parServiceProductID
					orderby dsServiceServiceFeature.SortOrder
					select new
						{
						dsServiceServiceFeature.Id,
						dsServiceServiceFeature.Title,
						dsServiceServiceFeature.SortOrder,
						dsServiceServiceFeature.CSDHeading,
						dsServiceServiceFeature.CSDDescription,
						dsServiceServiceFeature.ContentLayerValue,
						dsServiceServiceFeature.ContentPredecessorFeatureId
						};

				if(rsServiceFeatures.Count() == 0) // No Service Features were found
					{
					return listServiceFeature;
					}

				foreach(var record in rsServiceFeatures)
					{
					ServiceFeature objServiceFeature = new ServiceFeature();
					objServiceFeature.ID = record.Id;
					objServiceFeature.Title = record.Title;
					objServiceFeature.SortOrder = record.SortOrder;
					objServiceFeature.CSDheading = record.CSDHeading;
					objServiceFeature.CSDdescription = record.CSDDescription;
					objServiceFeature.ContentLayerValue = record.ContentLayerValue;
					objServiceFeature.ContentPredecessorFeatureID = record.ContentPredecessorFeatureId;

					if(objServiceFeature.ContentPredecessorFeatureID != null
					&& parGetContentLayers == true)
						{
						ServiceFeature objLayer1up = new ServiceFeature();
						objLayer1up.PopulateObject(parDatacontextSDDP, objServiceFeature.ContentPredecessorFeatureID, parGetContentLayers);
						objServiceFeature.Layer1up = objLayer1up;
						}
					else
						{
						objServiceFeature.Layer1up = null;
						}
					listServiceFeature.Add(objServiceFeature);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listServiceFeature;
			} // end if ObtainListOfObjects


		} // end Class ServiceFeature

	/// #############################################
	/// ### Deliverables Object
	/// #############################################
	/// <summary>
	/// This object represent an entry in the Deliverables SharePoint List.
	/// </summary>
	class Deliverable
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
		public List<string> SupportingSystems {get; set;}
		public string WhatHasChanged{get; set;}
		public string ContentLayerValue{get; set;}
		public string ContentStatus{get; set;}
		public Dictionary<int, string> GlossaryAndAcronyms{get; set;}
		public int? ContentPredecessorDeliverableID{get; set;}
		public Deliverable Layer1up{get; set;}
		public List<int?> RACIaccountables{get; set;}
		public List<int?> RACIresponsibles{get; set;}
		public List<int?> RACIinformeds{get; set;}
		public List<int?> RACIconsulteds{get; set;}

		// ----------------------------------------------
		// Deliverable - Populate method
		// ----------------------------------------------
		/// <summary>
		/// 
		/// </summary>
		/// <param name="parDatacontexSDDP"></param>
		/// <param name="parID"></param>
		/// <param name="parGetLayer1up"></param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parGetLayer1up = false,
			bool parGetRACI = false)
			{
			try
				{
				// Access the Deliverables List
				var dsDeliverables = parDatacontexSDDP.Deliverables
					.Expand(dlv => dlv.SupportingSystems)
					.Expand(dlv => dlv.GlossaryAndAcronyms)
					.Expand(dlv => dlv.Responsible_RACI)
					.Expand(dlv => dlv.Accountable_RACI)
					.Expand(dlv => dlv.Consulted_RACI)
					.Expand(dlv => dlv.Informed_RACI);

				var rsDeliverables =
					from dsDeliverable in dsDeliverables
					where dsDeliverable.Id == parID
					select dsDeliverable;

				var recDeliverable = rsDeliverables.FirstOrDefault();
				if(recDeliverable == null) // Service Element was not found
					{
					//throw new DataEntryNotFoundException("Content for Deliverable ID:" + parID + " could not be found in SharePoint.");
					this.ID = 0;
					}
				else
					{
					this.ID = recDeliverable.Id;
					this.Title = recDeliverable.Title;
					this.SortOrder = recDeliverable.SortOrder;
					this.ISDheading = recDeliverable.ISDHeading;
					this.ISDsummary = recDeliverable.ISDSummary;
					this.ISDdescription = recDeliverable.ISDDescription;
					this.CSDheading = recDeliverable.CSDHeading;
					this.CSDsummary = recDeliverable.CSDSummary;
					this.CSDdescription = recDeliverable.CSDDescription;
					this.SoWheading = recDeliverable.ContractHeading;
					this.SoWsummary = recDeliverable.ContractSummary;
					this.SoWdescription = recDeliverable.ContractDescription;
					this.TransitionDescription = recDeliverable.TransitionDescription;
					this.Inputs = recDeliverable.Inputs;
					this.Outputs = recDeliverable.Outputs;
					this.DDobligations = recDeliverable.SPObligations;
					this.ClientResponsibilities = recDeliverable.ClientResponsibilities;
					this.Exclusions = recDeliverable.Exclusions;
					this.GovernanceControls = recDeliverable.GovernanceControls;
					this.WhatHasChanged = recDeliverable.WhatHasChanged;
					this.ContentStatus = recDeliverable.ContentStatusValue;
					this.ContentLayerValue = recDeliverable.ContentLayerValue;
					this.ContentPredecessorDeliverableID = recDeliverable.ContentPredecessor_DeliverableId;

					// Add the Glossary and Acronym terms to the Deliverable object
					if(recDeliverable.GlossaryAndAcronyms.Count > 0)
						{
						foreach(var entry in recDeliverable.GlossaryAndAcronyms)
							{
							if(this.GlossaryAndAcronyms == null)
								{
								this.GlossaryAndAcronyms = new Dictionary<int, string>();
								}
							if(this.GlossaryAndAcronyms.ContainsKey(entry.Id) != true)
								this.GlossaryAndAcronyms.Add(entry.Id, entry.Title);
							}
						}

					if(recDeliverable.SupportingSystems != null)
						{
						this.SupportingSystems = new List<string>();
						foreach(var systemItem in recDeliverable.SupportingSystems)
							{
							this.SupportingSystems.Add(systemItem.Value);
							}
						}

					//Only poulate the RACI tables if required
					if(parGetRACI)
						{
						//Populate the RACI dictionaries
						// --- RACIresponsibles
						if(recDeliverable.Responsible_RACI.Count > 0)
							{
							RACIresponsibles = new List<int?>();
							foreach(var entry in recDeliverable.Responsible_RACI)
								{
								RACIresponsibles.Add(entry.Id);
								}
							}

						// --- RACIaccountables
						RACIaccountables = new List<int?>();
						if(recDeliverable.Accountable_RACI != null)
							{
							RACIaccountables.Add(recDeliverable.Accountable_RACIId);
							}

						// --- RACIconsulteds
						if(recDeliverable.Consulted_RACI.Count > 0)
							{
							RACIconsulteds = new List<int?>();
							foreach(var entry in recDeliverable.Consulted_RACI)
								{
								RACIconsulteds.Add(entry.Id);
								}
							}

						// --- RACIinformeds
						if(recDeliverable.Informed_RACI.Count > 0)
							{
							RACIinformeds = new List<int?>();
							foreach(var entry in recDeliverable.Informed_RACI)
								{

								RACIinformeds.Add(entry.Id);
								}
							}
						}

					// Add the recursive relationship of Content Predecessor if required
					if(parGetLayer1up == true
					&& recDeliverable.ContentPredecessor_DeliverableId != null)
						{
						Deliverable objDeliverableLayer1up = new Deliverable();
						try
							{
							objDeliverableLayer1up.PopulateObject(
								parDatacontexSDDP: parDatacontexSDDP,
								parID: recDeliverable.ContentPredecessor_DeliverableId,
								parGetLayer1up: true);

							this.Layer1up = objDeliverableLayer1up;
							}
						catch(DataEntryNotFoundException)
							{
							this.Layer1up = null;
							}
						catch(Exception exc)
							{
							Console.WriteLine("Exception consumed: {0} - {1}", exc.HResult, exc.Message);
							this.Layer1up = null;
							}
						}
					else
						{
						this.Layer1up = null;
						}
					} //if(recDeliverable != null) // Deliverable was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint. Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			} // end of Method PopulateObject


		} // end Class Deliverables

	// ####################################################
	// ### Deliverable Service Levels class
	// ####################################################
	/// <summary>
	/// 
	/// </summary>
	class DeliverableServiceLevel
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

		// ----------------------------------------
		// DeliverableServiceLevel Populate Method
		//-----------------------------------------
		public bool Populate(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parGetLayer1up = false,
			bool parPopulateServiceLevelObject = false,
			bool parPopulateDeliverableObject = false,
			bool parPopulateServiceProductObject = false)
			{
			try
				{
				// Access the DeliverableServiceLevels List
				var dsDeliverableServiceLevels = parDatacontexSDDP.DeliverableServiceLevels
					.Expand(delSL => delSL.Deliverable_)
					.Expand(delSL => delSL.Service_Level)
					.Expand(delSL => delSL.Service_Product);

				var rsDeliverableServiceLevels =
					from dsDeliverableServiceLevel in dsDeliverableServiceLevels
					where dsDeliverableServiceLevel.Id == parID
					select dsDeliverableServiceLevel;

				var record = rsDeliverableServiceLevels.FirstOrDefault();
				if(record == null) // No record/entry was found
					{
					this.ID = 0;
					return false;
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.Optionality = record.OptionalityValue;
					this.AssociatedDeliverableID = record.Deliverable_Id;
					this.AssociatedServiceLevelID = record.Service_LevelId;
					this.AssociatedServiceProductID = record.Service_ProductId;
					this.ContentStatus = record.ContentStatusValue;
					this.Optionality = record.OptionalityValue;
					this.AdditionalConditions = record.AdditionalConditions;
					//Populate the Associated Service Level object if required
					if(parPopulateServiceLevelObject)
						{
						ServiceLevel objServiceLevel = new ServiceLevel();
						objServiceLevel.PopulateObject(parDatacontexSDDP, record.Service_LevelId);
						if(objServiceLevel != null && objServiceLevel.ID != 0)
							this.AssociatedServiceLevel = objServiceLevel;
						}

					// Populate the Associated Deliverable object if required
					if(parPopulateDeliverableObject)
						{
						Deliverable objDeliverable = new Deliverable();
						objDeliverable.PopulateObject(parDatacontexSDDP, record.Deliverable_Id, parGetLayer1up);
						if(objDeliverable != null && objDeliverable.ID != 0)
							this.AssociatedDeliverable = objDeliverable;
						}

					//Populate the Associated ServiceProduct object
					if(parPopulateServiceProductObject)
						{
						ServiceProduct objServiceProduct = new ServiceProduct();
						objServiceProduct.PopulateObject(parDatacontexSDDP, record.Service_ProductId);
						if(objServiceProduct != null && objServiceProduct.ID != 0)
							this.AssociatedServiceProduct = objServiceProduct;
						}

					} //if(record != null) // Feature Deliverable was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			} // end Populate method

		//-------------------------------------------------------
		// DeliverableServiceLevel - ObtainListOfServiceLevels_Summary 
		//-------------------------------------------------------
		/// <summary>
		/// Obtain a SUMMARY list of all the ServiceLevel objects that are associated with a SPECIFIC DeliverableServiceLevel - based on the parDeliverableID 
		/// parameter that must be provided. Only the following properties for each ServiceLevel will be returned: 
		/// ID, Title, ContentStatus - all the other properties of the Deliverable objects will be null.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parDeliverableID">Specify the the Deliverable's ID for which the List of Service Lvels must be retrieved and returned.</param>
		/// <returns>a List consisting of Service Level objects.</returns>
		public static List<ServiceLevel> ObtainListOfServiceLevels_Summary(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parDeliverableID,
			int parServiceProductID)
			{
			List<ServiceLevel> listServiceLevels = new List<ServiceLevel>();

			var dsDeliverableServiceLeverls = parDatacontextSDDP.DeliverableServiceLevels
				.Expand(delSL => delSL.Deliverable_)
				.Expand(delSL => delSL.Service_Product);

			try
				{
				// Access the Feature List
				var rsDeliverableSericeLevel =
					from dsDelSLs in dsDeliverableServiceLeverls
					where dsDelSLs.Deliverable_Id == parDeliverableID && dsDelSLs.Service_ProductId == parServiceProductID
					orderby dsDelSLs.Title
					select dsDelSLs;

				if(rsDeliverableSericeLevel.Count() == 0) // no records was found
					{
					return listServiceLevels;
					}

				foreach(var record in rsDeliverableSericeLevel)
					{
					ServiceLevel objServiceLevel = new ServiceLevel();
					objServiceLevel.ID = record.Service_Level.Id;
					objServiceLevel.Title = record.Service_Level.Title;
					objServiceLevel.ContentStatus = record.Service_Level.ContentStatusValue;
					listServiceLevels.Add(objServiceLevel);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listServiceLevels;
			} // end if ObtainListOfObjects
		}// end of DeliverableServiceLevels class

	// ####################################################
	// ### Deliverable Activities class
	// ####################################################
	/// <summary>
	/// 
	/// </summary>
	class DeliverableActivity
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public Activity AssociatedActivity{get; set;}
		public int? AssociatedActivityID{get; set;}
		// -------------------------------------
		// DeliverableActivity Populate Method
		//--------------------------------------
		public bool Populate(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parDeliverableActivityID,
			bool parPopulateDeliverableObject = false,
			bool parPopulateActivityObject = false,
			bool parGetLayer1up = false)
			{
			try
				{
				// Access the DeliverableServiceLevels List
				var dsDeliverableActivities = parDatacontexSDDP.DeliverableActivities
					.Expand(delAct => delAct.Deliverable_)
					.Expand(delAct => delAct.Activity_);

				var rsDeliverableActivities =
					from dsDeliverableActivity in dsDeliverableActivities
					where dsDeliverableActivity.Id == parDeliverableActivityID
					select dsDeliverableActivity;

				var record = rsDeliverableActivities.FirstOrDefault();
				if(record == null) // No record/entry was found
					{
					this.ID = 0;
					this.Title = "Deliverable Activity ID: " + parDeliverableActivityID + " could not be located in SharePoint ";
					return false;
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.Optionality = record.OptionalityValue;
					this.AssociatedDeliverableID = record.Deliverable_Id;
					this.AssociatedActivityID = record.Activity_Id;
					this.Optionality = record.OptionalityValue;
					//Populate the Associated Activity object if required
					if(parPopulateActivityObject)
						{
						Activity objActivity = new Activity();
						objActivity.PopulateObject(parDatacontexSDDP, record.Activity_Id);
						if(objActivity != null && objActivity.ID != 0)
							this.AssociatedActivity = objActivity;
						}

					// Populate the Associated Deliverable object if required
					if(parPopulateDeliverableObject)
						{
						Deliverable objDeliverable = new Deliverable();
						objDeliverable.PopulateObject(parDatacontexSDDP, record.Deliverable_Id, parGetLayer1up);
						if(objDeliverable != null && objDeliverable.ID != 0)
							this.AssociatedDeliverable = objDeliverable;
						}
					} //if(record != null) // Feature Deliverable was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			} // end Populate method

		//-------------------------------------------------------
		// DeliverableActivity - ObtainListOfActivities_Summary 
		//-------------------------------------------------------
		/// <summary>
		/// Obtain a SUMMARY list of all the Activity objects that are associated with a SPECIFIC Deliverable - based on the parDeliverableID 
		/// parameter that must be provided. Only the following properties for each ServiceLevel will be returned: 
		/// ID, Title, Optionality - all the other properties of the Deliverable objects will be null.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parDeliverableID">Specify the the Deliverable's ID for which the List of Activities must be retrieved and returned.</param>
		/// <returns>a List consisting of Activity objects.</returns>
		public static List<Activity> ObtainListOfActivities_Summary(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parDeliverableID)
			{
			List<Activity> listActivities = new List<Activity>();

			try
				{
				// Access the DeliverableActivities List
				var dsDeliverableActivities = parDatacontextSDDP.DeliverableActivities
				.Expand(delAct => delAct.Deliverable_)
				.Expand(delAct => delAct.Activity_);

				var rsDeliverableActivities =
					from dsDelActs in dsDeliverableActivities
					where dsDelActs.Deliverable_Id == parDeliverableID
					orderby dsDelActs.Title
					select dsDelActs;

				if(rsDeliverableActivities.Count() == 0) // no records was found
					{
					return listActivities;
					}

				foreach(var record in rsDeliverableActivities)
					{
					Activity objActivity = new Activity();
					objActivity.ID = record.Activity_.Id;
					objActivity.Title = record.Activity_.Title;
					objActivity.ContentStatus = record.Activity_.ContentStatusValue;
					listActivities.Add(objActivity);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listActivities;
			} // end if ObtainListOfObjects
		}// end of DeliverableActivities class

	//##########################################################
	/// <summary>
	/// This object represents an entry in the DeliverableTechnologies SharePoint List
	/// Each entry in the list is a DeliverableTechnology object.
	/// </summary>
	class DeliverableTechnology
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Considerations{get; set;}
		public TechnologyProduct TechnologyProduct {get; set;}
		public Deliverable	Deliviverable {get; set;}
		public string RoadmapStatus{get; set;}

		// ----------------------------
		// PopulateObject method
		//-----------------------------
		/// <summary>
		/// Populate the properties of the DeliverableTechnology object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the DeliverableTechnology that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parPopultateDeliverable = false)
			{
			try
				{
				// Access the DeliverableTechnologies List
				var dsDeliverableTechnologies = parDatacontexSDDP.DeliverableTechnologies
					.Expand(tp => tp.Deliverable_)
					.Expand(tp => tp.TechnologyProducts);

				var rsDeliverableTechnologies =
					from dsDeliverableTechnology in dsDeliverableTechnologies
					where dsDeliverableTechnology.Id == parID
					select dsDeliverableTechnology;

				var record = rsDeliverableTechnologies.FirstOrDefault();
				if(record == null) // was not found
					{
					this.ID = 0;
					this.Title = "DeliverableTechnology ID: " + parID + " could not be located in the SharePoint List";
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.Considerations = record.TechnologyConsiderations;
					if(parPopultateDeliverable)
						{
						if(record.Deliverable_ != null)
							{
							Deliverable objDeliverable = new Deliverable();
							objDeliverable.PopulateObject(
								parDatacontexSDDP: parDatacontexSDDP,
								parID: record.Deliverable_Id,
								parGetLayer1up: false,
								parGetRACI: false);
							this.Deliviverable = objDeliverable;
							}
						}

					if(record.TechnologyProducts != null)
						{
						TechnologyProduct objTechnologyProduct = new TechnologyProduct();
						objTechnologyProduct.ID = record.TechnologyProducts.Id;
						objTechnologyProduct.Title = record.TechnologyProducts.Title;
						objTechnologyProduct.Prerequisites = record.TechnologyProducts.TechnologyPrerequisites;
						if(record.TechnologyProducts.TechnologyCategory != null)
							{
							TechnologyCategory objTechnologCategory = new TechnologyCategory();
							objTechnologCategory.ID = record.TechnologyProducts.TechnologyCategory.Id;
							objTechnologCategory.Title = record.TechnologyProducts.TechnologyCategory.Title;
							objTechnologyProduct.Category = objTechnologCategory;
							}
						if(record.TechnologyProducts.TechnologyVendor != null)
							{
							TechnologyVendor objTechnologyVendor = new TechnologyVendor();
							objTechnologyVendor.ID = record.TechnologyProducts.TechnologyVendor.Id;
							objTechnologyVendor.Title = record.TechnologyProducts.TechnologyVendor.Title;
							objTechnologyProduct.Vendor = objTechnologyVendor;
							}
						this.TechnologyProduct = objTechnologyProduct;
						}
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			} // end of PopulateObject method

		//-------------------------------------------------------------------
		// DeliverableTechnology - ObtainListOfTechnologyProducts_Summary 
		//-------------------------------------------------------------------
		/// <summary>
		/// Obtain a Summary List of all the Technology Product objects that are associated with a SPECIFIC Deliverable - 
		/// The parDeliverableID parameter that must be provided. Only the following properties for each TechnologyProduct will be returned: 
		/// ID, Title, Category, Vendor,  - all the properties of the Deliverable object will be null.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parDeliverableID">Specify the the Deliverable's ID for which the List of Technology Products must be retrieved and returned.</param>
		/// <returns>a List consisting of TechnologyProduct objects.</returns>
		public static List<DeliverableTechnology> ObtainListOfTechnologyProducts_Summary(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parDeliverableID)
			{
			List<DeliverableTechnology> listDeliverableTechnologies = new List<DeliverableTechnology>();
			try
				{
				// Access the DeliverableTechnologies List
				var dsDeliverableTechnologies = parDatacontextSDDP.DeliverableTechnologies
				.Expand(delTech => delTech.TechnologyProducts)
				.Expand(delTech => delTech.TechnologyProducts.TechnologyVendor)
				.Expand(delTech => delTech.TechnologyProducts.TechnologyCategory);
				
				var rsDeliverableTechnologies =
					from dsDelTech in dsDeliverableTechnologies
					where dsDelTech.Deliverable_Id == parDeliverableID
					orderby dsDelTech.Title
					select dsDelTech;

				if(rsDeliverableTechnologies.Count() == 0) // no records was found
					{
					return listDeliverableTechnologies;
					}

				foreach(var record in rsDeliverableTechnologies)
					{
					DeliverableTechnology objDeliverableTechnology = new DeliverableTechnology();
					objDeliverableTechnology.ID = record.Id;
					objDeliverableTechnology.Title = record.Title;
					objDeliverableTechnology.RoadmapStatus = record.TechnologyRoadmapStatusValue;
					objDeliverableTechnology.Considerations = record.TechnologyConsiderations;
					// obtain the details of the TechnologyProduct and assign it to the TechnologyProduct object property of the DeliverableTechnology object.
					if(record.TechnologyProducts != null)
						{
						TechnologyProduct objTechnologyProduct = new TechnologyProduct();
						objTechnologyProduct.ID = record.TechnologyProducts.Id;
						objTechnologyProduct.Title = record.TechnologyProducts.Title;
						objTechnologyProduct.Prerequisites = record.TechnologyProducts.TechnologyPrerequisites;
						if(record.TechnologyProducts.TechnologyVendor != null)
							{
							TechnologyVendor objTechnologyVendor = new TechnologyVendor();
							objTechnologyVendor.ID = record.TechnologyProducts.TechnologyVendor.Id;
							objTechnologyVendor.Title = record.TechnologyProducts.TechnologyVendor.Title;
							objTechnologyProduct.Vendor = objTechnologyVendor;
							}
						if(record.TechnologyProducts.TechnologyCategory != null)
							{
							TechnologyCategory objTechnologyCategory = new TechnologyCategory();
							objTechnologyCategory.ID = record.TechnologyProducts.TechnologyCategory.Id;
							objTechnologyCategory.Title = record.TechnologyProducts.TechnologyCategory.Title;
							objTechnologyProduct.Category = objTechnologyCategory;
							}
						objDeliverableTechnology.TechnologyProduct = objTechnologyProduct;
						}
					// add the objDeliverableTechnology object to the listDeliverableTechnologies.
					listDeliverableTechnologies.Add(objDeliverableTechnology);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listDeliverableTechnologies;
			} // end if ObtainListOfObjects

		} // end of TechnologyProduct class


	//##########################################################
	//### FeatureDeliverable class
	//#########################################################
	/// <summary>
	/// The FeatureDeliverable object is the junction table or the cross-reference table between Service Features and Deliverables.
	/// </summary>
	class FeatureDeliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public ServiceFeature AssociatedFeature{get; set;}
		public int? AssociatedFeatureID{get; set;}

		// ----------------------------
		// Populate Method
		//-----------------------------
		public bool Populate(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parGetLayer1up = false,
			bool parPopulateFeatureObject = false,
			bool parPopulateDeliverableObject = false)
			{
			try
				{
				// Access the FeatureDeliverables List
				var dsFeatureDeliverables = parDatacontexSDDP.FeatureDeliverables
					.Expand(elDel => elDel.Deliverable_)
					.Expand(elDel => elDel.Service_Feature);

				var rsFeatureDeliverables =
					from dsFeatureDeliverable in dsFeatureDeliverables
					where dsFeatureDeliverable.Id == parID
					select dsFeatureDeliverable;

				var record = rsFeatureDeliverables.FirstOrDefault();
				if(record == null) // Feature Deliverable was not found
					{
					this.ID = 0;
					return false;
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.Optionality = record.OptionalityValue;
					this.AssociatedFeatureID = record.Service_FeatureId;
					this.AssociatedDeliverableID = record.Deliverable_Id;
					//Populate the Associated Service Feature object if required
					if(parPopulateFeatureObject)
						{
						ServiceFeature objServiceFeature = new ServiceFeature();
						objServiceFeature.PopulateObject(parDatacontexSDDP, record.Service_FeatureId, parGetLayer1up);
						if(objServiceFeature == null || objServiceFeature.ID == 0)
							{
							this.AssociatedFeature = null;
							}
						else
							{
							this.AssociatedFeature = objServiceFeature;
							}
						}
					else
						{
						this.AssociatedFeature = null;
						}
					// Populate the Associated Deliverable object if required
					if(parPopulateDeliverableObject)
						{
						Deliverable objDeliverable = new Deliverable();
						objDeliverable.PopulateObject(parDatacontexSDDP, record.Deliverable_Id, parGetLayer1up);
						if(objDeliverable == null || objDeliverable.ID == 0)
							{
							this.AssociatedDeliverable = null;
							}
						else
							{
							this.AssociatedDeliverable = objDeliverable;
							}
						}
					else
						{
						this.AssociatedDeliverable = null;
						}
					} //if(record != null) // Feature Deliverable was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			} // end Populate method

		//-------------------------------------------------------
		// FeatureDeliverable - ObtainListOfDeliverables_Detailed 
		//-------------------------------------------------------
		/// <summary>
		/// Obtain a comprehensive list of all the Deliverable objects that are associated with a SPECIFIC ServiceFeature - based on the parServiceFeatureID 
		/// parameter that must be provided.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parServiceFeatureID">Specify the the Service Feature for which the ListofDeliverables must be retrived and returned.</param>
		/// <param name="parGetContentLayers">When TRUE, the content layers of the each returned Deliverable object will be populated, else only the Deliverable object is returned and not any content layers if applicable on an object.</param>
		/// <returns>a List consisting of Deliverable objects.</returns>
		public static List<Deliverable> ObtainListOfDeliverables_Detailed(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parServiceFeatureID,
			bool parGetContentLayers = true)
			{
			List<Deliverable> listDeliverables = new List<Deliverable>();

			try
				{
				// Access the FeatureDeliverables List
				var rsFeatureDeliverables =
					from datasetFeautreDeliverables in parDatacontextSDDP.FeatureDeliverables
					where datasetFeautreDeliverables.Service_FeatureId == parServiceFeatureID
					orderby datasetFeautreDeliverables.Title
					select new
						{
						datasetFeautreDeliverables.Id,
						datasetFeautreDeliverables.Title,
						datasetFeautreDeliverables.OptionalityValue,
						datasetFeautreDeliverables.Deliverable_Id
						};

				if(rsFeatureDeliverables.Count() == 0) // no records was found
					{
					return listDeliverables;
					}

				foreach(var record in rsFeatureDeliverables)
					{
					Deliverable objDeliverable = new Deliverable();
					objDeliverable.PopulateObject(parDatacontextSDDP, record.Deliverable_Id, parGetContentLayers);
					if(objDeliverable == null || objDeliverable.ID == 0)
						{
						objDeliverable.ID = 0;
						objDeliverable.Title = "Deliverable Id: " + record.Id + " could not be found.";
						}
					listDeliverables.Add(objDeliverable);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listDeliverables;
			} // end if ObtainListOfObjects

		//-------------------------------------------------------
		// FeatureDeliverable - ObtainListOfDeliverables_Summary 
		//-------------------------------------------------------
		/// <summary>
		/// Obtain a SUMMARY list of all the Deliverable objects that are associated with a SPECIFIC ServiceFeature - based on the parServiceElemmentID 
		/// parameter that must be provided. Only the following properties for each deliverable will be returned: ID, Title, ISDsummary, CSDsummary, SoWsummary, ContentStatus - all the other properties of the Deliverable objects will be null. It will also not have the ContentLayers and Layer1up object populated.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parServiceFeatureID">Specify the the Service Feature for which the ListofDeliverables must be retrived and returned.</param>
		/// <param name="parGetContentLayers">When TRUE, the content layers of the each returned Deliverable object will be populated, else only the Deliverable object is returned and not any content layers if applicable on an object.</param>
		/// <returns>a List consisting of Deliverable objects.</returns>
		public static List<Deliverable> ObtainListOfDeliverables_Summary(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parServiceFeatureID)
			{
			List<Deliverable> listDeliverables = new List<Deliverable>();

			var dsFeatureDeliverables = parDatacontextSDDP.FeatureDeliverables
				.Expand(eldel => eldel.Deliverable_)
				.Expand(eldel => eldel.Deliverable_.ContentStatus);

			try
				{
				// Access the Feature List
				var rsFeatureDeliverables =
					from datasetEDs in dsFeatureDeliverables
					where datasetEDs.Service_FeatureId == parServiceFeatureID
					orderby datasetEDs.Title
					select datasetEDs;

				if(rsFeatureDeliverables.Count() == 0) // no records was found
					{
					return listDeliverables;
					}

				foreach(var record in rsFeatureDeliverables)
					{
					Deliverable objDeliverable = new Deliverable();
					objDeliverable.ID = record.Deliverable_.Id;
					objDeliverable.Title = record.Deliverable_.Title;
					objDeliverable.DeliverableType = record.Deliverable_.DeliverableTypeValue;
					objDeliverable.ISDsummary = record.Deliverable_.ISDSummary;
					objDeliverable.CSDsummary = record.Deliverable_.CSDSummary;
					objDeliverable.SoWsummary = record.Deliverable_.ContractSummary;
					objDeliverable.ContentStatus = record.Deliverable_.ContentStatusValue;
					listDeliverables.Add(objDeliverable);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listDeliverables;
			} // end if ObtainListOfObjects

		} // end of FeatureDeliverable class

	//##########################################################
	//### ElementDeliverable class
	//#########################################################
	/// <summary>
	/// The ElementDeliverable objects is the junction table or the cross-reference table between Service Elements and Deliverables.
	/// </summary>
	class ElementDeliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Optionality{get; set;}
		public Deliverable AssociatedDeliverable{get; set;}
		public int? AssociatedDeliverableID{get; set;}
		public ServiceElement AssociatedElement{get; set;}
		public int? AssociatedElementID{get; set;}

		// ----------------------------
		// Populate Method
		//-----------------------------
		public bool Populate(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID,
			bool parGetLayer1up = false,
			bool parPopulateElementObject = false,
			bool parPopulateDeliverableObject = false)
			{
			try
				{
				// Access the ElementDeliverables List
				var dsElementDeliverables = parDatacontexSDDP.ElementDeliverables
					.Expand(elDel => elDel.Deliverable_)
					.Expand(elDel => elDel.Service_Element);

				var rsElementDeliverables =
					from dsElementDeliverable in dsElementDeliverables
					where dsElementDeliverable.Id == parID
					select dsElementDeliverable;

				var record = rsElementDeliverables.FirstOrDefault();
				if(record == null) // Element Deliverable was not found
					{
					this.ID = 0;
					return false;
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.Optionality = record.OptionalityValue;
					this.AssociatedElementID = record.Service_ElementId;
					this.AssociatedDeliverableID = record.Deliverable_Id;
					//Populate the Associated Service Element object if required
					if(parPopulateElementObject)
						{
						ServiceElement objServiceElement = new ServiceElement();
						objServiceElement.PopulateObject(parDatacontexSDDP, record.Service_ElementId, parGetLayer1up);
						if(objServiceElement == null || objServiceElement.ID == 0)
							{
							this.AssociatedElement = null;
							}
						else
							{
							this.AssociatedElement = objServiceElement;
							}
						}
					else
						{
						this.AssociatedElement = null;
						}
					// Populate the Associated Deliverable object if required
					if(parPopulateDeliverableObject)
						{
						Deliverable objDeliverable = new Deliverable();
						objDeliverable.PopulateObject(parDatacontexSDDP, record.Deliverable_Id, parGetLayer1up);
						if(objDeliverable == null || objDeliverable.ID == 0)
							{
							this.AssociatedDeliverable = null;
							}
						else
							{
							this.AssociatedDeliverable = objDeliverable;
							}
						}
					else
						{
						this.AssociatedDeliverable = null;
						}
					} //if(record != null) // Element Deliverable was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			} // end Populate method

		//-------------------------------------------------------
		// ElementDeliverable - ObtainListOfDeliverables_Detailed 
		//-------------------------------------------------------
		/// <summary>
		/// Obtain a comprehensive list of all the Deliverable objects that are associated with a SPECIFIC ServiceElement - based on the parServiceElemmentID 
		/// parameter that must be provided.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parServiceElementID">Specify the the Service Element for which the ListofDeliverables must be retrived and returned.</param>
		/// <param name="parGetContentLayers">When TRUE, the content layers of the each returned Deliverable object will be populated, else only the Deliverable object is returned and not any content layers if applicable on an object.</param>
		/// <returns>a List consisting of Deliverable objects.</returns>
		public static List<Deliverable> ObtainListOfDeliverables_Detailed(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parServiceElementID,
			bool parGetContentLayers = true)
			{
			List<Deliverable> listDeliverables = new List<Deliverable>();

			try
				{
				// Access the ElementDeliverables List
				var rsElementDeliverables =
					from datasetEDs in parDatacontextSDDP.ElementDeliverables
					where datasetEDs.Service_ElementId == parServiceElementID
					orderby datasetEDs.Title
					select new
						{
						datasetEDs.Id,
						datasetEDs.Title,
						datasetEDs.OptionalityValue,
						datasetEDs.Deliverable_Id
						};

				if(rsElementDeliverables.Count() == 0) // no records was found
					{
					return listDeliverables;
					}

				foreach(var record in rsElementDeliverables)
					{
					Deliverable objDeliverable = new Deliverable();
					objDeliverable.PopulateObject(parDatacontextSDDP, record.Deliverable_Id, parGetContentLayers);
					if(objDeliverable == null || objDeliverable.ID == 0)
						{
						objDeliverable.ID = 0;
						objDeliverable.Title = "Deliverable Id: " + record.Id + " could not be found.";
						}
					listDeliverables.Add(objDeliverable);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listDeliverables;
			} // end if ObtainListOfObjects

		//-------------------------------------------------------
		// ElementDeliverable - ObtainListOfDeliverables_Summary 
		//-------------------------------------------------------
		/// <summary>
		/// Obtain a SUMMARY list of all the Deliverable objects that are associated with a SPECIFIC ServiceElement - based on the parServiceElemmentID 
		/// parameter that must be provided. Only the following properties for each deliverable will be returned: ID, Title, ISDsummary, CSDsummary, SoWsummary, ContentStatus - all the other properties of the Deliverable objects will be null. It will also not have the ContentLayers and Layer1up object populated.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parServiceElementID">Specify the the Service Element for which the ListofDeliverables must be retrived and returned.</param>
		/// <param name="parGetContentLayers">When TRUE, the content layers of the each returned Deliverable object will be populated, else only the Deliverable object is returned and not any content layers if applicable on an object.</param>
		/// <returns>a List consisting of Deliverable objects.</returns>
		public static List<Deliverable> ObtainListOfDeliverables_Summary(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parServiceElementID)
			{
			List<Deliverable> listDeliverables = new List<Deliverable>();

			var dsElementDeliverables = parDatacontextSDDP.ElementDeliverables
				.Expand(eldel => eldel.Deliverable_)
				.Expand(eldel => eldel.Deliverable_.ContentStatus);

			try
				{
				// Access the ElementDeliverables List
				var rsElementDeliverables =
					from datasetEDs in dsElementDeliverables
					where datasetEDs.Service_ElementId == parServiceElementID
					orderby datasetEDs.Title
					select datasetEDs;

				if(rsElementDeliverables.Count() == 0) // no records was found
					{
					return listDeliverables;
					}

				foreach(var record in rsElementDeliverables)
					{
					Deliverable objDeliverable = new Deliverable();
					objDeliverable.ID = record.Deliverable_.Id;
					objDeliverable.Title = record.Deliverable_.Title;
					objDeliverable.DeliverableType = record.Deliverable_.DeliverableTypeValue;
					objDeliverable.ISDsummary = record.Deliverable_.ISDSummary;
					objDeliverable.CSDsummary = record.Deliverable_.CSDSummary;
					objDeliverable.SoWsummary = record.Deliverable_.ContractSummary;
					objDeliverable.ContentStatus = record.Deliverable_.ContentStatusValue;
					listDeliverables.Add(objDeliverable);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listDeliverables;
			} // end if ObtainListOfObjects

		} // end of ElementDeliverable class

	// ###################################
	// ### Mapping Object
	// ###################################

	/// <summary>
	/// The Mapping object represents an entry in the Mappings List in SharePoint.
	/// </summary>
	class Mapping
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string ClientName{get; set;}

		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the Mapping object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				var dsMappings = parDatacontexSDDP.Mappings
					.Expand(map => map.Client_);

				// Access the Mappings List
				var rsMappings =
					from dsMapping in dsMappings
					where dsMapping.Id == parID
					select dsMapping;

				var recMapping = rsMappings.FirstOrDefault();
				if(recMapping == null) // Mapping was not found
					{
					throw new DataEntryNotFoundException("Client Requirements Mapping entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recMapping.Id;
					this.Title = recMapping.Title;
					this.ClientName = recMapping.Client_.DocGenClientName;
					} //if(recFeature != null) // Mapping was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			}
		} // end Class Mapping

	//###############################################
	/// <summary>
	/// The MappingServiceTower object represents an entry in the Mapping Service Towers List in SharePoint.
	/// </summary>
	class MappingServiceTower
		{
		public int ID{get; set;}
		public string Title{get; set;}
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the MappingServiceTower object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping Service Tower that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				// Access the Mapping Service Towers List
				var rsMappingTowers =
					from dsTower in parDatacontexSDDP.MappingServiceTowers
					where dsTower.Id == parID
					select new
						{
						dsTower.Id,
						dsTower.Title
						};

				var recTower = rsMappingTowers.FirstOrDefault();
				if(recTower == null) // MappingTower was not found
					{
					throw new DataEntryNotFoundException("Mapping Tower entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recTower.Id;
					this.Title = recTower.Title;
					} //if(recTower != null) // Mapping Tower was found
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			} // end of PopulateObject method

		//-----------------------------------------------
		// MappingServiceTower - ObtainListOfObjects 
		//----------------------------------------------
		/// <summary>
		/// Obtain a list containing all the MappingServiceTower objects associated with the value in the parMappingID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parMappingID"></param>
		/// <returns></returns>
		public static List<MappingServiceTower> ObtainListOfObjects(DesignAndDeliveryPortfolioDataContext parDatacontextSDDP, int parMappingID)
			{
			List<MappingServiceTower> listMappingTowers = new List<MappingServiceTower>();

			try
				{
				// Access the Mapping Service Towers List
				var rsMappingTowers =
					from dsTower in parDatacontextSDDP.MappingServiceTowers
					where dsTower.Mapping_Id == parMappingID
					orderby dsTower.Title
					select new
						{
						dsTower.Id,
						dsTower.Title
						};

				if(rsMappingTowers.Count() == 0) // MappingTowers was not found
					{
					//throw new DataEntryNotFoundException("No Mapping Tower entries for Mapping ID:" +
					//	parMappingID + " could be found in SharePoint.");
					return listMappingTowers;
					}

				foreach(var recTower in rsMappingTowers)
					{
					MappingServiceTower objMappingTower = new MappingServiceTower();
					objMappingTower.ID = recTower.Id;
					objMappingTower.Title = recTower.Title;
					listMappingTowers.Add(objMappingTower);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return listMappingTowers;
			} // end if ObtainListOfObjects

		} // end Class Mapping Service Towers

	//##########################################
	/// <summary>
	/// The MappingRequirement object represents an entry in the MappingRequirements List.
	/// </summary>
	class MappingRequirement
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string RequirementText{get; set;}
		public string RequirementServiceLevel{get; set;}
		public string SourceReference{get; set;}
		public string ComplianceStatus{get; set;}
		public string ComplianceComments{get; set;}

		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the MappingRequirement object
		/// </summary>
		/// <param name="parDatacontexSDDP">Pass a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping Requirement that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				// Access the Mapping Requirements List
				var rsRequirements =
					from dsRequirement in parDatacontexSDDP.MappingRequirements
					where dsRequirement.Id == parID
					select new
						{
						dsRequirement.Id,
						dsRequirement.Title,
						dsRequirement.RequirementText,
						dsRequirement.RequirementServiceLevel,
						dsRequirement.SourceReference,
						dsRequirement.ComplianceStatusValue,
						dsRequirement.ComplianceComments
						};

				if(rsRequirements.Count() == 0) // Mapping Requirement was not found
					{
					throw new DataEntryNotFoundException("Mapping Requirement entry ID:" +
						parID + " could not be found in SharePoint.");
					}

				var recRequirement = rsRequirements.FirstOrDefault();
				if(recRequirement == null) // Mapping Requirement was not found
					{
					throw new DataEntryNotFoundException("Mapping Requirement entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recRequirement.Id;
					this.Title = recRequirement.Title;
					this.RequirementText = recRequirement.RequirementText;
					this.RequirementServiceLevel = recRequirement.RequirementServiceLevel;
					this.SourceReference = recRequirement.SourceReference;
					this.ComplianceStatus = recRequirement.ComplianceStatusValue;
					this.ComplianceComments = recRequirement.ComplianceComments;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			} // end of PopulateObject Method

		///----------------------------------------------
		/// <summary>
		/// Obtain a list containing all the MappingRequirement objects associated with the value in the parMappingTowerID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parMappingTowerID">The ID od the MappingServiceTower for which the list of MappingRequirementObjects must be returned</param>
		/// <returns></returns>
		public static List<MappingRequirement> ObtainListOfObjects(DesignAndDeliveryPortfolioDataContext parDatacontextSDDP, int parMappingTowerID)
			{
			List<MappingRequirement> listMappingRequirements = new List<MappingRequirement>();

			try
				{
				// Access the Mapping Requirements List
				var rsMappingRequirements =
					from dsRequirement in parDatacontextSDDP.MappingRequirements
					where dsRequirement.Mapping_TowerId == parMappingTowerID
					orderby dsRequirement.SortOrder
					select new
						{
						dsRequirement.Id,
						dsRequirement.Title,
						dsRequirement.RequirementText,
						dsRequirement.RequirementServiceLevel,
						dsRequirement.SourceReference,
						dsRequirement.ComplianceStatusValue,
						dsRequirement.ComplianceComments
						};

				//if(rsMappingRequirements.Count() == 0) // No MappingRequirements was not found
				//	{
				//	throw new DataEntryNotFoundException("No Mapping Requirement entries for Mapping Service Tower ID:" +
				//		parMappingTowerID + " could be found in SharePoint.");
				//	}

				foreach(var recRequirement in rsMappingRequirements)
					{
					MappingRequirement objMappingRequirement = new MappingRequirement();
					objMappingRequirement.ID = recRequirement.Id;
					objMappingRequirement.Title = recRequirement.Title;
					objMappingRequirement.RequirementText = recRequirement.RequirementText;
					objMappingRequirement.RequirementServiceLevel = recRequirement.RequirementServiceLevel;
					objMappingRequirement.SourceReference = recRequirement.SourceReference;
					objMappingRequirement.ComplianceStatus = recRequirement.ComplianceStatusValue;
					objMappingRequirement.ComplianceComments = recRequirement.ComplianceComments;
					listMappingRequirements.Add(objMappingRequirement);
					}

				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return listMappingRequirements;
			} // end if ObtainListOfObjects
		} // end Class Mapping Requirements

	//############################################
	/// <summary>
	/// The Mapping Deliverable is the class used to for the Mapping Deliverables SharePoint List.
	/// </summary>
	//############################################
	class MappingDeliverable
		{
		public int ID{get; set;}
		public string Title{get; set;}
		/// <summary>
		/// Represents the translated value in the Deliverable Choice column of the MappingDeliverable List. TRUE if "New" else FALSE
		/// </summary>
		public bool NewDeliverable{get; set;}
		public string ComplianceComments{get; set;}
		public String NewRequirement{get; set;}
		/// <summary>
		/// This Property represents a complete Deliverable Object
		/// </summary>
		public Deliverable MappedDeliverable{get; set;}
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the MappingDeliverable object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping Deliverable that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{

				// Retrieve the data from the Mapping Deliverable List
				var rsMappingDeliverables =
					from dsMappingDeliverable in parDatacontexSDDP.MappingDeliverables
					where dsMappingDeliverable.Id == parID
					select new
						{
						dsMappingDeliverable.Id,
						dsMappingDeliverable.Title,
						dsMappingDeliverable.DeliverableChoiceValue,
						dsMappingDeliverable.DeliverableRequirement,
						dsMappingDeliverable.ComplianceComments,
						dsMappingDeliverable.Mapped_DeliverableId
						};

				var recMappingDeliverable = rsMappingDeliverables.FirstOrDefault();
				if(recMappingDeliverable == null) // Mapping Deliverable was not found
					{
					throw new DataEntryNotFoundException("Mapping Deliverable entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recMappingDeliverable.Id;
					this.Title = recMappingDeliverable.Title;
					if(recMappingDeliverable.DeliverableChoiceValue.Contains("New"))
						{
						this.NewDeliverable = true;
						this.NewRequirement = recMappingDeliverable.DeliverableRequirement;
						}
					else
						{
						this.NewDeliverable = false;
						this.MappedDeliverable = new Deliverable();
						try
							{
							// Populate the MappedDeliverable
							this.MappedDeliverable.PopulateObject(
								parDatacontexSDDP: parDatacontexSDDP,
								parID: recMappingDeliverable.Mapped_DeliverableId);
							}
						catch(DataEntryNotFoundException exc)
							{
							this.MappedDeliverable = null;
							}
						}
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			}

		///----------------------------------------------
		/// <summary>
		/// Obtain a list containing all the MappingDeliverable objects associated with the value in the parMappingRequirementID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parMappingRequirementID">The ID od the MappingServiceTower for which the list of MappingRequirementObjects must be returned</param>
		/// <returns></returns>
		public static List<MappingDeliverable> ObtainListOfObjects(
			DesignAndDeliveryPortfolioDataContext parDatacontextSDDP,
			int parMappingRequirementID)
			{
			List<MappingDeliverable> listMappingDeliverables = new List<MappingDeliverable>();

			try
				{
				// Access the Mapping Deliverables List
				var rsMappingDeliverables =
					from dsMappingDeliverable in parDatacontextSDDP.MappingDeliverables
					where dsMappingDeliverable.Mapping_RequirementId == parMappingRequirementID
					orderby dsMappingDeliverable.Title
					select new
						{
						dsMappingDeliverable.Id,
						dsMappingDeliverable.Title,
						dsMappingDeliverable.DeliverableChoiceValue,
						dsMappingDeliverable.DeliverableRequirement,
						dsMappingDeliverable.Mapped_DeliverableId,
						dsMappingDeliverable.ComplianceComments
						};

				if(rsMappingDeliverables.Count() == 0) // No MappingRequirements was not found
					{
					//throw new DataEntryNotFoundException("No Mapping Requirement entries for Mapping Service Tower ID:" +
					//	parMappingRequirementID + " could be found in SharePoint.");
					return listMappingDeliverables;
					}

				// Process all the relevant entries and add them to the list of Mapped Deliverables
				foreach(var recMappingDeliverable in rsMappingDeliverables)
					{
					MappingDeliverable objMappingDeliverable = new MappingDeliverable();
					objMappingDeliverable.ID = recMappingDeliverable.Id;
					objMappingDeliverable.Title = recMappingDeliverable.Title;
					objMappingDeliverable.ComplianceComments = recMappingDeliverable.ComplianceComments;

					if(recMappingDeliverable.DeliverableChoiceValue.Contains("New"))
						{
						objMappingDeliverable.NewDeliverable = true;
						objMappingDeliverable.NewRequirement = recMappingDeliverable.DeliverableRequirement;
						}
					else //...Contains("Existing"))
						{
						objMappingDeliverable.NewDeliverable = false;
						objMappingDeliverable.MappedDeliverable = new Deliverable();
						try
							{
							// Populate the MappedDeliverable with Deliverable Data
							objMappingDeliverable.MappedDeliverable.PopulateObject(
								parDatacontexSDDP: parDatacontextSDDP,
								parID: recMappingDeliverable.Mapped_DeliverableId, parGetLayer1up: true);
							if(objMappingDeliverable.MappedDeliverable.ID == 0)
								{
								objMappingDeliverable.MappedDeliverable = null;
								}
							}

						catch(DataEntryNotFoundException)
							{
							objMappingDeliverable.MappedDeliverable = null;
							}
						}
					listMappingDeliverables.Add(objMappingDeliverable);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return listMappingDeliverables;
			} // end if ObtainListOfObjects
		}

	//#############################################
	/// <summary>
	/// The MappingAssumption represents an entry of the Mapping Assumptions List in SharePoint
	/// </summary>
	class MappingAssumption
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Description{get; set;}
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the MappingAssumption object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping Assumption that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{

				// Access the Mapping Assumptions List
				var rsAssumptions =
					from dsAssumption in parDatacontexSDDP.MappingAssumptions
					where dsAssumption.Id == parID
					select new
						{
						dsAssumption.Id,
						dsAssumption.Title,
						dsAssumption.AssumptionDescription,
						};

				var recAssumption = rsAssumptions.FirstOrDefault();
				if(recAssumption == null) // Mapping Assumption was not found
					{
					throw new DataEntryNotFoundException("Mapping Assumption entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recAssumption.Id;
					this.Title = recAssumption.Title;
					this.Description = recAssumption.AssumptionDescription;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			}

		///----------------------------------------------
		/// <summary>
		/// Obtain a list containing all the MappingAssumption objects associated with the value in the parMappingRequirementID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parMappingRequirementID">The ID of the MappingRequirement for which the list of MappingAssumption Objects must be returned</param>
		/// <returns>List of MappingRisks object</returns>
		public static List<MappingAssumption> ObtainListOfObjects(DesignAndDeliveryPortfolioDataContext parDatacontextSDDP, int parMappingRequirementID)
			{
			List<MappingAssumption> listMappingAssumptions = new List<MappingAssumption>();

			try
				{
				// Access the Mapping Assumption List
				var rsMappingAssumptions =
					from dsAssumption in parDatacontextSDDP.MappingAssumptions
					where dsAssumption.Mapping_RequirementId == parMappingRequirementID
					orderby dsAssumption.Title
					select new
						{
						dsAssumption.Id,
						dsAssumption.Title,
						dsAssumption.AssumptionDescription
						};

				//if(rsMappingAssumptions.Count() == 0) // No Mapping Assumptions were not found
				//	{
				//	throw new DataEntryNotFoundException("No Mapping Assumption entries for Mapping Requirement ID:" +
				//		parMappingRequirementID + " could be found in SharePoint.");
				//	}

				foreach(var recMappingAssumption in rsMappingAssumptions)
					{
					MappingAssumption objMappingAssumption = new MappingAssumption();
					objMappingAssumption.ID = recMappingAssumption.Id;
					objMappingAssumption.Title = recMappingAssumption.Title;
					objMappingAssumption.Description = recMappingAssumption.AssumptionDescription;
					listMappingAssumptions.Add(objMappingAssumption);
					}

				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return listMappingAssumptions;
			} // end if ObtainListOfObjects
		}
	//##################################################
	/// <summary>
	/// Mapping Risk Object
	/// </summary>
	class MappingRisk
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Statement{get; set;}
		public string Mitigation{get; set;}
		public double? ExposureValue{get; set;}
		public string Status{get; set;}
		public string Exposure{get; set;}
		public string ComplianceStatus{get; set;}
		public string ComplianceComments{get; set;}
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the MappingRisk object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping Risk that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{

				// Access the Service Features List
				var rsRisks =
					from dsRisk in parDatacontexSDDP.MappingRisks
					where dsRisk.Id == parID
					select new
						{
						dsRisk.Id,
						dsRisk.Title,
						dsRisk.RiskExposureValue0,
						dsRisk.RiskExposureValue,
						dsRisk.RiskMitigation,
						dsRisk.RiskStatement,
						dsRisk.RiskStatusValue
						};

				var recRisk = rsRisks.FirstOrDefault();
				if(recRisk == null) // Mapping Requirement was not found
					{
					throw new DataEntryNotFoundException("Mapping Requirement entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recRisk.Id;
					this.Title = recRisk.Title;
					this.Statement = recRisk.RiskStatement;
					this.Mitigation = recRisk.RiskMitigation;
					this.Exposure = recRisk.RiskExposureValue;
					this.ExposureValue = recRisk.RiskExposureValue0;
					this.Status = recRisk.RiskStatusValue;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			} // end of Method PopulateObject

		///----------------------------------------------
		/// <summary>
		/// Obtain a list containing all the MappingRisks objects associated with the value in the parMappingRequirementID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parMappingRequirementID">The ID of the MappingRequirement for which the list of MappingRisk Objects must be returned</param>
		/// <returns>List of MappingRisks object</returns>
		public static List<MappingRisk> ObtainListOfObjects(DesignAndDeliveryPortfolioDataContext parDatacontextSDDP, int parMappingRequirementID)
			{
			List<MappingRisk> listMappingRisks = new List<MappingRisk>();
			try
				{
				// Access the Mapping Risk List
				var rsMappingRisks =
					from dsRisk in parDatacontextSDDP.MappingRisks
					where dsRisk.Mapping_RequirementId == parMappingRequirementID
					orderby dsRisk.Title
					select new
						{
						dsRisk.Id,
						dsRisk.Title,
						dsRisk.RiskStatement,
						dsRisk.RiskMitigation,
						dsRisk.RiskExposureValue,
						dsRisk.RiskExposureValue0,
						dsRisk.RiskStatusValue
						};

				//if(rsMappingRisks.Count() == 0) // No MappingRequirements was not found
				//	{
				//	throw new DataEntryNotFoundException("No Mapping Risk entries for Mapping Requirement ID:" +
				//		parMappingRequirementID + " could be found in SharePoint.");
				//	}

				foreach(var recMappingRisk in rsMappingRisks)
					{
					MappingRisk objMappingRisk = new MappingRisk();
					objMappingRisk.ID = recMappingRisk.Id;
					objMappingRisk.Title = recMappingRisk.Title;
					objMappingRisk.Statement = recMappingRisk.RiskStatement;
					objMappingRisk.Mitigation = recMappingRisk.RiskMitigation;
					objMappingRisk.Exposure = recMappingRisk.RiskExposureValue;
					objMappingRisk.ExposureValue = recMappingRisk.RiskExposureValue0;
					objMappingRisk.Status = recMappingRisk.RiskStatusValue;
					listMappingRisks.Add(objMappingRisk);
					}

				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return listMappingRisks;
			} // end if ObtainListOfObjects
		} // End of Class MappingRisk


	/// <summary>
	/// The Mapping Service Level is the class used to for the Mapping Service Levels SharePoint List.
	/// </summary>
	class MappingServiceLevel
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string RequirementText{get; set;}
		public bool NewServiceLevel{get; set;}
		public string ServiceLevelText{get; set;}
		/// <summary>
		/// This property represents a complete Service Level object.
		/// </summary>
		public ServiceLevel MappedServiceLevel{get; set;}
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the MappingServiceLevel object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Mapping Service Level that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			bool? newServiceLevel = false;
			try
				{
				var dsMappingServiceLevels = parDatacontexSDDP.MappingServiceLevels
					.Expand(map => map.Service_Level);

				// Access the Mapping Service Levels List
				var rsMappingServiceLevels =
					from dsServiceLevel in dsMappingServiceLevels
					where dsServiceLevel.Id == parID
					select dsServiceLevel;

				var recServiceLevel = rsMappingServiceLevels.FirstOrDefault();
				if(recServiceLevel == null) // Mapping Service Level was not found
					{
					throw new DataEntryNotFoundException("Mapping Service Levels entry ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.MappedServiceLevel = new ServiceLevel();
					this.ID = recServiceLevel.Id;
					this.Title = recServiceLevel.Title;
					newServiceLevel = recServiceLevel.NewServiceLevel;
					if(newServiceLevel != null)
						{
						if(newServiceLevel == true)
							{
							this.RequirementText = recServiceLevel.ServiceLevelRequirement;
							this.NewServiceLevel = true;
							}
						else
							{
							this.NewServiceLevel = false;
							this.RequirementText = recServiceLevel.Service_Level.CSDHeading;
							ServiceLevel objServiceLevel = new ServiceLevel();
							objServiceLevel.PopulateObject(parDatacontexSDDP: parDatacontexSDDP, ServiceLevelID: recServiceLevel.Service_LevelId);
							if(objServiceLevel.Title != null)
								{
								this.MappedServiceLevel = objServiceLevel;
								}
							}
						}
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			}

		///----------------------------------------------
		/// <summary>
		/// Returns a list containing all the MappingServiceLevel objects associated with the value in the parMappingRequirementID parameter.
		/// </summary>
		/// <param name="parDatacontextSDDP"></param>
		/// <param name="parMappingDeliverableID">The ID of the MappingDeliverable for which the list of MappingServiceLevel Objects must be returned</param>
		/// <returns>List of MappingRisks object</returns>
		public static List<MappingServiceLevel> ObtainListOfObjects(DesignAndDeliveryPortfolioDataContext parDatacontextSDDP, int parMappingDeliverableID)
			{
			List<MappingServiceLevel> listMappingServiceLevels = new List<MappingServiceLevel>();

			bool? newServiceLevel = false;
			try
				{
				var dsMappingServiceLevels = parDatacontextSDDP.MappingServiceLevels
					.Expand(map => map.Mapping_Deliverable)
					.Expand(map => map.Service_Level);

				// Access the Mapping Service Levels List
				var rsMappingServiceLevels =
					from dsMappingSL in dsMappingServiceLevels
					where dsMappingSL.Mapping_DeliverableId == parMappingDeliverableID
					orderby dsMappingSL.Title
					select dsMappingSL;

				//if(rsMappingServiceLevels.Count() == 0) // No MappingServiceLevels were found
				//	{
				//	throw new DataEntryNotFoundException("No Mapping Service Level entries for Mapping Deliverable ID:" +
				//		parMappingDeliverableID + " could be found in SharePoint.");
				//	}

				foreach(var recMappingSL in rsMappingServiceLevels)
					{
					MappingServiceLevel objMappingServiceLevel = new MappingServiceLevel();
					objMappingServiceLevel.ID = recMappingSL.Id;
					objMappingServiceLevel.Title = recMappingSL.Title;
					newServiceLevel = recMappingSL.NewServiceLevel;
					if(newServiceLevel != null)
						{
						if(newServiceLevel == true)
							{
							objMappingServiceLevel.RequirementText = recMappingSL.ServiceLevelRequirement;
							objMappingServiceLevel.NewServiceLevel = true;
							}
						else
							{
							objMappingServiceLevel.NewServiceLevel = false;
							objMappingServiceLevel.RequirementText = recMappingSL.Service_Level.CSDHeading;
							ServiceLevel objServiceLevel = new ServiceLevel();
							objServiceLevel.PopulateObject(parDatacontexSDDP: parDatacontextSDDP, ServiceLevelID: recMappingSL.Service_LevelId);
							if(objServiceLevel.Title != null)
								{
								objMappingServiceLevel.MappedServiceLevel = objServiceLevel;
								}
							}
						}
					listMappingServiceLevels.Add(objMappingServiceLevel);
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return listMappingServiceLevels;
			} // end if ObtainListOfObjects

		}

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Service Levels SharePoint List
	/// </summary>
	class ServiceLevel{public int ID{get; set;}
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
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the ServiceLevel object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Service Level that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? ServiceLevelID)
			{
			try
				{
				// Access the Service Levels List
				var dsServiceLevels = parDatacontexSDDP.ServiceLevels
					.Expand(level => level.Service_Hour);

				var rsServiceLevels =
					from dsServiceLevel in dsServiceLevels
					where dsServiceLevel.Id == ServiceLevelID
					select dsServiceLevel;

				var recServiceLevel = rsServiceLevels.FirstOrDefault();
				if(recServiceLevel == null) // Service Level was not found
					{
					throw new DataEntryNotFoundException("Service Levels entry ID:" +
						ServiceLevelID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recServiceLevel.Id;
					this.Title = recServiceLevel.Title;
					this.ContentStatus = recServiceLevel.ContentStatusValue;
					this.ISDheading = recServiceLevel.ISDHeading;
					this.ISDdescription = recServiceLevel.ISDDescription;
					this.CSDheading = recServiceLevel.CSDHeading;
					this.CSDdescription = recServiceLevel.CSDDescription;
					this.SOWheading = recServiceLevel.ContractHeading;
					this.SOWdescription = recServiceLevel.ContractDescription;
					this.Measurement = recServiceLevel.ServiceLevelMeasurement;
					this.MeasurementInterval = recServiceLevel.MeasurementIntervalValue;
					this.ReportingInterval = recServiceLevel.ReportingIntervalValue;
					this.CalcualtionMethod = recServiceLevel.CalculationMethod;
					this.CalculationFormula = recServiceLevel.CalculationFormula;
					this.ServiceHours = recServiceLevel.Service_Hour.Title;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			// Load the Service Level Performance Thresholds
			this.PerfomanceThresholds = new List<ServiceLevelTarget>();
			try
				{
				var dsThresholds =
					from dsThreshold in parDatacontexSDDP.ServiceLevelTargets
					where dsThreshold.Service_LevelId == this.ID && dsThreshold.ThresholdOrTargetValue == "Threshold"
					orderby dsThreshold.Title
					select dsThreshold;

				foreach(var thresholdItem in dsThresholds)
					{
					ServiceLevelTarget objSLthreshold = new ServiceLevelTarget();
					objSLthreshold.ID = thresholdItem.Id;
					objSLthreshold.Title = thresholdItem.Title.Substring(thresholdItem.Title.IndexOf(": ", 0) + 2, (thresholdItem.Title.Length - thresholdItem.Title.IndexOf(": ", 0) + 2));
					objSLthreshold.Type = thresholdItem.ThresholdOrTarget.Value;
					objSLthreshold.ContentStatus = thresholdItem.ContentStatusValue;
					this.PerfomanceThresholds.Add(objSLthreshold);
					}
				}
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			// Load the Service Level Performance Targets
			this.PerformanceTargets = new List<ServiceLevelTarget>();
			try
				{
				var dsTargetss =
					from dsThreshold in parDatacontexSDDP.ServiceLevelTargets
					where dsThreshold.Service_LevelId == this.ID && dsThreshold.ThresholdOrTargetValue == "Target"
					orderby dsThreshold.Title
					select dsThreshold;

				foreach(var targetItem in dsTargetss)
					{
					ServiceLevelTarget objSLtarget = new ServiceLevelTarget();
					objSLtarget.ID = targetItem.Id;
					objSLtarget.Title = targetItem.Title.Substring(targetItem.Title.IndexOf(": ", 0) + 2, (targetItem.Title.Length - targetItem.Title.IndexOf(": ", 0) + 2));
					objSLtarget.Type = targetItem.ThresholdOrTarget.Value;
					objSLtarget.ContentStatus = targetItem.ContentStatusValue;
					this.PerfomanceThresholds.Add(objSLtarget);
					}
				}
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return;
			} // end of PopulateObject method
		} // end of Service Levels class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Activities SharePoint List
	/// </summary>
	class ServiceLevelTarget
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
	class Activity
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

		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the Activities object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parActivityID">Receives the Identifier of the Activity that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parActivityID)
			{
			try
				{
				// Access the Activities List

				var dsActivities = parDatacontexSDDP.Activities
					.Expand(act => act.Responsible_RACI)
					.Expand(act => act.Accountable_RACI)
					.Expand(act => act.Consulted_RACI)
					.Expand(act => act.Informed_RACI)
					.Expand(act => act.Activity_Category);

				var rsActivities =
					from dsActivity in dsActivities
					where dsActivity.Id == parActivityID
					select dsActivity;

				var record = rsActivities.FirstOrDefault();
				if(record == null) // Activity was not found
					{
					this.ID = 0;
					this.Title = "Activity ID: " + parActivityID + " could not be located in the SharePoint List";
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.SortOrder = record.SortOrder;
					this.Optionality = record.ActivityOptionalityValue;
					this.ISDheading = record.ISDHeading;
					this.ISDdescription = record.ISDDescription;
					this.CSDheading = record.CSDHeading;
					this.CSDdescription = record.CSDDescription;
					this.SOWheading = record.ContractHeading;
					this.SOWdescription = record.ContractDescription;
					this.ContentStatus = record.ContentStatusValue;
					this.Input = record.ActivityInput;
					this.Output = record.ActivityOutput;
					this.Catagory = record.Activity_Category.Title;
					this.Assumptions = record.ActivityAssumptions;
					this.OLAvariations = record.OLAVariations;
					// Add the RACI Accountable entry to the list if there are any associated.
					if(record.Accountable_RACI.Title != null)
						{
						this.RACI_Accountable = new List<JobRole>();
						JobRole objJobRole = new JobRole();
						objJobRole.ID = record.Accountable_RACI.Id;
						objJobRole.Title = record.Accountable_RACI.Title;
						this.RACI_Accountable.Add(objJobRole);
						}
					// add the RACI Responsible entries to the list if there are any associated.
					if(record.Responsible_RACI.Count > 0)
						{
						this.RACI_Responsible = new List<JobRole>();
						foreach(var item in record.Responsible_RACI)
							{
							JobRole objJobRole = new JobRole();
							objJobRole.ID = item.Id;
							objJobRole.Title = item.Title;
							this.RACI_Responsible.Add(objJobRole);
							}
						}
					// add the RACI Consulted entries to the list if there are any associated.
					if(record.Consulted_RACI.Count > 0)
						{
						this.RACI_Consulted = new List<JobRole>();
						foreach(var item in record.Consulted_RACI)
							{
							JobRole objJobRole = new JobRole();
							objJobRole.ID = item.Id;
							objJobRole.Title = item.Title;
							this.RACI_Consulted.Add(objJobRole);
							}
						}
					// add the RACI Informed entries to the list if there are any associated.
					if(record.Informed_RACI.Count > 0)
						{
						this.RACI_Informed = new List<JobRole>();
						foreach(var item in record.Informed_RACI)
							{
							JobRole objJobRole = new JobRole();
							objJobRole.ID = item.Id;
							objJobRole.Title = item.Title;
							this.RACI_Informed.Add(objJobRole);
							}
						}
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return;
			} // end of PopulateObject method
		} // end of Activitiy class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Job Framewotk Alignment SharePoint List
	/// But each entry is essentially a JobRole, therefore the class is named JobRole
	/// </summary>
	class JobRole
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string DeliveryDomain{get; set;}
		public string SpecificRegion{get; set;}
		public string RelevantBusinessUnit{get; set;}
		public string OtherJobTitles{get; set;}
		public string JobFrameworkLink{get; set;}
		// ----------------------------
		// Methods
		//-----------------------------
		/// <summary>
		/// Populate the properties of the Activities object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parJobID">Receives the Identifier of the Activity that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parJobID)
			{
			try
				{
				// Access the Job Framework Alignment List
				var dsJobFrameworks = parDatacontexSDDP.JobFrameworkAlignment
					.Expand(jf => jf.JobDeliveryDomain);

				var rsJobFrameworks =
					from dsJobFramework in dsJobFrameworks
					where dsJobFramework.Id == parJobID
					select dsJobFramework;

				var record = rsJobFrameworks.FirstOrDefault();
				if(record == null) // Job was not found
					{
					this.ID = 0;
					this.Title = "Job Framework ID: " + parJobID + " could not be located in the SharePoint List";
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.OtherJobTitles = record.RelatedRoleTitle;
					if(record.JobDeliveryDomain.Title != null)
						this.DeliveryDomain = record.JobDeliveryDomain.Title;
					if(record.RelevantBusinessUnitValue != null)
						this.RelevantBusinessUnit = record.RelevantBusinessUnitValue;
					if(record.SpecificRegionValue != null)
						this.SpecificRegion = record.SpecificRegionValue;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return;
			} // end of PopulateObject method
		} // end of JobRole class

	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Technology Categories SharePoint List
	/// Each entry in the list is a Technology Category object.
	/// </summary>
	class TechnologyCategory
		{
		public int ID{get; set;}
		public string Title{get; set;}

		// ----------------------------
		// PopulateObject method
		//-----------------------------
		/// <summary>
		/// Populate the properties of the TechnologyCategory object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parTechnologyCategoryID">Receives the Identifier of the Technology Category that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parTechnologyCategoryID)
			{
			try
				{
				// Access the Technology Categories List
				var rsTechCategories =
					from dsTechCategory in parDatacontexSDDP.TechnologyCategories
					where dsTechCategory.Id == parTechnologyCategoryID
					select dsTechCategory;

				var record = rsTechCategories.FirstOrDefault();
				if(record == null) // was not found
					{
					this.ID = 0;
					this.Title = "Technology Category ID: " + parTechnologyCategoryID + " could not be located in the SharePoint List";
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return;
			} // end of PopulateObject method


		} // end of TechnologyCategory class


	//##########################################################
	/// <summary>
	/// This object repsents an entry in the Technology Vendors SharePoint List
	/// Each entry in the list is a Technology Vendor object.
	/// </summary>
	class TechnologyVendor
		{
		public int ID{get; set;}
		public string Title{get; set;}
		
		// ----------------------------
		// PopulateObject method
		//-----------------------------
		/// <summary>
		/// Populate the properties of the TechnologyVendor object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Technology Vendor that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				// Access the Technology Vendors List
				var rsTechVendors =
					from dsTechVendor in parDatacontexSDDP.TechnologyVendors
					where dsTechVendor.Id == parID
					select dsTechVendor;

				var record = rsTechVendors.FirstOrDefault();
				if(record == null) // was not found
					{
					this.ID = 0;
					this.Title = "Technology Vendor ID: " + parID + " could not be located in the SharePoint List";
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return;
			} // end of PopulateObject method


		} // end of TechnologyVendor class

	//##########################################################
	/// <summary>
	/// This object represents an entry in the Technology Products SharePoint List
	/// Each entry in the list is a Technology Product object.
	/// </summary>
	class TechnologyProduct
		{
		public int ID{get; set;}
		public string Title{get; set;}
		public string Prerequisites{get; set;}
		public TechnologyCategory Category{get; set;}
		public TechnologyVendor Vendor {get; set;}

		// ----------------------------
		// PopulateObject method
		//-----------------------------
		/// <summary>
		/// Populate the properties of the TechnologyProduct object
		/// </summary>
		/// <param name="parDatacontexSDDP">Receives a predefined DataContext object which is used to access the SharePoint Data</param>
		/// <param name="parID">Receives the Identifier of the Technology Product that need to be retrieved from SharePoint</param>
		public void PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				// Access the Technology Products List
				var dsTechProducts = parDatacontexSDDP.TechnologyProducts
					.Expand(tp => tp.TechnologyVendor)
					.Expand(tp => tp.TechnologyCategory);

				var rsTechProducts =
					from dsTechProduct in parDatacontexSDDP.TechnologyProducts
					where dsTechProduct.Id == parID
					select dsTechProduct;

				var record = rsTechProducts.FirstOrDefault();
				if(record == null) // was not found
					{
					this.ID = 0;
					this.Title = "Technology Product ID: " + parID + " could not be located in the SharePoint List";
					}
				else
					{
					this.ID = record.Id;
					this.Title = record.Title;
					this.Prerequisites = record.TechnologyPrerequisites;
					if(record.TechnologyCategory != null)
						{
						TechnologyCategory objTechnologyCategory = new TechnologyCategory();
						objTechnologyCategory.ID = record.TechnologyCategory.Id; 
						objTechnologyCategory.Title = record.TechnologyCategory.Title;
						this.Category = objTechnologyCategory;
						}

					if(record.TechnologyVendor != null)
						{
						TechnologyVendor objTechnologyVendor = new TechnologyVendor();
						objTechnologyVendor.ID = record.TechnologyVendor.Id;
						objTechnologyVendor.Title = record.TechnologyVendor.Title;
						this.Vendor = objTechnologyVendor;
                              }
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			} // end of PopulateObject method
		} // end of TechnologyProduct class

	class CompleteDataSet
		{
		public Dictionary<int, JobRole> dsJobroles{get; set;}
		public Dictionary<int, GlossaryAcronym> dsGlossaryAcronyms{get; set;}
		public Dictionary<int,ServicePortfolio> dsPortfolios {get; set;}
		public Dictionary<int,ServiceFamily> dsFamilies{get; set;}
		public Dictionary<int, ServiceProduct> dsProducts {get; set;}
		public Dictionary<int, ServiceElement> dsElements{get; set;}
		public Dictionary<int, ServiceFeature> dsFeatures{get; set;}
		public Dictionary<int, Deliverable> dsDeliverables{get; set;}
		public Dictionary<int, ElementDeliverable> dsElementDeliverables {get; set;}
		public Dictionary<int, FeatureDeliverable> dsFeatureDeliverables{get; set;}
		public Dictionary<int, Activity> dsActivities{get; set;}
		public Dictionary<int, DeliverableActivity> dsDeliverableActivities{get; set;}
		public Dictionary<int, TechnologyProduct> dsTechnologyProducts{get; set;}
		public Dictionary<int, DeliverableTechnology> dsDeliverableTechnologies{get; set;}
		public Dictionary<int, ServiceLevel> dsServiceLevels{get; set;}
		public Dictionary<int, DeliverableServiceLevel> dsDeliverableServiceLevels{get; set;}
		public bool PopulateObject(
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
				Console.Write("\tPopulating the complete DataSet...");

				// -------------------------
				// Populate GlossaryAcronyms
				Console.Write("\n\t + Glossary & Acronyms...");
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
                    Console.Write("\t {0} - {1}", this.dsGlossaryAcronyms.Count, DateTime.Now - setStart);

				// Populate JobRoles
				Console.Write("\n\t + JobRoles...");
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
                    Console.Write("\t {0} - {1}", this.dsJobroles.Count, DateTime.Now - setStart);

				// -------------------------
				// Populate TechnologyProdcuts
				Console.Write("\n\t + TechnologyProducts...");
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
                    Console.Write("\t {0} - {1}", this.dsTechnologyProducts.Count, DateTime.Now - setStart);

				//--------------------------------
				// Populate the Service Portfolios
				Console.Write("\n\t + ServicePortfolios...");
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
                    Console.Write("\t {0} - {1}", this.dsPortfolios.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Families
				Console.Write("\n\t + ServiceFamilies...");
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
                    Console.Write("\t {0} - {1}", this.dsFamilies.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Products
				Console.Write("\n\t + ServiceProducts...");
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
                    Console.Write("\t {0} - {1}", this.dsProducts.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Element 
				Console.Write("\n\t + ServiceElements...");
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
						objElement.ServiceProductID = recElement.Service_PortfolioId;
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
                    Console.Write("\t {0} - {1}", this.dsElements.Count, DateTime.Now - setStart);

				//--------------------------	
				// Populate Service Feature 
				Console.Write("\n\t + ServiceFeatures...");
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
						objFeature.ID = recFeature.Id;
						intLastReadID = recFeature.Id;
						boolFetchMore = true;
						objFeature.Title = recFeature.Title;
						objFeature.ServiceProductID = recFeature.Service_PortfolioId;
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
					Console.Write("\t {0} - {1}", this.dsFeatures.Count, DateTime.Now - setStart);
					
				//-----------------------
				// Populate Deliverables
				Console.Write("\n\t + Deliverables...");
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
				Console.Write("\t {0} - {1}", this.dsDeliverables.Count, DateTime.Now - setStart);

				//--------------------------------------
				// Populate Service Element Deliverables
				Console.Write("\n\t + ElementDeliverables...");
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
				Console.Write("\t {0} - {1}", this.dsElementDeliverables.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate Service Feature Deliverables
				Console.Write("\n\t + FeatureDeliverables...");
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
                    Console.Write("\t {0} - {1}", this.dsFeatureDeliverables.Count, DateTime.Now - setStart);

				//---------------------------------------
				// Populate DeliverableTechnologies
				Console.Write("\n\t + DeliverableTechnologies...");
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
                    Console.Write("\t {0} - {1}", this.dsDeliverableTechnologies.Count, DateTime.Now - setStart);

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
                    Console.Write("\t {0} - {1}", this.dsActivities.Count, DateTime.Now - setStart);


				//---------------------------------------
				// Populate DeliverableActivities
				//---------------------------------------
				Console.Write("\n\t + DeliverableActivities...");
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
                    Console.Write("\t {0} - {1}", this.dsDeliverableActivities.Count, DateTime.Now - setStart);

				// -------------------------
				// Populate ServiceLevels
				// -------------------------
				Console.Write("\n\t + ServiceLevels...");
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

					this.dsServiceLevels = new Dictionary<int, ServiceLevel>();
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
						objServiceLevel.ServiceHours = record.Service_Hour.Title;
						objServiceLevel.PerfomanceThresholds = new List<ServiceLevelTarget>();

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

                                        objSLthreshold.Title = thresholdItem.Title.Substring(thresholdItem.Title.IndexOf(": ", 0) + 2, thresholdItem.Title.Length - thresholdItem.Title.IndexOf(": ", 0) - 2);
								objSLthreshold.Type = thresholdItem.ThresholdOrTargetValue;
								objSLthreshold.ContentStatus = thresholdItem.ContentStatusValue;
								objServiceLevel.PerfomanceThresholds.Add(objSLthreshold);
								}
							}

						// Load the Service Level Performance Targets
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
								objSLtarget.Title = targetEntry.Title.Substring(targetEntry.Title.IndexOf(": ", 0) + 2, (targetEntry.Title.Length - targetEntry.Title.IndexOf(": ", 0) - 2));
								objSLtarget.Type = targetEntry.ThresholdOrTargetValue;
								objSLtarget.ContentStatus = targetEntry.ContentStatusValue;
								objServiceLevel.PerformanceTargets.Add(objSLtarget);
								}
							}
						this.dsServiceLevels.Add(key: record.Id, value: objServiceLevel);
						}
					} while(boolFetchMore);
                    Console.Write("\t {0} - {1}", this.dsServiceLevels.Count, DateTime.Now - startTime);

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
                    Console.WriteLine("\t {0} - {1}", this.dsDeliverableServiceLevels.Count, DateTime.Now - setStart);
					
				Console.WriteLine("\tPopulating the complete DataSet took ended at {0} and took {1}.", DateTime.Now, DateTime.Now - startTime);
				return true;
				}
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			}
		}
	}
