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
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string ISDheading
			{
			get; set;
			}

		public string ISDdescription
			{
			get; set;
			}

		public string CSDheading
			{
			get; set;
			}

		public string CSDdescription
			{
			get; set;
			}

		public string SOWheading
			{
			get; set;
			}

		public string SOWdescription
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID)
			{
			try
				{
				// Access the Service Portfolios List
				var rsPortfolios =
					from dsPortfolio in parDatacontexSDDP.ServicePortfolios
					where dsPortfolio.Id == parID
					select new
						{
						dsPortfolio.Id,
						dsPortfolio.Title,
						dsPortfolio.ISDHeading,
						dsPortfolio.ISDDescription,
						dsPortfolio.CSDHeading,
						dsPortfolio.CSDDescription,
						dsPortfolio.ContractHeading,
						dsPortfolio.ContractDescription
						};

				var recPortfolio = rsPortfolios.FirstOrDefault();
				if(recPortfolio == null) // Service Portfolio was not found
					{
					throw new DataEntryNotFoundException("Service Portfolio content for ID:" +
						parID + " could not be found in SharePoint.");
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
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			}

		} // end of class ServicePortfolio

	class ServiceFamily
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string ISDheading
			{
			get; set;
			}

		public string ISDdescription
			{
			get; set;
			}

		public string CSDheading
			{
			get; set;
			}

		public string CSDdescription
			{
			get; set;
			}

		public string SOWheading
			{
			get; set;
			}

		public string SOWdescription
			{
			get; set;
			}

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
					select new
						{
						dsFamilies.Id,
						dsFamilies.Title,
						dsFamilies.ISDHeading,
						dsFamilies.ISDDescription,
						dsFamilies.CSDHeading,
						dsFamilies.CSDDescription,
						dsFamilies.ContractHeading,
						dsFamilies.ContractDescription
						};

				var recFamily = rsFamilies.FirstOrDefault();
				if(recFamily == null) // Service Family was not found
					{
					throw new DataEntryNotFoundException("Service Family content for ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recFamily.Id;
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

	class ServiceProduct
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string ISDheading
			{
			get; set;
			}

		public string ISDdescription
			{
			get; set;
			}

		public string KeyDDbenefits
			{
			get; set;
			}

		public string KeyClientBenefits
			{
			get; set;
			}

		public string CSDheading
			{
			get; set;
			}

		public string CSDdescription
			{
			get; set;
			}

		public string SOWheading
			{
			get; set;
			}

		public string SOWdescription
			{
			get; set;
			}

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
					select new
						{
						dsProduct.Id,
						dsProduct.Title,
						dsProduct.ISDHeading,
						dsProduct.ISDDescription,
						dsProduct.KeyDDBenefits,
						dsProduct.KeyClientBenefits,
						dsProduct.CSDHeading,
						dsProduct.CSDDescription,
						dsProduct.ContractHeading,
						dsProduct.ContractDescription
						};

				var recProduct = rsProducts.FirstOrDefault();
				if(recProduct == null) // Service Product was not found
					{
					throw new DataEntryNotFoundException("Service Product content for ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recProduct.Id;
					this.Title = recProduct.Title;
					this.ISDheading = recProduct.ISDHeading;
					this.ISDdescription = recProduct.ISDDescription;
					this.KeyClientBenefits = recProduct.KeyClientBenefits;
					this.KeyDDbenefits = recProduct.KeyDDBenefits;
					this.CSDheading = recProduct.CSDHeading;
					this.CSDdescription = recProduct.CSDDescription;
					this.SOWheading = recProduct.ContractHeading;
					this.SOWdescription = recProduct.ContractDescription;
					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return true;
			}

		} // end of class ServiceProduct

		
	class ServiceElement
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public double? SortOrder
			{
			get; set;
			}

		public string ISDheading
			{
			get; set;
			}

		public string ISDdescription
			{
			get; set;
			}

		public string Objectives
			{
			get; set;
			}

		public string KeyClientAdvantages
			{
			get; set;
			}

		public string KeyClientBenefits
			{
			get; set;
			}

		public string KeyDDbenefits
			{
			get; set;
			}

		public string KeyPerformanceIndicators
			{
			get; set;
			}

		public string CriticalSuccessFactors
			{
			get; set;
			}

		public string ProcessLink
			{
			get; set;
			}

		public string ContentLayerValue
			{
			get; set;
			}

		public int? ContentPredecessorElementID
			{
			get; set;
			}


		public ServiceElement Layer1up
			{
			get; set;
			}

		// ----------------------------
		// Methods
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
					select new
						{
						dsElement.Id,
						dsElement.Title,
						dsElement.SortOrder,
						dsElement.ISDHeading,
						dsElement.ISDDescription,
						dsElement.Objective,
						dsElement.KeyClientAdvantages,
						dsElement.KeyClientBenefits,
						dsElement.KeyDDBenefits,
						dsElement.KeyPerformanceIndicators,
						dsElement.CriticalSuccessFactors,
						dsElement.ProcessLink,
						dsElement.ContentLayerValue,
						dsElement.ContentPredecessorElementId
						};

				var recElement = rsElements.FirstOrDefault();
				if(recElement == null) // Service Element was not found
					{
					throw new DataEntryNotFoundException("Service Element content for ID:" +
						parID + " could not be found in SharePoint.");
					}
				else
					{
					this.ID = recElement.Id;
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
			}
		} // end Class ServiceElement

	class ServiceFeature
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public double? SortOrder
			{
			get; set;
			}

		public string CSDheading
			{
			get; set;
			}

		public string CSDdescription
			{
			get; set;
			}

		public string SOWheading
			{
			get; set;
			}

		public string SOWdescription
			{
			get; set;
			}

		public string ContentLayerValue
			{
			get; set;
			}

		public int? ContentPredecessorFeatureID
			{
			get; set;
			}

		public ServiceFeature Layer1up
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID, bool parGetLayer1up = false)
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
						dsFeature.ContentPredecessorFeatureId
						};

				var recFeature = rsFeatures.FirstOrDefault();
				if(recFeature == null) // Service Feature was not found
					{
					throw new DataEntryNotFoundException("Service Feature content for ID:" +
						parID + " could not be found in SharePoint.");
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
					
					//this.ContentLayerValue = this.ContentLayerValue;
					this.ContentLayerValue = recFeature.ContentLayerValue;
					this.ContentPredecessorFeatureID = recFeature.ContentPredecessorFeatureId;
					if(parGetLayer1up == true && recFeature.ContentPredecessorFeatureId != null)
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
			}
		} // end Class ServiceFeature

	class Deliverable
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string ISDheading
			{
			get; set;
			}

		public string ISDdescription
			{
			get; set;
			}

		public string ISDsummary
			{
			get; set;
			}

		public string CSDheading
			{
			get; set;
			}

		public string CSDdescription
			{
			get; set;
			}

		public string CSDsummary
			{
			get; set;
			}

		public string SoWheading
			{
			get; set;
			}

		public string SoWdescription
			{
			get; set;
			}

		public string SoWsummary
			{
			get; set;
			}

		public string DeliverableType
			{
			get; set;
			}

		public string Inputs
			{
			get; set;
			}

		public string Outputs
			{
			get; set;
			}

		public string DDobligations
			{
			get; set;
			}

		public string ClientResponsibilities
			{
			get; set;
			}

		public string Exclusions
			{
			get; set;
			}

		public string GovernanceControls
			{
			get; set;
			}

		public double? SortOrder
			{
			get; set;
			}

		public string TransitionDescription
			{
			get; set;
			}

		public string WhatHasChanged
			{
			get; set;
			}

		public string ContentLayerValue
			{
			get; set;
			}

		private Dictionary<int, String> _glossaryAndAcronyms = new Dictionary<int, string>();
		public Dictionary<int, string> GlossaryAndAcronyms
			{
			get{return this._glossaryAndAcronyms;}
			set{this._glossaryAndAcronyms = value;}
			}

		public int? ContentPredecessorDeliverableID
			{
			get; set;
			}

		public Deliverable Layer1up
			{
			get; set;
			}

		public bool PopulateObject(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int? parID, bool parGetLayer1up = false)
			{
			try
				{
				// Access the Service Elements List
				var dsDeliverables = parDatacontexSDDP.Deliverables
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
					throw new DataEntryNotFoundException("Content for Deliverable ID:" +
						parID + " could not be found in SharePoint.");
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
					this.ContentLayerValue = recDeliverable.ContentLayerValue;
					this.ContentPredecessorDeliverableID = recDeliverable.ContentPredecessor_DeliverableId;

					// Add the Glossary and Acronym terms to the Deliverable object
					if(recDeliverable.GlossaryAndAcronyms.Count > 0)
						{
						foreach(var entry in recDeliverable.GlossaryAndAcronyms)
							{
							if(this.GlossaryAndAcronyms.ContainsKey(entry.Id) != true)
								this.GlossaryAndAcronyms.Add(entry.Id, entry.Title);
							}
						}
					// Add the recursive relationship of Content Predecessors
					if(parGetLayer1up == true && recDeliverable.ContentPredecessor_DeliverableId != null)
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
			return true;
			} // end of Method PopulateObject

		} // end Class Deliverables


	class Mapping
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}


		public string ClientName
			{
			get; set;
			}	

		// ----------------------------
		// Methods
		//-----------------------------
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



	class MappingServiceTower
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
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
						{dsTower.Id,
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
			}
		} // end Class Mapping Service Towers


	class MappingRequirement
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string RequirementText
			{
			get; set;
			}

		public string RequirementServiceLevel
			{
			get; set;
			}

		public string SourceReference
			{
			get; set;
			}

		public string ComplianceStatus
			{
			get; set;
			}

		public string ComplianceComments
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
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
						{dsRequirement.Id,
						dsRequirement.Title,
						dsRequirement.RequirementText,
						dsRequirement.RequirementServiceLevel,
						dsRequirement.SourceReference,
						dsRequirement.ComplianceStatusValue,
						dsRequirement.ComplianceComments
						};

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
			}
		} // end Class Mapping Requirements

	class MappingAssumption
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string Description
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
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
		}

	class MappingRisk
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string Statement
			{
			get; set;
			}

		public string Mitigation
			{
			get; set;
			}

		public double? ExposureValue
			{
			get; set;
			}

		public string Status
			{
			get; set;
			}

		public string Exposure
			{
			get; set;
			}

		public string ComplianceStatus
			{
			get; set;
			}

		public string ComplianceComments
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
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
			}
		}

	class MappingServiceLevel
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string RequirementText
			{
			get; set;
			}

		public string ServiceLevelText
			{
			get; set;
			}

		public ServiceLevel MappedServiceLevel
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
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
						if(newServiceLevel == true)
							this.RequirementText = recServiceLevel.ServiceLevelRequirement;
						else
							{
							this.RequirementText = recServiceLevel.Service_Level.CSDHeading;
							ServiceLevel objServiceLevel = new ServiceLevel();
							objServiceLevel.PopulateObject(parDatacontexSDDP: parDatacontexSDDP, ServiceLevelID: recServiceLevel.Id);

							}
					

					}
				} // try
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}
			return;
			}
		}


	class ServiceLevel
		{
		public int ID
			{
			get; set;
			}

		public string Title
			{
			get; set;
			}

		public string ISDheading
			{
			get; set;
			}

		public string ISDdescription
			{
			get; set;
			}

		public string CSDheading
			{
			get; set;
			}

		public string CSDdescription
			{
			get; set;
			}

		public string SOWheading
			{
			get; set;
			}

		public string SOWdescription
			{
			get; set;
			}

		public string Measurement
			{
			get; set;
			}

		public string MeasurementInterval
			{
			get; set;
			}

		public string ReportingInterval
			{
			get; set;
			}

		public string CalcualtionMethod
			{
			get; set;
			}

		public string CalculationFormula
			{
			get; set;
			}

		public string ServiceHours
			{
			get; set;
			}

		public List<string> PerfomanceThresholds
			{
			get; set;
			}

		public List<string> PerformanceTargets
			{
			get; set;
			}

		public string BasicConditions
			{
			get; set;
			}

		public string AdditionalConditions
			{
			get; set;
			}

		// ----------------------------
		// Methods
		//-----------------------------
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
			this.PerfomanceThresholds = new List<string>();
			try
				{
				var dsThresholds =
					from dsThreshold in parDatacontexSDDP.ServiceLevelTargets
					where dsThreshold.Service_LevelId == this.ID && dsThreshold.ThresholdOrTargetValue == "Threshold"
					orderby dsThreshold.Title
					select dsThreshold;
				
				foreach(var thresholdItem in dsThresholds)
					{
					this.PerfomanceThresholds.Add(thresholdItem.Title.Substring(thresholdItem.Title.IndexOf(": ",0) + 2, (thresholdItem.Title.Length - thresholdItem.Title.IndexOf(": ", 0) + 2)));
					}
				}
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			// Load the Service Level Performance Targets
			this.PerformanceTargets = new List<string>();
			try
				{
				var dsTargetss =
					from dsThreshold in parDatacontexSDDP.ServiceLevelTargets
					where dsThreshold.Service_LevelId == this.ID && dsThreshold.ThresholdOrTargetValue == "Target"
					orderby dsThreshold.Title
					select dsThreshold;

				foreach(var targetItem in dsTargetss)
					{
					this.PerformanceTargets.Add(targetItem.Title.Substring(targetItem.Title.IndexOf(": ", 0) + 2, (targetItem.Title.Length - targetItem.Title.IndexOf(": ", 0) + 2)));
					}
				}
			catch(DataServiceClientException exc)
				{
				throw new DataServiceClientException("Unable to access SharePoint Error: " + exc.HResult + " - " + exc.Message);
				}

			return;
			} // end of PopulateObject method
		} // end of Service Levels class
	}
