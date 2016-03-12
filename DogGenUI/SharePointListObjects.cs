using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocGenerator.SDDPServiceReference;

namespace DocGenerator
	{
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
	}
