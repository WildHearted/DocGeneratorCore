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

		public string CSDheading
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
					this.ContentLayerValue = this.ContentLayerValue;
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

		public string iSDsummary
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

		public decimal SortOrder
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

		} // end Class Deliverables
	}
