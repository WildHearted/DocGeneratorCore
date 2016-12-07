using System;
using System.Collections.Generic;
using System.Data.Services.Client;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocGeneratorCore.SDDPServiceReference;
using VelocityDb;
using VelocityDb.Session;
using DocGeneratorCore.Database.Classes;

namespace DocGeneratorCore.Database.Functions
	{
	class UpdateLocalDatabase
		{

		#region Methods
		public static void UpadateData(DesignAndDeliveryPortfolioDataContext parSDDPdatacontext)
			{
			try
				{
				Stopwatch stopwatchCompleteDataSet = Stopwatch.StartNew();

				//+ Please Note:
				//---G
				//- SharePoint's REST API has a limit which returns only 1000 entries at a time
				//- therefore a paging principle had to be implemented to return all the entries in the List.
				//---G
				Stopwatch stopwatch = Stopwatch.StartNew();
				int entriesCounter = 0;
				int totalEntriesUpdated = 0;
				int lastReadID = 0;
				bool fetchMoreIndicator = true;

				using (ServerClientSession dbSession = new ServerClientSession(
					systemDir: Properties.Settings.Default.CurrentDatabaseLocation,
					systemHost: Properties.Settings.Default.CurrentDatabaseHost))
					{
					//---G
					//+ Populate **GlossaryAcronyms**
					
					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsGlossaryAcronyms =
								from dsGlossaryAcronym in parSDDPdatacontext.GlossaryAndAcronyms
								where dsGlossaryAcronym.Id > lastReadID
								&& dsGlossaryAcronym.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsGlossaryAcronym;

							entriesCounter = 0;

							foreach (GlossaryAndAcronymsItem record in rsGlossaryAcronyms)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								GlossaryAcronym objGlossaryAcronym = (from theEntry in dbSession.AllObjects<GlossaryAcronym>()
																		where theEntry.IDsp == record.Id
																		select theEntry).FirstOrDefault();

								if (objGlossaryAcronym == null)
									objGlossaryAcronym = new GlossaryAcronym();

								objGlossaryAcronym.IDsp = record.Id;
								objGlossaryAcronym.Term = record.Title;
								objGlossaryAcronym.Acronym = record.Acronym;
								objGlossaryAcronym.Meaning = record.Definition;
								dbSession.Persist(objGlossaryAcronym);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating GlossaryAcronyms ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with GlossaryAcronyms");
						}
					stopwatch.Stop();
					Console.Write("\n\t + Glossary & Acronyms...\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **JobRoles**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					try
						{
						dbSession.BeginUpdate();
						var dsJobFrameworks = parSDDPdatacontext.JobFrameworkAlignment
							.Expand(jf => jf.JobDeliveryDomain);

						while (fetchMoreIndicator)
							{
							var rsJobFrameworks =
								from dsJobFramework in dsJobFrameworks
								where dsJobFramework.Id > lastReadID
								&& dsJobFramework.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsJobFramework;

							entriesCounter = 0;

							foreach (JobFrameworkAlignmentItem record in rsJobFrameworks)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								JobRole objJobRole = (from theEntry in dbSession.AllObjects<JobRole>()
														where theEntry.IDsp == record.Id
														select theEntry).FirstOrDefault();

								if (objJobRole == null)
									objJobRole = new JobRole();

								objJobRole.IDsp = record.Id;
								objJobRole.Title = record.Title;
								objJobRole.OtherJobTitles = record.RelatedRoleTitle;
								if (record.JobDeliveryDomain.Title != null)
									objJobRole.DeliveryDomain = record.JobDeliveryDomain.Title;
								if (record.RelevantBusinessUnitValue != null)
									objJobRole.RelevantBusinessUnit = record.RelevantBusinessUnitValue;
								if (record.SpecificRegionValue != null)
									objJobRole.SpecificRegion = record.SpecificRegionValue;

								dbSession.Persist(objJobRole);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating JobRoles ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with JobRoles");
						}

					stopwatch.Stop();
					Console.Write("\n\t + JobRoles...\t\t\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **TechnologyCategories**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsTechnologyCategories =
								from dsTechCategory in parSDDPdatacontext.TechnologyCategories
								where dsTechCategory.Id > lastReadID
								&& dsTechCategory.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsTechCategory;

							entriesCounter = 0;

							foreach (TechnologyCategoriesItem record in rsTechnologyCategories)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = Convert.ToInt16(record.Id);

								TechnologyCategory technologyCategory = (from theEntry in dbSession.AllObjects<TechnologyCategory>()
																	where theEntry.IDsp == record.Id
																	select theEntry).FirstOrDefault();
								if (technologyCategory == null)
									technologyCategory = new TechnologyCategory();

								technologyCategory.IDsp = record.Id;
								technologyCategory.Title = record.Title;

								dbSession.Persist(technologyCategory);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating TechnologyCategories ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with TechnologyCategories");
						}

					stopwatch.Stop();
					Console.Write("\n\t + TechnologyCategories...\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **TechnologyVendors**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsTechnologyVendors =
								from dsTechVendor in parSDDPdatacontext.TechnologyVendors
								where dsTechVendor.Id > lastReadID
								&& dsTechVendor.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsTechVendor;

							entriesCounter = 0;

							foreach (TechnologyVendorsItem record in rsTechnologyVendors)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;

								TechnologyVendor technologyVendor = (from theEntry in dbSession.AllObjects<TechnologyVendor>()
																	where theEntry.IDsp == record.Id
																	select theEntry).FirstOrDefault();
								if (technologyVendor == null)
									technologyVendor = new TechnologyVendor();

								technologyVendor.IDsp = record.Id;
								lastReadID = record.Id;
								technologyVendor.Title = record.Title;

								dbSession.Persist(technologyVendor);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating TechnologyVendors ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with TechnologyVendors");
						}

					stopwatch.Stop();
					Console.Write("\n\t + TechnologyVendors...\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **TechnologyProducts**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						var dsTechnologyProducts = parSDDPdatacontext.TechnologyProducts
							.Expand(tp => tp.TechnologyCategory)
							.Expand(tp => tp.TechnologyVendor);

						while (fetchMoreIndicator)
							{
							var rsTechnologyProducts =
								from dsTechProduct in dsTechnologyProducts
								where dsTechProduct.Id > lastReadID
								&& dsTechProduct.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsTechProduct;

							entriesCounter = 0;

							foreach (TechnologyProductsItem record in rsTechnologyProducts)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;

								TechnologyProduct objTechProduct = (from theEntry in dbSession.AllObjects<TechnologyProduct>()
																	where theEntry.IDsp == record.Id
																	select theEntry).FirstOrDefault();
								if (objTechProduct == null)
									objTechProduct = new TechnologyProduct();

								objTechProduct.IDsp = record.Id;
								lastReadID = record.Id;
								objTechProduct.Title = record.Title;
								objTechProduct.Prerequisites = record.TechnologyPrerequisites;
								//-|Create and Embed the TechnologyVendor
								TechnologyVendor objTechVendor = TechnologyVendor.Read(parIDsp: Convert.ToInt16(record.TechnologyVendorId));
								objTechProduct.Vendor = objTechVendor;
								//-| Create and embed the Technology Category
								TechnologyCategory objTechCategory = TechnologyCategory.Read(parIDsp: Convert.ToInt16(record.TechnologyCategoryId));
								objTechProduct.Category = objTechCategory;

								dbSession.Persist(objTechProduct);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating TechnologyProducts ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with TechnologyProducts");
						}

					stopwatch.Stop();
					Console.Write("\n\t + TechnologyProducts...\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ActivityCategories**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsActivityCategories =
								from dsActivityCategory in parSDDPdatacontext.ActivityCategories
								where dsActivityCategory.Id > lastReadID
								&& dsActivityCategory.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsActivityCategory;

							entriesCounter = 0;

							foreach (ActivityCategoriesItem record in rsActivityCategories)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								ActivityCategory activityCategory = (from theEntry in dbSession.AllObjects<ActivityCategory>()
																	where theEntry.IDsp == record.Id
																	select theEntry).FirstOrDefault();
								if (activityCategory == null)
									activityCategory = new ActivityCategory();

								activityCategory.IDsp = record.Id;
								activityCategory.Title = record.Title;

								dbSession.Persist(activityCategory);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ActivityCategories ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with ActivityCategories");
						}

					stopwatch.Stop();
					Console.Write("\n\t + ActivityCategories...\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ServiceLevelCategories**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsServiceLevelCategories =
								from dsSLcategory in parSDDPdatacontext.ServiceLevelCatagories
								where dsSLcategory.Id > lastReadID
								&& dsSLcategory.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsSLcategory;

							entriesCounter = 0;

							foreach (ServiceLevelCatagoriesItem record in rsServiceLevelCategories)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								ServiceLevelCategory serviceLeveLCategory = (from theEntry in dbSession.AllObjects<ServiceLevelCategory>()
																	where theEntry.IDsp == record.Id
																	select theEntry).FirstOrDefault();
								if (serviceLeveLCategory == null)
									serviceLeveLCategory = new ServiceLevelCategory();

								serviceLeveLCategory.IDsp = record.Id;
								serviceLeveLCategory.Title = record.Title;

								dbSession.Persist(serviceLeveLCategory);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceLevelCategory ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with ServiceLevelCategory");
						}

					stopwatch.Stop();
					Console.Write("\n\t + ServiceLevelCategories...\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ServicePortfolios**
					stopwatch.Restart();
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsPortfolios =
								from dsPortfolio in parSDDPdatacontext.ServicePortfolios
								where dsPortfolio.Id > lastReadID
								&& dsPortfolio.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsPortfolio;

							entriesCounter = 0;

							foreach (var recordPortfolio in rsPortfolios)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = recordPortfolio.Id;

								ServicePortfolio servicePortfolio = (from sp in dbSession.AllObjects<ServicePortfolio>()
																		where sp.IDsp == recordPortfolio.Id
																		select sp).FirstOrDefault();
								if (servicePortfolio == null)
									servicePortfolio = new ServicePortfolio();

								servicePortfolio.IDsp = recordPortfolio.Id;
								servicePortfolio.Title = recordPortfolio.Title;
								servicePortfolio.PortfolioType = recordPortfolio.PortfolioTypeValue;
								servicePortfolio.ISDheading = recordPortfolio.ISDHeading;
								servicePortfolio.ISDdescription = recordPortfolio.ISDDescription;
								servicePortfolio.CSDheading = recordPortfolio.ContractHeading;
								servicePortfolio.CSDdescription = recordPortfolio.CSDDescription;
								servicePortfolio.SOWheading = recordPortfolio.ContractHeading;
								servicePortfolio.SOWdescription = recordPortfolio.ContractDescription;

								dbSession.Persist(servicePortfolio);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServicePortfolios ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local Database with ServicePortfolios");
						}

					stopwatch.Stop();
					Console.Write("\n\t + ServicePortfolios\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ServiceFamilies**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							dbSession.BeginUpdate();
							var rsFamilies = from dsFamily in parSDDPdatacontext.ServiceFamilies
												where dsFamily.Id > lastReadID
												&& dsFamily.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
												select dsFamily;

							entriesCounter = 0;

							foreach (var record in rsFamilies)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								ServiceFamily serviceFamily = (from thisEntry in dbSession.AllObjects<ServiceFamily>()
																where thisEntry.IDsp == record.Id
																select thisEntry).FirstOrDefault();
								if (serviceFamily == null)
									serviceFamily = new ServiceFamily();

								serviceFamily.IDsp = record.Id;
								serviceFamily.Title = record.Title;
								serviceFamily.ServicePortfolioIDsp = record.Service_PortfolioId;
								serviceFamily.ISDheading = record.ISDHeading;
								serviceFamily.ISDdescription = record.ISDDescription;
								serviceFamily.CSDheading = record.ContractHeading;
								serviceFamily.CSDdescription = record.CSDDescription;
								serviceFamily.SOWheading = record.ContractHeading;
								serviceFamily.SOWdescription = record.ContractDescription;

								dbSession.Persist(serviceFamily);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceFamilies ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceFamilies");
						}

					stopwatch.Stop();
					Console.Write("\n\t + ServiceFamilies...\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ServiceProducts**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsProducts =
								from dsProduct in parSDDPdatacontext.ServiceProducts
								where dsProduct.Id > lastReadID
								&& dsProduct.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsProduct;

							entriesCounter = 0;

							foreach (var recordProduct in rsProducts)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = recordProduct.Id;

								ServiceProduct objProduct = (from thisEntry in dbSession.AllObjects<ServiceProduct>()
																where thisEntry.IDsp == recordProduct.Id
																select thisEntry).FirstOrDefault();
								if (objProduct == null)
									objProduct = new ServiceProduct();

								objProduct.IDsp = recordProduct.Id;
								objProduct.Title = recordProduct.Title;
								objProduct.ServiceFamilyIDsp = recordProduct.Service_FamilyId;
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

								dbSession.Persist(objProduct);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceProduct ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceProduct");
						}
					stopwatch.Stop();
					Console.Write("\n\t + ServiceProducts...\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					stopwatch.Restart();

					//---G
					//+ Populate the **ServiceElements**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();

						while (fetchMoreIndicator)
							{
							var rsElements = from dsElement in parSDDPdatacontext.ServiceElements
												where dsElement.Id > lastReadID
												&& dsElement.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
												select dsElement;

							entriesCounter = 0;

							foreach (var record in rsElements)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								ServiceElement objElement = (from thisentry in dbSession.AllObjects<ServiceElement>()
																where thisentry.IDsp == record.Id
																select thisentry).FirstOrDefault();
								if (objElement == null)
									objElement = new ServiceElement();

								objElement.IDsp = record.Id;
								objElement.Title = record.Title;
								objElement.ServiceProductIDsp = record.Service_ProductId;
								objElement.SortOrder = record.SortOrder;
								objElement.ISDheading = record.ISDHeading;
								objElement.ISDdescription = record.ISDDescription;
								objElement.Objectives = record.Objective;
								objElement.KeyClientAdvantages = record.KeyClientAdvantages;
								objElement.KeyClientBenefits = record.KeyClientBenefits;
								objElement.KeyDDbenefits = record.KeyDDBenefits;
								objElement.CriticalSuccessFactors = record.CriticalSuccessFactors;
								objElement.ProcessLink = record.ProcessLink;
								objElement.KeyPerformanceIndicators = record.KeyPerformanceIndicators;
								objElement.ContentLayer = record.ContentLayerValue;
								objElement.ContentPredecessorElementIDsp = record.ContentPredecessorElementId;
								objElement.ContentStatus = record.ContentStatusValue;

								dbSession.Persist(objElement);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceElements ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceElements");
						}
					stopwatch.Stop();
					Console.Write("\n\t + ServiceElements...\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ServiceFeatures**

					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsFeatures = from dsFeature in parSDDPdatacontext.ServiceFeatures
												where dsFeature.Id > lastReadID
												&& dsFeature.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
												select dsFeature;

							entriesCounter = 0;

							foreach (var record in rsFeatures)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								ServiceFeature serviceFeature = (from thisEntry in dbSession.AllObjects<ServiceFeature>()
																	where thisEntry.IDsp == record.Id
																	select thisEntry).FirstOrDefault();
								if(serviceFeature == null)
									serviceFeature = new ServiceFeature();
																	
								serviceFeature.IDsp = record.Id;
								serviceFeature.Title = record.Title;
								serviceFeature.ServiceProductIDsp = Convert.ToInt16(record.Service_ProductId);
								serviceFeature.SortOrder = record.SortOrder;
								serviceFeature.CSDheading = record.ContractHeading;
								serviceFeature.CSDdescription = record.CSDDescription;
								serviceFeature.SOWheading = record.ContractHeading;
								serviceFeature.SOWdescription = record.ContractDescription;
								serviceFeature.ContentLayer = record.ContentLayerValue;
								serviceFeature.ContentPredecessorFeatureIDsp = record.ContentPredecessorFeatureId;
								serviceFeature.ContentStatus = record.ContentStatusValue;

								dbSession.Persist(serviceFeature);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceFeatures ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceFeatures");
						}
					stopwatch.Stop();
					Console.Write("\n\t + ServiceFeatures...\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);


					stopwatch.Restart();
					//---G
					//+Populate **Deliverables**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					try
						{
						dbSession.BeginUpdate();

						var dsDeliverables = parSDDPdatacontext.Deliverables
							.Expand(dlv => dlv.SupportingSystems)
							.Expand(dlv => dlv.GlossaryAndAcronyms)
							.Expand(dlv => dlv.Responsible_RACI)
							.Expand(dlv => dlv.Accountable_RACI)
							.Expand(dlv => dlv.Consulted_RACI)
							.Expand(dlv => dlv.Informed_RACI);

						while (fetchMoreIndicator)
							{
							var rsDeliverables =
								from dsDeliverable in dsDeliverables
								where dsDeliverable.Id > lastReadID
								&& dsDeliverable.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsDeliverable;

							entriesCounter = 0;

							foreach (DeliverablesItem record in rsDeliverables)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								Deliverable deliverable = new Deliverable();

								deliverable = (from thisEntry in dbSession.AllObjects<Deliverable>()
												where thisEntry.IDsp == record.Id
												select thisEntry).FirstOrDefault();
									
								if(deliverable == null)
									deliverable = new Deliverable();

								deliverable.IDsp = record.Id;
								deliverable.Title = record.Title;
								deliverable.DeliverableType = record.DeliverableTypeValue;
								deliverable.SortOrder = record.SortOrder;
								deliverable.ISDheading = record.ISDHeading;
								deliverable.ISDsummary = record.ISDSummary;
								deliverable.ISDdescription = record.ISDDescription;
								deliverable.CSDheading = record.CSDHeading;
								deliverable.CSDsummary = record.CSDSummary;
								deliverable.CSDdescription = record.CSDDescription;
								deliverable.SOWheading = record.ContractHeading;
								deliverable.SOWsummary = record.ContractSummary;
								deliverable.SOWdescription = record.ContractDescription;
								deliverable.TransitionDescription = record.TransitionDescription;
								deliverable.Inputs = record.Inputs;
								deliverable.Outputs = record.Outputs;
								deliverable.DDobligations = record.SPObligations;
								deliverable.ClientResponsibilities = record.ClientResponsibilities;
								deliverable.Exclusions = record.Exclusions;
								deliverable.GovernanceControls = record.GovernanceControls;
								deliverable.WhatHasChanged = record.WhatHasChanged;
								deliverable.ContentStatus = record.ContentStatusValue;
								deliverable.ContentLayer = record.ContentLayerValue;
								deliverable.ContentPredecessorDeliverableIDsp = record.ContentPredecessor_DeliverableId;

								//-| Add the GlossaryAcronym terms to the Deliverable object
								if (record.GlossaryAndAcronyms.Count > 0)
									{
									deliverable.GlossaryAndAcronyms = new List<int>();
									foreach (GlossaryAndAcronymsItem recGlossAcronym in record.GlossaryAndAcronyms)
										{
										if (deliverable.GlossaryAndAcronyms.Contains(recGlossAcronym.Id) == false)
											deliverable.GlossaryAndAcronyms.Add(recGlossAcronym.Id);
										}
									}
								//-| Add the Supporting systems
								if (record.SupportingSystems.Count > 0)
									{
									deliverable.SupportingSystems = new List<string>();
									foreach (var recSupportingSystem in record.SupportingSystems)
										{
										deliverable.SupportingSystems.Add(recSupportingSystem.Value);
										}
									}

								//-| Populate the RACI Lists
								//-| + RACIresponsibles
								if (record.Responsible_RACI.Count > 0)
									{
									deliverable.RACIresponsibles = new List<int>();
									foreach (var recJobRole in record.Responsible_RACI)
										{
										deliverable.RACIresponsibles.Add(recJobRole.Id);
										}
									}

								//-| + RACIaccountables
								if (record.Accountable_RACI != null)
									{
									deliverable.RACIaccountables = new List<int>();
									if (record.Accountable_RACI != null)
										{
										deliverable.RACIaccountables.Add(Convert.ToInt16(record.Accountable_RACIId));
										}
									}

								//-| +RACIconsulteds
								if (record.Consulted_RACI.Count > 0)
									{
									deliverable.RACIconsulteds = new List<int>();
									foreach (var recJobRole in record.Consulted_RACI)
										{
										deliverable.RACIconsulteds.Add(recJobRole.Id);
										}
									}

								//-| + RACIinformeds
								if (record.Informed_RACI.Count > 0)
									{
									deliverable.RACIinformeds = new List<int>();
									foreach (var recJobRole in record.Informed_RACI)
										{
										deliverable.RACIinformeds.Add(recJobRole.Id);
										}
									}

								dbSession.Persist(deliverable);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("...Entries Processed: {0}, Last ID processed: {1} ", entriesCounter, lastReadID);
						Console.WriteLine("### Exception while populating Deliverables ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with Deliverables");
						}

					stopwatch.Stop();
					Console.Write("\n\t + Deliverables...\t\t\t\t {0} \t {1}", entriesCounter.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **Element Deliverables**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsElementDeliverable =
								from dsElementDeliverable in parSDDPdatacontext.ElementDeliverables
								where dsElementDeliverable.Id > lastReadID
								&& dsElementDeliverable.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsElementDeliverable;

							entriesCounter = 0;

							foreach (var record in rsElementDeliverable)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								ElementDeliverable elementDeliverable = (from thisEntry in dbSession.AllObjects<ElementDeliverable>()
																			where thisEntry.IDsp == record.Id
																			select thisEntry).FirstOrDefault();
								if (elementDeliverable == null)
									elementDeliverable = new ElementDeliverable();

								elementDeliverable.IDsp = record.Id;
								elementDeliverable.Title = record.Title;
								elementDeliverable.AssociatedDeliverableIDsp = record.Deliverable_Id;
								elementDeliverable.AssociatedElementIDsp = record.Service_ElementId;
								elementDeliverable.Optionality = record.OptionalityValue;

								dbSession.Persist(elementDeliverable);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ElementDeliverables ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ElementDeliverables");
						}

					stopwatch.Stop();
					Console.Write("\n\t + ElementDeliverables...\t\t {0} \t {1}", entriesCounter.ToString("D3"), stopwatch.Elapsed);

					//+Populate **Feature Deliverables**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsFeatureDeliverable =
								from dsFeatureDeliverable in parSDDPdatacontext.FeatureDeliverables
								where dsFeatureDeliverable.Id > lastReadID
								&& dsFeatureDeliverable.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsFeatureDeliverable;

							entriesCounter = 0;

							foreach (var record in rsFeatureDeliverable)
								{
								entriesCounter += 1;
								totalEntriesUpdated += 1;
								lastReadID = record.Id;

								FeatureDeliverable objFeatureDeliverable = (from theEntry in dbSession.AllObjects<FeatureDeliverable>()
																			where theEntry.IDsp == record.Id
																			select theEntry).FirstOrDefault();
								if (objFeatureDeliverable == null)
									objFeatureDeliverable = new FeatureDeliverable();

								objFeatureDeliverable.IDsp = record.Id;
								objFeatureDeliverable.Title = record.Title;
								objFeatureDeliverable.AssociatedDeliverableIDsp = record.Deliverable_Id;
								objFeatureDeliverable.AssociatedFeatureIDsp = record.Service_FeatureId;
								objFeatureDeliverable.Optionality = record.OptionalityValue;

								dbSession.Persist(objFeatureDeliverable);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating FeatureDeliverables ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local database with FeatureDeliverables");
						}
					stopwatch.Stop();
					Console.Write("\n\t + FeatureDeliverables...\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **DeliverableTechnologies**

					stopwatch.Restart();
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsDeliverableTechnologies =
								from dsDeliverableTechnology in parSDDPdatacontext.DeliverableTechnologies
								where dsDeliverableTechnology.Id > lastReadID
								&& dsDeliverableTechnology.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsDeliverableTechnology;

							entriesCounter = 0;

							foreach (var recDeliverableTechnology in rsDeliverableTechnologies)
								{
								totalEntriesUpdated += 1;
								entriesCounter += 1;
								lastReadID = recDeliverableTechnology.Id;
									
								DeliverableTechnology deliverableTechnology = (from theEntry in dbSession.AllObjects<DeliverableTechnology>()
																				where theEntry.IDsp == recDeliverableTechnology.Id
																				select theEntry).FirstOrDefault();
								if (deliverableTechnology == null)
									deliverableTechnology = new DeliverableTechnology();

								deliverableTechnology.IDsp = recDeliverableTechnology.Id;
								deliverableTechnology.Title = recDeliverableTechnology.Title;
								deliverableTechnology.Considerations = recDeliverableTechnology.TechnologyConsiderations;
								deliverableTechnology.RoadmapStatus = recDeliverableTechnology.TechnologyRoadmapStatusValue;
								deliverableTechnology.AssociatedDeliverableIDsp = recDeliverableTechnology.Deliverable_Id;
								deliverableTechnology.AssociatedTechnologyProductIDsp = recDeliverableTechnology.TechnologyProductsId;

								dbSession.Persist(deliverableTechnology);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceElements ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceElements");
						}

					stopwatch.Stop();
					Console.Write("\n\t + DeliverableTechnologies...\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);


					//---g
					//+ Populate **Activities**
					stopwatch.Restart();
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					try
						{
						dbSession.BeginUpdate();

						var dsActivities = parSDDPdatacontext.Activities
							.Expand(ac => ac.Activity_Category)
							.Expand(ac => ac.OLA_);

						while (fetchMoreIndicator)
							{
							var rsActivities =
								from dsActivity in dsActivities
								where dsActivity.Id > lastReadID
								&& dsActivity.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsActivity;

							entriesCounter = 0;

							foreach (ActivitiesItem record in rsActivities)
								{
								totalEntriesUpdated += 1;
								entriesCounter += 1;
								lastReadID = record.Id;

								Activity activity = (from theEntry in dbSession.AllObjects<Activity>()
														where theEntry.IDsp == record.Id
														select theEntry).FirstOrDefault();
								if (activity == null)
									activity = new Activity();

								activity.IDsp = record.Id;
								activity.Title = record.Title;
								activity.SortOrder = record.SortOrder;
								activity.Category = record.Activity_Category.Title;
								activity.Assumptions = record.ActivityAssumptions;
								activity.ContentStatus = record.ContentStatusValue;
								activity.ISDheading = record.ISDHeading;
								activity.ISDdescription = record.ISDDescription;
								activity.Inputs = record.ActivityInput;
								activity.Outputs = record.ActivityOutput;
								activity.CSDheading = record.CSDHeading;
								activity.CSDdescription = record.CSDDescription;
								activity.SOWheading = record.CSDDescription;
								if (record.OLA_ != null)
									activity.OLA = record.OLA_.Title;
								activity.OLAvariations = record.OLAVariations;
								activity.Optionality = record.ActivityOptionalityValue;
								activity.OwningEntity = record.OwningEntityValue;
								if (record.Accountable_RACI != null)
									{
									activity.RACIaccountables = new List<int?>();
									activity.RACIaccountables.Add(record.Accountable_RACIId);
									}
								if (record.Responsible_RACI != null && record.Responsible_RACI.Count() > 0)
									{
									activity.RACIresponsibles = new List<int?>();
									foreach (var entryJobRole in record.Responsible_RACI)
										{
										activity.RACIresponsibles.Add(entryJobRole.Id);
										}
									}
								if (record.Consulted_RACI != null && record.Consulted_RACI.Count() > 0)
									{
									activity.RACIconsulteds = new List<int?>();
									foreach (var entryJobRole in record.Consulted_RACI)
										{
										activity.RACIconsulteds.Add(record.Id);
										}
									}
								if (record.Informed_RACI != null && record.Informed_RACI.Count() > 0)
									{
									activity.RACIinformeds = new List<int?>();
									foreach (var entryJobRole in record.Informed_RACI)
										{
										activity.RACIinformeds.Add(record.Id);
										}
									}

								dbSession.Persist(activity);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating Activities ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with Activities");
						}
					stopwatch.Stop();
					Console.Write("\n\t + Activities...\t\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---g
					//+ Populate **DeliverableActivities**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					stopwatch.Restart();

					try
						{
						dbSession.BeginUpdate();
						while (fetchMoreIndicator)
							{
							var rsDeliverableActivities =
								from dsDeliverableActivity in parSDDPdatacontext.DeliverableActivities
								where dsDeliverableActivity.Id > lastReadID
								&& dsDeliverableActivity.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsDeliverableActivity;

							entriesCounter = 0;

							foreach (var record in rsDeliverableActivities)
								{
								totalEntriesUpdated += 1;
								lastReadID = record.Id;
								entriesCounter += 1;
									
								DeliverableActivity deliverableActivity = (from theEntry in dbSession.AllObjects<DeliverableActivity>()
																			where theEntry.IDsp == record.Id
																			select theEntry).FirstOrDefault();
								if(deliverableActivity == null)
									deliverableActivity = new DeliverableActivity();
																		
								deliverableActivity.IDsp = record.Id;
								deliverableActivity.Title = record.Title;
								deliverableActivity.Optionality = record.OptionalityValue;
								deliverableActivity.AssociatedActivityIDsp = record.Activity_Id;
								deliverableActivity.AssociatedDeliverableIDsp = record.Deliverable_Id;

								dbSession.Persist(deliverableActivity);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceElements ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceElements");
						}

					stopwatch.Stop();
					Console.Write("\n\t + DeliverableActivities\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **ServiceLevels**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					stopwatch.Restart();
					try
						{
						dbSession.BeginUpdate();
						var datasetServiceLevels = parSDDPdatacontext.ServiceLevels
							.Expand(sl => sl.Service_Hour);

						while (fetchMoreIndicator)
							{
							var rsServiceLevels = from dsServiceLevel in datasetServiceLevels
								where dsServiceLevel.Id > lastReadID
								&& dsServiceLevel.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsServiceLevel;

							entriesCounter = 0;

							foreach (ServiceLevelsItem record in rsServiceLevels)
								{
								totalEntriesUpdated += 1;
								entriesCounter += 1;
								lastReadID = record.Id;

								ServiceLevel objServiceLevel = (from theEntry in dbSession.AllObjects<ServiceLevel>()
																where theEntry.IDsp == record.Id
																select theEntry).FirstOrDefault();
								if(objServiceLevel == null)
									objServiceLevel = new ServiceLevel();

								objServiceLevel.IDsp = record.Id;
								objServiceLevel.Title = record.Title;
								objServiceLevel.CategoryIDsp =  Convert.ToInt16(record.Service_Level_CategoryId);
								objServiceLevel.ISDheading = record.ISDHeading;
								objServiceLevel.ISDdescription = record.ISDDescription;
								objServiceLevel.CSDheading = record.CSDHeading;
								objServiceLevel.CSDdescription = record.CSDDescription;
								objServiceLevel.BasicConditions = record.BasicServiceLevelConditions;
								objServiceLevel.CalculationMethod = record.CalculationMethod;
								objServiceLevel.CalculationFormula = record.CalculationFormula;
								objServiceLevel.ContentStatus = record.ContentStatusValue;
								objServiceLevel.Measurement = record.ServiceLevelMeasurement;
								objServiceLevel.MeasurementInterval = record.MeasurementIntervalValue;
								objServiceLevel.SOWheading = record.ContractHeading;
								objServiceLevel.SOWdescription = record.ContractDescription;
								objServiceLevel.ReportingInterval = record.ReportingIntervalValue;
								if (record.Service_HourId != null)
									objServiceLevel.ServiceHours = record.Service_Hour.Title;
								objServiceLevel.BasicConditions = record.BasicServiceLevelConditions;

								//-| Load the Service Level Performance Thresholds
								var dsThresholds =
									from dsThreshold in parSDDPdatacontext.ServiceLevelTargets
									where dsThreshold.Service_LevelId == record.Id && dsThreshold.ThresholdOrTargetValue == "Threshold"
									orderby dsThreshold.Title
									select dsThreshold;

								if (dsThresholds.Count() > 0)
									{
									objServiceLevel.PerformanceThresholds = new List<ServiceLevelTarget>();
									foreach (var thresholdItem in dsThresholds)
										{
										ServiceLevelTarget objSLthreshold = new ServiceLevelTarget();
										objSLthreshold.IDsp = thresholdItem.Id;
										objSLthreshold.Title = thresholdItem.Title.Substring(thresholdItem.Title.IndexOf(": ", 0) + 2,
											thresholdItem.Title.Length - thresholdItem.Title.IndexOf(": ", 0) - 2);
										objSLthreshold.Type = thresholdItem.ThresholdOrTargetValue;
										objSLthreshold.ContentStatus = thresholdItem.ContentStatusValue;
										objServiceLevel.PerformanceThresholds.Add(objSLthreshold);
										}
									}
								//-| Load the Service Level Performance Targets
								var dsTargets =
									from dsThreshold in parSDDPdatacontext.ServiceLevelTargets
									where dsThreshold.Service_LevelId == record.Id && dsThreshold.ThresholdOrTargetValue == "Target"
									orderby dsThreshold.Title
									select dsThreshold;

								if (dsTargets.Count() > 0)
									{
									objServiceLevel.PerformanceTargets = new List<ServiceLevelTarget>();
									foreach (var targetEntry in dsTargets)
										{
										ServiceLevelTarget objSLtarget = new ServiceLevelTarget();
										objSLtarget.IDsp = targetEntry.Id;
										objSLtarget.Title = targetEntry.Title.Substring(targetEntry.Title.IndexOf(": ", 0) + 2,
											(targetEntry.Title.Length - targetEntry.Title.IndexOf(": ", 0) - 2));
										objSLtarget.Type = targetEntry.ThresholdOrTargetValue;
										objSLtarget.ContentStatus = targetEntry.ContentStatusValue;
										objServiceLevel.PerformanceTargets.Add(objSLtarget);
										}
									}

								dbSession.Persist(objServiceLevel);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating ServiceLevels ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with ServiceLevels");
						}

					stopwatch.Stop();
					Console.Write("\n\t + ServiceLevels...\t\t\t\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);

					//---G
					//+ Populate **DeliverableServiceLevels**
					lastReadID = 0;
					stopwatch.Restart();
					fetchMoreIndicator = true;
					totalEntriesUpdated = 0;
					stopwatch.Restart();

					try
						{
						dbSession.BeginUpdate();

						while (fetchMoreIndicator)
							{
							var rsDeliverableServiceLevels =
								from dsDeliverableServiceLevel in parSDDPdatacontext.DeliverableServiceLevels
								where dsDeliverableServiceLevel.Id > lastReadID
								&& dsDeliverableServiceLevel.Modified > Properties.Settings.Default.CurrentDatabaseLastRefreshedOn
								select dsDeliverableServiceLevel;

							entriesCounter = 0;

							foreach (var record in rsDeliverableServiceLevels)
								{
								totalEntriesUpdated += 1;
								lastReadID = record.Id;
								entriesCounter += 1;

								DeliverableServiceLevel deliverableServiceLevel = (from theEntry in dbSession.AllObjects<DeliverableServiceLevel>()
																						where theEntry.IDsp == record.Id
																						select theEntry).FirstOrDefault();
								if(deliverableServiceLevel == null)
									deliverableServiceLevel = new DeliverableServiceLevel();
																		
								deliverableServiceLevel.IDsp = record.Id;
								deliverableServiceLevel.Title = record.Title;
								deliverableServiceLevel.Optionality = record.OptionalityValue;
								deliverableServiceLevel.ContentStatus = record.ContentStatusValue;
								deliverableServiceLevel.AdditionalConditions = record.AdditionalConditions;
								deliverableServiceLevel.AssociatedDeliverableIDsp = record.Deliverable_Id;
								deliverableServiceLevel.AssociatedServiceLevelIDsp = record.Service_LevelId;
								deliverableServiceLevel.AssociatedServiceProductIDsp = record.Service_ProductId;

								dbSession.Persist(deliverableServiceLevel);
								}
							if (entriesCounter < 1000)
								break;
							}
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating DeliverableServiceLevel ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with DeliverbableServiceLevel");
						}
					stopwatch.Stop();
					Console.Write("\n\t + DeliverableServiceLevels...\t {0} \t {1}", totalEntriesUpdated.ToString("D3"), stopwatch.Elapsed);
					}

				//- -----------------------------------------------------------------------------------------------------------------
				//-|Setting the **CurrentdatabaseIsPopulated**
				//- ------------------------------------------------------------------------------------------------------------------
				Properties.Settings.Default.CurrentDatabaseIsPopulated = true;
				stopwatchCompleteDataSet.Stop();
				Console.Write("\n\nPopulating the Local {0} Database took {1}", Properties.Settings.Default.CurrentPlatform, stopwatchCompleteDataSet.Elapsed);
				//-| This is the **Main ** which does not terminate, it returns to the calling module where it continues with execution...
				}
			catch (DataServiceClientException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nStatusCode: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.StatusCode, exc.StackTrace);
				Properties.Settings.Default.CurrentDatabaseIsPopulated = false;
				}
			catch (DataServiceQueryException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nResponse: {2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				Properties.Settings.Default.CurrentDatabaseIsPopulated = false;
				}
			catch (DataServiceTransportException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nResponse:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				Properties.Settings.Default.CurrentDatabaseIsPopulated = false;
				}
			catch (System.Net.Sockets.SocketException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nTargetSite:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.TargetSite, exc.StackTrace);
				Properties.Settings.Default.CurrentDatabaseIsPopulated = false;
				}
			catch (LocalDatabaseExeption exc)
				{
				Console.Write("\n\n*** Exception from Local Database ***\n{0} - {1}\nSource:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Source, exc.StackTrace);
				Properties.Settings.Default.CurrentDatabaseIsPopulated = false;
				}
			catch (Exception exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1}\nSource:{2}\nStackTrace: {3}.", exc.HResult, exc.Message, exc.Source, exc.StackTrace);
				Properties.Settings.Default.CurrentDatabaseIsPopulated = false;
				}
			return;
			}


		//===G
		//++UpdateMappingData
		//---G
		public static bool UpdateMappingData(
			DesignAndDeliveryPortfolioDataContext parDatacontexSDDP,
			int parMapping)
			{
			int totalEntriesRead = 0;
			int lastReadEntryID = 0;
			int entryReadCounter = 0;
			bool fetchMoreEntries = true;
			bool result = false;
			bool deleteResult = false;
			//-| **Please Note:**
			//-| SharePoint's REST API has a limit which returns only 1000 entries at a time
			//-| therefore a paging mechanism is implemented to return all the entries in the List.
			Properties.Settings.Default.CurrentMappingIsPopulated = false;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(
						systemHost: Properties.Settings.Default.CurrentDatabaseHost,
						systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{

					Console.Write("\nPopulating the complete Mappings DataSet...");
					deleteResult = Mapping.DeleteAll();
					if(!deleteResult)
						{
						Console.WriteLine("### Exception while deleting Mapping from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the Mappings from the local databse.");
						}

					//+ Populate Mapping
					dbSession.BeginUpdate();
					Console.Write("\n\t + Mappings...");
					Stopwatch stopwatchOverall = Stopwatch.StartNew();
					Stopwatch stopwatchIndividual = Stopwatch.StartNew();

					var datasetMappings = parDatacontexSDDP.Mappings
						.Expand(m => m.Client_);

					var rsMappings =
						from dsMapping in datasetMappings
						where dsMapping.Id == parMapping
						select dsMapping;

					MappingsItem record = rsMappings.FirstOrDefault();

					if (record == null)
						{
						dbSession.Abort();
						result = false;
						goto Exit_Method;
						}

					lastReadEntryID = record.Id;
					totalEntriesRead += 1;
					entryReadCounter += 1;

					try
						{
						//-|Check if the Mapping exist in the Local database
						Mapping mapping = (from theEntry in dbSession.AllObjects<Mapping>()
										   where theEntry.IDsp == parMapping
										   select theEntry).FirstOrDefault();
						if (mapping == null)
							mapping = new Mapping();

						mapping.IDsp = record.Id;
						mapping.Title = record.Title;
						mapping.ClientName = record.Client_.DocGenClientName;
						dbSession.Persist(mapping);
						dbSession.Commit();
						}
					catch (Exception exc)
						{
						Console.WriteLine("### Exception while populating Mapping ### " + exc.HResult + " - " + exc.Message);
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while populating local databse with DeliverbableServiceLevel");
						}

					stopwatchIndividual.Stop();
					Console.Write("\t\t\t\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);

					//+ Populate **MappingServiceTowers**
					Console.Write("\n\t + MappingServiceTowers...");
					stopwatchIndividual.Restart();
					totalEntriesRead = 0;
					entryReadCounter = 0;
					lastReadEntryID = 0;

					deleteResult = MappingServiceTower.DeleteAll();
					if (!deleteResult)
						{
						Console.WriteLine("### Exception while deleting MappingServiceTowers from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the MappingServiceTowers from the local databse.");
						}

					dbSession.BeginUpdate();

					//-|Store all the MappingServiceTower IDsp values in a List to process the next layer...
					List<int> serviceTowerIDs = new List<int>();

					while(fetchMoreEntries)
						{
						var rsMappingServiceTowers =
							from dsMappingServiceTowers in parDatacontexSDDP.MappingServiceTowers
							where dsMappingServiceTowers.Mapping_Id == parMapping
							&& dsMappingServiceTowers.Id > lastReadEntryID
							select dsMappingServiceTowers;

						entryReadCounter = 0;

						foreach (var recordTower in rsMappingServiceTowers)
							{
							lastReadEntryID = recordTower.Id;
							serviceTowerIDs.Add(recordTower.Id);
							entryReadCounter += 1;
							totalEntriesRead += 1;

							MappingServiceTower objMappingServiceTower = (from theEntry in dbSession.AllObjects<MappingServiceTower>()
																		  where theEntry.IDsp == recordTower.Id
																		  select theEntry).FirstOrDefault();
							if(objMappingServiceTower == null)
								objMappingServiceTower = new MappingServiceTower();

							objMappingServiceTower.IDsp = recordTower.Id;
							objMappingServiceTower.Title = recordTower.Title;
							objMappingServiceTower.MappingIDsp = recordTower.Mapping_Id;
							dbSession.Persist(objMappingServiceTower);
							}
						if (entryReadCounter < 1000)
							break;
						}
					dbSession.Commit();
					stopwatchIndividual.Stop();
					Console.Write("\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);

					//+ Populate **MappingRequirements**
					Console.Write("\n\t + MappingRequirements...");
					stopwatchIndividual.Restart();
					totalEntriesRead = 0;
					List<int> requirementIDs = new List<int>();

					deleteResult = MappingRequirement.DeleteAll();
					if (!deleteResult)
						{
						Console.WriteLine("### Exception while deleting MappingRequirements from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the MappingRequirements from the local databse.");
						}

					dbSession.BeginUpdate();
					//-|Process all the Requirements for each MappingServiceTower IDsp in the list...
					foreach (int serviceTowerID in serviceTowerIDs)
						{
						entryReadCounter = 0;
						lastReadEntryID = 0;

						while (fetchMoreEntries)
							{
							var rsMappingRequirements =
								from dsMappingRequirements in parDatacontexSDDP.MappingRequirements
								where dsMappingRequirements.Mapping_TowerId == serviceTowerID
								select dsMappingRequirements;

							entryReadCounter = 0;

							foreach (var recordMR in rsMappingRequirements)
								{
								entryReadCounter += 1;
								totalEntriesRead += 1;
								lastReadEntryID = recordMR.Id;
								requirementIDs.Add(recordMR.Id);

								MappingRequirement objMappingRequirement = (from theEntry in dbSession.AllObjects<MappingRequirement>()
																			where theEntry.IDsp == recordMR.Id
																			select theEntry).FirstOrDefault();
								if(objMappingRequirement == null)	
									objMappingRequirement = new MappingRequirement();

								objMappingRequirement.IDsp = recordMR.Id;
								objMappingRequirement.Title = recordMR.Title;
								objMappingRequirement.MappingServiceTowerIDsp = recordMR.Mapping_TowerId;
								objMappingRequirement.ComplianceComments = recordMR.ComplianceComments;
								objMappingRequirement.ComplianceStatus = recordMR.ComplianceStatusValue;
								objMappingRequirement.RequirementServiceLevel = recordMR.RequirementServiceLevel;
								objMappingRequirement.RequirementText = recordMR.RequirementText;
								objMappingRequirement.SourceReference = recordMR.SourceReference;
								objMappingRequirement.SortOrder = recordMR.SortOrder;
								dbSession.Persist(objMappingRequirement);
								}
							if (entryReadCounter < 1000)
								break;
							}
						}
					dbSession.Commit();
					stopwatchIndividual.Stop();
					Console.Write("\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);

					//+ Populate MappingAssumptions
					Console.Write("\n\t + MappingAssumptions...");
					stopwatchIndividual.Restart();
					totalEntriesRead = 0;

					deleteResult = MappingAssumption.DeleteAll();
					if (!deleteResult)
						{
						Console.WriteLine("### Exception while deleting MappingAssumptions from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the MappingAssumptions from the local databse.");
						}

					dbSession.BeginUpdate();
					//-|Process the MappingAssumptions for all the Requirements in the List.
					foreach (int requirementID in requirementIDs)
						{
						entryReadCounter = 0;
						lastReadEntryID = 0;

						while (fetchMoreEntries)
							{
							var rsMappingAssumptions =
								from dsMappingAssumptions in parDatacontexSDDP.MappingAssumptions
								where dsMappingAssumptions.Mapping_RequirementId == requirementID
								select dsMappingAssumptions;

							entryReadCounter = 0;

							foreach (var recordMA in rsMappingAssumptions)
								{
								entryReadCounter += 1;
								totalEntriesRead += 1;
								lastReadEntryID = recordMA.Id;

								MappingAssumption objMappingAssumption = (from theEntry in dbSession.AllObjects<MappingAssumption>()
																		  where theEntry.IDsp == recordMA.Id
																		  select theEntry).FirstOrDefault();
								if(objMappingAssumption == null)
									objMappingAssumption = new MappingAssumption();

								objMappingAssumption.IDsp = recordMA.Id;
								objMappingAssumption.MappingRequirementIDsp = recordMA.Mapping_RequirementId;
								objMappingAssumption.Title = recordMA.Title;
								objMappingAssumption.Description = recordMA.AssumptionDescription;
								dbSession.Persist(objMappingAssumption);
								}
							if (entryReadCounter < 1000)
								break;
							} 
						}
					dbSession.Commit();
					stopwatchIndividual.Stop();
					Console.Write("\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);

					//+ Populate MappingRisks
					Console.Write("\n\t + MappingRisks...");
					stopwatchIndividual.Restart();
					totalEntriesRead = 0;

					deleteResult = MappingRisk.DeleteAll();
					if (!deleteResult)
						{
						Console.WriteLine("### Exception while deleting MappingRisks from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the MappingRisks from the local databse.");
						}

					dbSession.BeginUpdate();
					//-|Process the MappingRisks for all the Requirements in the List.
					foreach (int requirementID in requirementIDs)
						{
						entryReadCounter = 0;
						lastReadEntryID = 0;

						while (fetchMoreEntries)
							{
							var rsMappingRisks =
								from dsMappingRisks in parDatacontexSDDP.MappingRisks
								where dsMappingRisks.Mapping_RequirementId == requirementID
								select dsMappingRisks;

							entryReadCounter = 0;

							foreach (var recordRisk in rsMappingRisks)
								{
								entryReadCounter += 1;
								totalEntriesRead += 1;
								lastReadEntryID = recordRisk.Id;

								MappingRisk objMappingRisk = (from theEntry in dbSession.AllObjects<MappingRisk>()
															  where theEntry.IDsp == recordRisk.Id
															  select theEntry).FirstOrDefault();
								if(objMappingRisk == null)
									objMappingRisk = new MappingRisk();

								objMappingRisk.IDsp = recordRisk.Id;
								objMappingRisk.MappingRequirementIDsp = recordRisk.Mapping_RequirementId;
								objMappingRisk.Title = recordRisk.Title;
								objMappingRisk.Statement = recordRisk.RiskStatement;
								objMappingRisk.Status = recordRisk.RiskStatusValue;
								objMappingRisk.Mittigation = recordRisk.RiskMitigation;
								objMappingRisk.Exposure = recordRisk.RiskExposureValue;
								objMappingRisk.ExposureValue = recordRisk.RiskExposureValue0;
								dbSession.Persist(objMappingRisk);
								}
							if (entryReadCounter < 1000)
								break;
							}
						}
					dbSession.Commit();
					stopwatchIndividual.Stop();
					Console.Write("\t\t\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);

					//+ Populate **MappingDeliverables**
					Console.Write("\n\t + Mapping Deliverables...");
					stopwatchIndividual.Restart();
					totalEntriesRead = 0;

					deleteResult = MappingDeliverable.DeleteAll();
					if (!deleteResult)
						{
						Console.WriteLine("### Exception while deleting MappingDeliverables from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the MappingDeliverables from the local databse.");
						}

					List<int> mappingDeliverableIDs = new List<int>();

					dbSession.BeginUpdate();
					//-|Process the MappingAssumptions for all the Requirements in the List.
					foreach (int requirementID in requirementIDs)
						{
						entryReadCounter = 0;
						lastReadEntryID = 0;

						while (fetchMoreEntries)
							{
							var rsMappingDeliverables =
								from dsMappingDeliverable in parDatacontexSDDP.MappingDeliverables
								where dsMappingDeliverable.Mapping_RequirementId == requirementID
								select dsMappingDeliverable;

							entryReadCounter = 0;

							foreach (var recordMD in rsMappingDeliverables)
								{
								entryReadCounter += 1;
								totalEntriesRead += 1;
								lastReadEntryID = recordMD.Id;
								mappingDeliverableIDs.Add(recordMD.Id);

								MappingDeliverable objMappingDeliverable = (from theEntry in dbSession.AllObjects<MappingDeliverable>()
																			where theEntry.IDsp == recordMD.Id
																			select theEntry).FirstOrDefault();
								if(objMappingDeliverable == null)
									objMappingDeliverable = new MappingDeliverable();

								objMappingDeliverable.IDsp = recordMD.Id;
								objMappingDeliverable.MappingRequirementIDsp = recordMD.Mapping_RequirementId;
								objMappingDeliverable.Title = recordMD.Title;
								if (recordMD.DeliverableChoiceValue == "New")
									objMappingDeliverable.NewDeliverable = true;
								else
									objMappingDeliverable.NewDeliverable = false;
								objMappingDeliverable.MappedDeliverableID = recordMD.Mapped_DeliverableId;
								objMappingDeliverable.NewRequirement = recordMD.DeliverableRequirement;
								objMappingDeliverable.ComplianceComments = recordMD.ComplianceComments;
								dbSession.Persist(objMappingDeliverable);
								}
							if (entryReadCounter < 1000)
								break;
							} 
						}
					dbSession.Commit();
					stopwatchIndividual.Stop();
					Console.Write("\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);

					//+ Populate **MappingServiceLevels**
					Console.Write("\n\t + MappingServiceLevels");
					stopwatchIndividual.Restart();
					totalEntriesRead = 0;

					deleteResult = MappingServiceLevel.DeleteAll();
					if (!deleteResult)
						{
						Console.WriteLine("### Exception while deleting MappingServiceLevels from local database ### ");
						dbSession.Abort();
						throw new LocalDatabaseExeption(message: "Error while deleting the MappingServiceLevels from the local databse.");
						}

					dbSession.BeginUpdate();
					//-|Process the MappingServiceLevels for all the MappingDeliverables in the List.
					foreach (int mappingDeliverableID in mappingDeliverableIDs)
						{
						entryReadCounter = 0;
						lastReadEntryID = 0;

						while (fetchMoreEntries)
							{
							var rsMappingServiceLevels =
								from dsMappingServiceLevel in parDatacontexSDDP.MappingServiceLevels
								where dsMappingServiceLevel.Mapping_DeliverableId == mappingDeliverableID
								select dsMappingServiceLevel;

							entryReadCounter = 0;

							foreach (var recordMSL in rsMappingServiceLevels)
								{
								entryReadCounter += 1;
								totalEntriesRead += 1;
								lastReadEntryID = recordMSL.Id;

								MappingServiceLevel objMappingServiceLevel = (from theEntry in dbSession.AllObjects<MappingServiceLevel>()
																			  where theEntry.IDsp == recordMSL.Id
																			  select theEntry).FirstOrDefault();
								if(objMappingServiceLevel == null)
									objMappingServiceLevel = new MappingServiceLevel();

								objMappingServiceLevel.IDsp = recordMSL.Id;
								objMappingServiceLevel.Title = recordMSL.Title;
								objMappingServiceLevel.MappingDeliverableIDsp = recordMSL.Mapping_DeliverableId;
								objMappingServiceLevel.NewServiceLevel = recordMSL.NewServiceLevel;
								objMappingServiceLevel.MappedServiceLevelIDsp = recordMSL.Service_LevelId;
								objMappingServiceLevel.RequirementText = recordMSL.ServiceLevelRequirement;
								dbSession.Persist(objMappingServiceLevel);
								}
							if (entryReadCounter < 1000)
								break;
							}
						}
					dbSession.Commit();
					stopwatchIndividual.Stop();
					Console.Write("\t\t {0} \t {1}", totalEntriesRead.ToString("D3"), stopwatchIndividual.Elapsed);
					stopwatchOverall.Stop();
					Console.Write("\n\n\tPopulating the Mappings DataSet ended at {0} and took {1}.", DateTime.Now, stopwatchOverall.Elapsed);
					
					}
				Properties.Settings.Default.CurrentMappingIsPopulated = true;
				return true;
				}
			catch (DataServiceClientException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} - StatusCode:{2}\n{3}.", exc.HResult, exc.Message, exc.StatusCode, exc.StackTrace);
				Properties.Settings.Default.CurrentMappingIsPopulated = false;
				}
			catch (DataServiceQueryException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} - StatusCode:{2}\n{3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				Properties.Settings.Default.CurrentMappingIsPopulated = false;
				}
			catch (DataServiceTransportException exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} \n{3}.", exc.HResult, exc.Message, exc.Response, exc.StackTrace);
				Properties.Settings.Default.CurrentMappingIsPopulated = false;
				}
			catch (LocalDatabaseExeption exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} \n{3}.", exc.HResult, exc.Message);
				Properties.Settings.Default.CurrentMappingIsPopulated = false;
				}
			catch (Exception exc)
				{
				Console.Write("\n\n*** Exception ERROR ***\n{0} - {1} \n{3}.", exc.HResult, exc.Message, exc.StackTrace);
				Properties.Settings.Default.CurrentMappingIsPopulated = false;
				}

Exit_Method:
			return result;
			}
		#endregion
		
		}
	
	}
