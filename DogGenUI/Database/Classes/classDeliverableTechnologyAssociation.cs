using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class DeliverableTechnology : OptimizedPersistable
		{
		#region Properties
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}

		private string _Title;
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}

		private string _Considerations;
		public string Considerations {
			get { return this._Considerations; }
			set { UpdateNonIndexField(); this._Considerations = value; }
			}

		private string _RoadmapStatus;
		public string RoadmapStatus {
			get { return this._RoadmapStatus; }
			set { UpdateNonIndexField(); this._RoadmapStatus = value;
				}
			}

		[Index]
		private int? _AssociatedDeliverableIDsp;
		public int? AssociatedDeliverableIDsp {
			get { return this._AssociatedDeliverableIDsp; }
			set { Update(); this._AssociatedDeliverableIDsp = value; }
			}

		[Index]
		private int? _AssociatedTechnologyProductIDsp;
		public int? AssociatedTechnologyProductIDsp {
			get { return this._AssociatedTechnologyProductIDsp; }
			set { Update(); this._AssociatedTechnologyProductIDsp = value; }
			}
		#endregion

		#region Methods
		//---G
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			string parConsiderations,
			string parRoadmapStatus,
			int parAssociatedDeliverableIDsp,
			int parAssociatedTechnologyProductIDsp )

			{
			DeliverableTechnology newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<DeliverableTechnology>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new DeliverableTechnology();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.Considerations  = parConsiderations;
					newEntry.RoadmapStatus = parRoadmapStatus;
					newEntry.AssociatedTechnologyProductIDsp = parAssociatedTechnologyProductIDsp;
					newEntry.AssociatedDeliverableIDsp = parAssociatedDeliverableIDsp;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting DeliverableTechnology ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}
		//---G
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>Deliverable object is retrieved if it exist, else null is retured.</returns>
		public static DeliverableTechnology Read(int parIDsp)
			{
			DeliverableTechnology result = new DeliverableTechnology();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<DeliverableTechnology>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading Deliverable [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//---g
		//++GetDeliverablesForTechnologyProduct
		/// <summary>
		/// Read/retrieve all the deliverable entries that are associated with a specific TechnologyProduct. 
		/// Specify a parameter containing the SharePoint ID value of all the TechnologyProduct for which the deliverables are required.
		/// </summary>
		/// <param name="parTechnologyProductIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple<DeliverableTechnology, Deliverrable>> is retrurned.</returns>
		public static List<Tuple<DeliverableTechnology, Deliverable>> GetDeliverablesForTechnologyProduct(int parTechnologyProductIDsp)
			{
			List<Tuple<DeliverableTechnology, Deliverable>> results = new List<Tuple<DeliverableTechnology, Deliverable>>();
			if (parTechnologyProductIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the DeliverableTechnology objects with which the specified Service Element (parElementIDsp) is associated 
					var technologyProdDeliverables = from eld in dbSession.AllObjects<DeliverableTechnology>()
											  where eld.AssociatedTechnologyProductIDsp == parTechnologyProductIDsp select eld;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in technologyProdDeliverables)
						{
						Deliverable deliverableEntry = (from entry in dbSession.AllObjects<Deliverable>()
														where entry.IDsp == item.AssociatedDeliverableIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (deliverableEntry != null)
							{
							Tuple<DeliverableTechnology, Deliverable> result = new Tuple<DeliverableTechnology, Deliverable>
								(item, deliverableEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all DeliverableTechnologyProduct ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		//---g
		//++GetTechnologyProductsForDeliverable
		/// <summary>
		/// Read/retrieve all the Service Element entries that are associated with a specific Deliverable. 
		/// Specify a parameter containing the SharePoint ID value of all the Deliverable for which the Service Elements are required.
		/// </summary>
		/// <param name="parDeliverableIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple<int, DeliverableTechnology>> is retrurned.</returns>
		public static List<Tuple<DeliverableTechnology, TechnologyProduct>> GetTechnologyProductForDeliverable(int parDeliverableIDsp)
			{
			List<Tuple<DeliverableTechnology, TechnologyProduct>> results = new List<Tuple<DeliverableTechnology,TechnologyProduct>>();

			if (parDeliverableIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the DeliverableTechnology objects with which the specified Deliverable (parDeliverableIDsp) is associated 
					var deliverableTechnologyProducts = from deliverableTP in dbSession.AllObjects<DeliverableTechnology>()
											  where deliverableTP.AssociatedDeliverableIDsp == parDeliverableIDsp
											  select deliverableTP;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in deliverableTechnologyProducts)
						{
						TechnologyProduct technologyProductEntry = (from entry in dbSession.AllObjects<TechnologyProduct>()
														where entry.IDsp == item.AssociatedTechnologyProductIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (technologyProductEntry != null)
							{
							Tuple<DeliverableTechnology, TechnologyProduct> result = new Tuple<DeliverableTechnology, TechnologyProduct>
								(item, technologyProductEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all DeliverableTechnology ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		#endregion
		}
	}
