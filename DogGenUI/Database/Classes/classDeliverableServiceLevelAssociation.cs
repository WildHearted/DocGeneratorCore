using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Collection;
using VelocityDb.Collection.BTree;
using VelocityDb.Indexing;
using VelocityDb.Session;
using VelocityDb.TypeInfo;
using VelocityDBExtensions;

namespace DocGeneratorCore.Database.Classes
	{
	public class DeliverableServiceLevel : OptimizedPersistable
		{
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		private string _ContentStatus;
		private string _Optionality;
		private string _AdditionalConditions;
		[Index]
		private int? _AssociatedDeliverableIDsp;
		[Index]
		private int? _AssociatedServiceLevelIDsp;
		[Index]
		private int? _AssociatedServiceProductIDsp;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value; }
			}
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}
		public string ContentStatus {
			get { return this._ContentStatus; }
			set { UpdateNonIndexField(); this._ContentStatus = value; }
			}
		public string Optionality {
			get { return this._Optionality; }
			set { UpdateNonIndexField(); this._Optionality = value; }
			}
		public string AdditionalConditions {
			get { return this._AdditionalConditions; }
			set { UpdateNonIndexField(); this._AdditionalConditions = value; }
			}
		public int? AssociatedDeliverableIDsp {
			get { return this._AssociatedDeliverableIDsp; }
			set { Update(); this._AssociatedDeliverableIDsp = value; }
			}
		public int? AssociatedServiceLevelIDsp {
			get { return this._AssociatedServiceLevelIDsp; }
			set { Update(); this._AssociatedServiceLevelIDsp = value; }
			}
		public int? AssociatedServiceProductIDsp {
			get { return this._AssociatedServiceLevelIDsp; }
			set { Update(); this._AssociatedServiceLevelIDsp = value; }
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
			string parContentStatus,
			string parOptionality,
			string parAdditionalConditions,
			int parAssociatedDeliverableIDsp,
			int parAssociatedServiceLevelIDsp,
			int parAssociatedServiceProductIDsp)

			{
			DeliverableServiceLevel newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<DeliverableServiceLevel>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new DeliverableServiceLevel();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.ContentStatus  = parContentStatus;
					newEntry.Optionality = parOptionality;
					newEntry.AdditionalConditions = parAdditionalConditions;
					newEntry.AssociatedServiceLevelIDsp = parAssociatedServiceLevelIDsp;
					newEntry.AssociatedDeliverableIDsp = parAssociatedDeliverableIDsp;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting DeliverableServiceLevel ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}
		//---G
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>Deliverable object is retrieved if it exist, else null is retured.</returns>
		public static DeliverableServiceLevel Read(int parIDsp)
			{
			DeliverableServiceLevel result = new DeliverableServiceLevel();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<DeliverableServiceLevel>()
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
		//++GetDeliverablesForServiceLevels
		/// <summary>
		/// Read/retrieve all the deliverable entries that are associated with a specific ServiceLevel. 
		/// Specify a parameter containing the SharePoint ID value of all the ServiceLevel for which the deliverables are required.
		/// </summary>
		/// <param name="parServiceLevelIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple<DeliverableServiceLevel, Deliverrable>> is retrurned.</returns>
		public static List<Tuple<DeliverableServiceLevel, Deliverable>> GetDeliverablesForServiceLevel(int parServiceLevelIDsp)
			{
			List<Tuple<DeliverableServiceLevel, Deliverable>> results = new List<Tuple<DeliverableServiceLevel, Deliverable>>();
			if (parServiceLevelIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the DeliverableServiceLevel objects with which the specified Service Element (parElementIDsp) is associated 
					var technologyProdDeliverables = from eld in dbSession.AllObjects<DeliverableServiceLevel>()
											  where eld.AssociatedServiceLevelIDsp == parServiceLevelIDsp select eld;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in technologyProdDeliverables)
						{
						Deliverable deliverableEntry = (from entry in dbSession.AllObjects<Deliverable>()
														where entry.IDsp == item.AssociatedDeliverableIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (deliverableEntry != null)
							{
							Tuple<DeliverableServiceLevel, Deliverable> result = new Tuple<DeliverableServiceLevel, Deliverable>
								(item, deliverableEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all DeliverableServiceLevelProduct ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		//---g
		//++GetServiceLevelsForDeliverable
		/// <summary>
		/// Read/retrieve all the ServiceLevel entries that are associated with a specific Deliverable. 
		/// Specify a parameter containing the SharePoint ID value of all the Deliverable for which the Service Elements are required.
		/// </summary>
		/// <param name="parDeliverableIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple< DeliverableServiceLevel, ServiceLevel>> is retrurned.</returns>
		public static List<Tuple<DeliverableServiceLevel, ServiceLevel>> GetServiceLevelsForDeliverable(int parDeliverableIDsp)
			{
			List<Tuple<DeliverableServiceLevel, ServiceLevel>> results = new List<Tuple<DeliverableServiceLevel,ServiceLevel>>();

			if (parDeliverableIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the DeliverableServiceLevel objects with which the specified Deliverable (parDeliverableIDsp) is associated 
					var deliverableServiceLevels = from deliverableSL in dbSession.AllObjects<DeliverableServiceLevel>()
											  where deliverableSL.AssociatedDeliverableIDsp == parDeliverableIDsp
											  select deliverableSL;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in deliverableServiceLevels)
						{
						ServiceLevel serviceLevelEntry = (from entry in dbSession.AllObjects<ServiceLevel>()
														where entry.IDsp == item.AssociatedServiceLevelIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (serviceLevelEntry != null)
							{
							Tuple<DeliverableServiceLevel, ServiceLevel> result = new Tuple<DeliverableServiceLevel, ServiceLevel>
								(item, serviceLevelEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all DeliverableServiceLevel ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		#endregion
		}
	}
