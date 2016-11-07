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
	public class FeatureDeliverable : OptimizedPersistable
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
		private int? _AssociatedFeatureIDsp;
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
		public string Optionality {
			get { return this._Optionality; }
			set { UpdateNonIndexField(); this._Optionality = value; }
			}
		public string AdditionalConditions {
			get { return this._AdditionalConditions; }
			set { UpdateNonIndexField(); this._AdditionalConditions = value;
				}
			}
		public int? AssociatedDeliverableIDsp {
			get { return this._AssociatedDeliverableIDsp; }
			set { Update(); this._AssociatedDeliverableIDsp = value; }
			}
		public int? AssociatedFeatureIDsp {
			get { return this._AssociatedFeatureIDsp; }
			set { Update(); this._AssociatedFeatureIDsp = value; }
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
			string parOptionality,
			string parAdditionalConditions,
			int parAssociatedDeliverableIDsp,
			int parAssociatedFeaturesp
			)
			{
			FeatureDeliverable newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<FeatureDeliverable>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new FeatureDeliverable();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.Optionality  = parOptionality;
					newEntry.AdditionalConditions = parAdditionalConditions;
					newEntry.AssociatedFeatureIDsp = parAssociatedFeaturesp;
					newEntry.AssociatedDeliverableIDsp = parAssociatedDeliverableIDsp;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting FeatureDeliverable ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}
		//---G
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>FeatureDeliverable object is retrieved if it exist, else null is retured.</returns>
		public static FeatureDeliverable Read(int parIDsp)
			{
			FeatureDeliverable result = new FeatureDeliverable();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<FeatureDeliverable>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading FeatureDeliverable [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//---g
		//++GetDeliverablesForFeature
		/// <summary>
		/// Read/retrieve all the deliverable entries that are associated with a specific Service Element. 
		/// Specify a parameter containing the SharePoint ID value of all the Service Element for which the deliverables are required.
		/// </summary>
		/// <param name="parElementIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple<ElementDeliverable, Deliverrable>> is retrurned.</returns>
		public static List<Tuple<ElementDeliverable, Deliverable>> GetDeliverablesForElement(int parElementIDsp)
			{
			List<Tuple<ElementDeliverable, Deliverable>> results = new List<Tuple<ElementDeliverable, Deliverable>>();
			if (parElementIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the ElementDeliverable objects with which the specified Service Element (parElementIDsp) is associated 
					var elementDeliverables = from eld in dbSession.AllObjects<ElementDeliverable>()
											  where eld.AssociatedElementIDsp == parElementIDsp select eld;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in elementDeliverables)
						{
						Deliverable deliverableEntry = (from entry in dbSession.AllObjects<Deliverable>()
														where entry.IDsp == item.AssociatedDeliverableIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (deliverableEntry != null)
							{
							Tuple<ElementDeliverable, Deliverable> result = new Tuple<ElementDeliverable, Deliverable>
								(item, deliverableEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all Deliverable ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		//---g
		//++GetElementsForDeliverable
		/// <summary>
		/// Read/retrieve all the Service Element entries that are associated with a specific Deliverable. 
		/// Specify a parameter containing the SharePoint ID value of all the Deliverable for which the Service Elements are required.
		/// </summary>
		/// <param name="parDeliverableIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple<ElementDeliverable, Deliverrable>> is retrurned.</returns>
		public static List<Tuple<ElementDeliverable, ServiceElement>> GetElementsForDeliverable(int parDeliverableIDsp)
			{
			List<Tuple<ElementDeliverable, ServiceElement>> results = new List<Tuple<ElementDeliverable, ServiceElement>>();

			if (parDeliverableIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the ElementDeliverable objects with which the specified Service Element (parElementIDsp) is associated 
					var elementDeliverables = from eld in dbSession.AllObjects<ElementDeliverable>()
											  where eld.AssociatedDeliverableIDsp == parDeliverableIDsp
											  select eld;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in elementDeliverables)
						{
						ServiceElement elementEntry = (from entry in dbSession.AllObjects<ServiceElement>()
														where entry.IDsp == item.AssociatedDeliverableIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (elementEntry != null)
							{
							Tuple<ElementDeliverable, ServiceElement> result = new Tuple<ElementDeliverable, ServiceElement>
								(item, elementEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all ElementDeliverable ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		#endregion
		}
	}
