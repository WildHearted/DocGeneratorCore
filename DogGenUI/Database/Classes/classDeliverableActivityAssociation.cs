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
	public class DeliverableActivity : OptimizedPersistable
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
		private int? _AssociatedActivityIDsp;
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
		public string Optionality {
			get { return this._Optionality; }
			set { UpdateNonIndexField(); this._Optionality = value; }
			}
		public int? AssociatedDeliverableIDsp {
			get { return this._AssociatedDeliverableIDsp; }
			set { Update(); this._AssociatedDeliverableIDsp = value; }
			}
		public int? AssociatedActivityIDsp {
			get { return this._AssociatedActivityIDsp; }
			set { Update(); this._AssociatedActivityIDsp = value; }
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
			int parAssociatedDeliverableIDsp,
			int parAssociatedActivityIDsp)
		
			{
			DeliverableActivity newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<DeliverableActivity>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new DeliverableActivity();

					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.Optionality = parOptionality;
					newEntry.AssociatedActivityIDsp = parAssociatedActivityIDsp;
					newEntry.AssociatedDeliverableIDsp = parAssociatedDeliverableIDsp;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting DeliverableActivity ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}
		//---G
		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>Deliverable object is retrieved if it exist, else null is retured.</returns>
		public static DeliverableActivity Read(int parIDsp)
			{
			DeliverableActivity result = new DeliverableActivity();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<DeliverableActivity>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading DeliverableActivity [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//---g
		//++GetDeliverablesForActivity
		/// <summary>
		/// Read/retrieve all the deliverable entries that are associated with a specific Activity. 
		/// Specify a parameter containing the SharePoint ID value of all the Activity for which the deliverables are required.
		/// </summary>
		/// <param name="parActivityIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple<DeliverableActivity, Deliverrable>> is retrurned.</returns>
		public static List<Tuple<DeliverableActivity, Deliverable>> GetDeliverablesForActivity(int parActivityIDsp)
			{
			List<Tuple<DeliverableActivity, Deliverable>> results = new List<Tuple<DeliverableActivity, Deliverable>>();
			if (parActivityIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the DeliverableActivity objects with which the specified Service Element (parElementIDsp) is associated 
					var activityDeliverables = from actDeliverable in dbSession.AllObjects<DeliverableActivity>()
											  where actDeliverable.AssociatedDeliverableIDsp == parActivityIDsp select actDeliverable;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in activityDeliverables)
						{
						Deliverable deliverableEntry = (from entry in dbSession.AllObjects<Deliverable>()
														where entry.IDsp == item.AssociatedDeliverableIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (deliverableEntry != null)
							{
							Tuple<DeliverableActivity, Deliverable> result = new Tuple<DeliverableActivity, Deliverable>
								(item, deliverableEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all DeliverableActivity ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		//---g
		//++GetActivitiesForDeliverable
		/// <summary>
		/// Read/retrieve all the Activity entries that are associated with a specific Deliverable. 
		/// Specify a parameter containing the SharePoint ID value of all the Deliverable for which the Service Elements are required.
		/// </summary>
		/// <param name="parDeliverableIDsp">pass an integer of the IDsp (SharePoint ID) of the Service Element for which the Deliverables need to be retrieved.</param>
		/// <returns>a List<Tuple< DeliverableActivity, Activity>> is retrurned.</returns>
		public static List<Tuple<DeliverableActivity, Activity>> GetActivitiesForDeliverable(int parDeliverableIDsp)
			{
			List<Tuple<DeliverableActivity, Activity>> results = new List<Tuple<DeliverableActivity,Activity>>();

			if (parDeliverableIDsp == 0)
				return results;

			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Obtain the DeliverableActivity objects with which the specified Deliverable (parDeliverableIDsp) is associated 
					var deliverableActivities = from deliverableAct in dbSession.AllObjects<DeliverableActivity>()
											  where deliverableAct.AssociatedDeliverableIDsp == parDeliverableIDsp
											  select deliverableAct;
					//-|Process each entry and retrived all the Deliverables... 
					foreach (var item in deliverableActivities)
						{
						Activity serviceLevelEntry = (from entry in dbSession.AllObjects<Activity>()
														where entry.IDsp == item.AssociatedActivityIDsp
														select entry).FirstOrDefault();
						//-|If the Deliverable is retrieved, add it to the results...
						if (serviceLevelEntry != null)
							{
							Tuple<DeliverableActivity, Activity> result = new Tuple<DeliverableActivity, Activity>
								(item, serviceLevelEntry);
							results.Add(result);
							}
						}
					dbSession.Commit();
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all DeliverableActivity ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}

		#endregion
		}
	}
