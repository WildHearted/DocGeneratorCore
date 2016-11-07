using System;
using System.Collections.Generic;
using System.Linq;
using VelocityDb;
using VelocityDb.Indexing;
using VelocityDb.Session;

namespace DocGeneratorCore.Database.Classes
	{
	public class TechnologyProduct : OptimizedPersistable
		{
		/// <summary>
		/// This class is used to store a single object that contains a TechnologyProduct as mapped to the SharePoint List named TechnologyProducts.
		/// </summary>
		#region Variables
		[Index]
		[UniqueConstraint]
		private int _IDsp;
		private string _Title;
		[Index]
		private TechnologyCategory _Category;
		[Index]
		private TechnologyVendor _Vendor;
		private string _Prerequisites;
		#endregion

		#region Properties
		public int IDsp {
			get { return this._IDsp; }
			set { Update(); this._IDsp = value;}
			}
		public string Title {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}
		public TechnologyCategory Category {
			get { Session?.LoadFields(pObj: this._Category); return this._Category; }
			set { Update(); this._Category = value; }
			}
		public TechnologyVendor Vendor {
			get { Session?.LoadFields(pObj: this._Vendor); return this._Vendor; }
			set { Update(); this._Vendor= value;
				}
			}
		public string Prerequisites {
			get { return this._Title; }
			set { UpdateNonIndexField(); this._Title = value; }
			}
		#endregion

		//===G
		#region Methods
		//++Store
		/// <summary>
		/// Store/Save a new Object in the database, use the same Store method for New and Updates.
		/// </summary>
		public static bool Store(
			int parIDsp,
			string parTitle,
			string parPrerequisites, 
			TechnologyCategory parCategory,
			TechnologyVendor parVendor)

			{
			TechnologyProduct newEntry;
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginUpdate();
					newEntry = (from objEntry in dbSession.AllObjects<TechnologyProduct>()
								where objEntry.IDsp == parIDsp
								select objEntry).FirstOrDefault();
					if (newEntry == null)
						newEntry = new TechnologyProduct();
					newEntry.IDsp = parIDsp;
					newEntry.Title = parTitle;
					newEntry.Category = parCategory;
					newEntry.Vendor = parVendor;
					newEntry.Prerequisites = parPrerequisites;
					dbSession.Persist(newEntry);
					dbSession.Commit();
					return true;
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database persisting TechnologyProduct ### - {0} - {1}", exc.HResult, exc.Message);
				return false;
				}
			}

		//++Read
		/// <summary>
		/// Read/retrieve all the entries from the database
		/// </summary>
		/// <returns>DataStatus object is retrieved if it exist, else null is retured.</returns>
		public static TechnologyProduct Read(int parIDsp)
			{
			TechnologyProduct result = new TechnologyProduct();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();

					result = (from thisEntry in dbSession.AllObjects<TechnologyProduct>()
							  where thisEntry.IDsp == parIDsp
							  select thisEntry).FirstOrDefault();
					}
				}
			catch (Exception exc)
				{
				result = null;
				Console.WriteLine("### Exception Database reading TechnologyProduct [{0}] ### - {1} - {2}", parIDsp, exc.HResult, exc.Message);
				}
			return result;
			}

		//++ReadAll
		/// <summary>
		/// Read/retrieve all the entries from the database. 
		/// Specify a List of intergers containing the SharePoint ID values of all the TechnologyProduct objects
		/// that need to be retrived and added to the list.
		/// </summary>
		/// <param name="parIDs">pass a List<int> of all the IDsp (SharePoint ID) that need to be retrieved and returned.
		/// If all TechnologyProducts must be retrieve, pass an empty List (with count = 0) to return all objects.</int> </param>
		/// <returns>a List<TechnologyProduct> objects are retrurned.</returns>
		public static List<TechnologyProduct> ReadAll(List<int> parIDs)
			{
			List<TechnologyProduct> results = new List<TechnologyProduct>();
			try
				{
				using (ServerClientSession dbSession = new ServerClientSession(systemDir: Properties.Settings.Default.CurrentDatabaseLocation))
					{
					dbSession.BeginRead();
					//-|Return all Products if no product is specified
					if (parIDs.Count == 0)
						{
						foreach (TechnologyProduct entry in dbSession.AllObjects<TechnologyProduct>())
							{
							results.Add(entry);
							}
						}
					else //-| Specific entries were specified.
						{
						foreach (int item in parIDs)
							{
							TechnologyProduct entry = new TechnologyProduct();
							entry = (from thisEntry in dbSession.AllObjects<TechnologyProduct>()
									 where thisEntry.IDsp == item
									 select thisEntry).FirstOrDefault();
							results.Add(entry);
							}
						}
					}
				}
			catch (Exception exc)
				{
				Console.WriteLine("### Exception Database reading all TechnologyProduct ### - {0} - {1}", exc.HResult, exc.Message);
				}
			return results;
			}
		#endregion
		}
	}
