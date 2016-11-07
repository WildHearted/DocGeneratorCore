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
	class DatabaseSchema
		{

		public UInt64 Persist(
			Placement parPlace, 
			SessionBase parSession, 
			bool persistRefs = true,
			bool disableFlush = false, 
			Queue<IOptimizedPersistable> toPersist = null)
			{
			parSession.RegisterClass(type: typeof(GlossaryAcronym));
			parSession.RegisterClass(type: typeof(JobRole));
			parSession.RegisterClass(type: typeof(ActivityCategory));
			parSession.RegisterClass(type: typeof(ServiceLevelCategory));
			parSession.RegisterClass(type: typeof(TechnologyCategory));
			parSession.RegisterClass(type: typeof(TechnologyVendor));
			parSession.RegisterClass(type: typeof(TechnologyProduct));

			parSession.RegisterClass(type: typeof(ServicePortfolio));
			parSession.RegisterClass(type: typeof(ServiceFamily));
			parSession.RegisterClass(type: typeof(ServiceProduct));
			parSession.RegisterClass(type: typeof(ServiceElement));
			parSession.RegisterClass(type: typeof(ServiceFeature));
			parSession.RegisterClass(type: typeof(Deliverable));
			parSession.RegisterClass(type: typeof(ElementDeliverable));
			parSession.RegisterClass(type: typeof(FeatureDeliverable));
			parSession.RegisterClass(type: typeof(DeliverableTechnology));
			parSession.RegisterClass(type: typeof(ServiceLevel));
			parSession.RegisterClass(type: typeof(ServiceLevelTarget));
			parSession.RegisterClass(type: typeof(DeliverableServiceLevel));

			parSession.RegisterClass(type: typeof(Activity));
			parSession.RegisterClass(type: typeof(DeliverableActivity));

			parSession.RegisterClass(typeof(AutoPlacement));



			return 0;
			}

		}
	}
