using System;
using System.Data;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DogGenUI
	{
	class Comic
		{
		public string Name
			{
			get;
			set;
			}
		public int Issue
			{
			get;
			set;
			}
		public static IEnumerable<Comic> BuildCatalogue()
			{
			return new List<Comic>
				{
				new Comic {Name="Johnny America vs. the Pinco", Issue=6 },
				new Comic {Name="Rock and Roll (limited edition)", Issue=19 },
				new Comic {Name="Woman's Work", Issue=36 },
				new Comic {Name="Hippie Madness(misprinted)", Issue=57 },
				new Comic {Name="Revenge of the New Wave Freak (damaged)", Issue=68},
				new Comic {Name="Black Monday", Issue=74},
				new Comic {Name="Tribal Tattoo Madnes", Issue=83 },
				new Comic {Name="The Death of an Object", Issue=97 }
				};
			}
		public static Dictionary<int, decimal> GetPrices()
			{
			return new Dictionary<int, decimal>
				{
					{6, 3600M },
					{19, 500M },
					{36, 650M },
					{57, 1325M },
					{68, 250M },
					{74, 75M },
					{83, 25.75M },
					{97, 35.25M }
				};
			}
		}
	}
