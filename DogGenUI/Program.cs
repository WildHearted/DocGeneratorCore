using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocGeneratorCore
	{
	static class Program
		{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
			{
			MainController objMainControl = new MainController();
			//List<DocumentCollection> objListdocumentCollections = new List<DocumentCollection>();
			//enumPlatform objPlatform = enumPlatform.Production;
			objMainControl.MainProcess();
			}
		}
	}
