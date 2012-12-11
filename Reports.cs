#region Imported Namespaces
using System;
using System.Threading;
#endregion


namespace TeamSiteReports
{

	class Reports
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]			
		static void Main(string[] args)
		{

			ReportTeamSiteOwners rOwners = new ReportTeamSiteOwners();
			ReportTeamSiteSize rSize = new ReportTeamSiteSize();
			ReportTeamSiteAll rMembers = new ReportTeamSiteAll();
			ReportTeamSiteSubSite rSubsite = new ReportTeamSiteSubSite();
			ReportTeamSiteLastModified rModified = new ReportTeamSiteLastModified();

			Thread thread1 = new Thread(new ThreadStart(rOwners.GenerateReport));
			Thread thread2 = new Thread(new ThreadStart(rSize.GenerateReport));
			Thread thread3 = new Thread(new ThreadStart(rMembers.GenerateReport));	
			Thread thread4 = new Thread(new ThreadStart(rModified.GenerateReport));
			Thread thread5 = new Thread(new ThreadStart(rSubsite.GenerateReport));
			
			// start them
			thread1.Start();
			thread2.Start();
			thread3.Start();
			thread4.Start();
			thread5.Start();


		}

	}
}
