#region Imported Namespaces
using System;
using System.Drawing;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using Logger;
#endregion

namespace TeamSiteReports
{

	/// <summary>
	/// Class to produce the "master" teamsite report on the date the last list item modification occurred
	/// </summary>
	public class ReportTeamSiteLastModified : Report
	{
		
		public override void GenerateReport()
		{
			TeamSiteFile tf = new TeamSiteFile();			
			string fileName = System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteLastModifiedReportLocation"];
			string lastModified;

			StreamWriter fwriter = File.CreateText( fileName ); 
			Console.WriteLine("Creating file "+System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteLastModifiedReportLocation"]);

			HtmlTextWriter txtWriter=new HtmlTextWriter(fwriter);

			HtmlTable reportHtmlTable = new HtmlTable();	
			reportHtmlTable.BgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.White);
			reportHtmlTable.Border = 1;
			reportHtmlTable.BorderColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.LightGray );
			reportHtmlTable.Style.Add("font-family", "Verdana");
			reportHtmlTable.Style.Add("font-size", "9pt");

			HtmlTableRow trMessage = new HtmlTableRow();
			HtmlTableCell tcMessage = new HtmlTableCell();
			tcMessage.ColSpan = 4;
			tcMessage.InnerText = @"Last run: " + System.DateTime.Now.ToString();
			tcMessage.Style.Add("font-style","italic");
			trMessage.Cells.Add(tcMessage);
			reportHtmlTable.Rows.Add(trMessage);


			reportHtmlTable.BgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.White);
			reportHtmlTable.Border = 1;
			reportHtmlTable.BorderColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.LightGray );
			reportHtmlTable.Style.Add("font-family", "Verdana");
			reportHtmlTable.Style.Add("font-size", "9pt");

			HtmlTableRow trHeader = new HtmlTableRow();
			trHeader.Style.Add("font-weight","bold");

			//teamsite name
			HtmlTableCell tcHeader1 = new HtmlTableCell();
			tcHeader1.InnerText = "teamsite name";
			trHeader.Cells.Add(tcHeader1);
			
			//teamsite url
			HtmlTableCell tcHeader2 = new HtmlTableCell();
			tcHeader2.InnerText = "teamsite url";
			trHeader.Cells.Add(tcHeader2);

			//teamsite brand
			HtmlTableCell tcHeader3 = new HtmlTableCell();
			tcHeader3.InnerText = "brand";
			trHeader.Cells.Add(tcHeader3);

			//teamsite lastModified
			HtmlTableCell tcHeader4 = new HtmlTableCell();
			tcHeader4.InnerText = "last modified";
			trHeader.Cells.Add(tcHeader4);
			reportHtmlTable.Rows.Add(trHeader);

			Console.WriteLine("Connecting to site...");
			SPSite siteCollection = new SPSite(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteUrl"]);
			SPWebCollection sites = siteCollection.AllWebs;
			try
			{
				foreach (SPWeb site in sites)
				{
					lastModified = "01/01/1900 01:01:01";

					//go through the lists in the site for the later lastmodified date
					foreach (SPList list in site.Lists)
					{
						if (System.DateTime.Parse(lastModified) < list.LastItemModifiedDate)
							lastModified = list.LastItemModifiedDate.ToString();
					}

					Console.WriteLine("Site:"+site.Name+" last modified:"+lastModified);
					HtmlTableRow trData = new HtmlTableRow();

					//teamsite name
					HtmlTableCell tcData1 = new HtmlTableCell();
					tcData1.InnerText = site.Name;
					trData.Cells.Add(tcData1);

					//teamsite url
					HtmlTableCell tcData2 = new HtmlTableCell();
					HtmlAnchor ha1 = new HtmlAnchor();
					ha1.InnerText=site.Url;			
					ha1.HRef=site.Url;
					tcData2.Controls.Add(ha1);	
					trData.Cells.Add(tcData2);

					//teamsite brand
					HtmlTableCell tcData3 = new HtmlTableCell();
					string brand = site.Url.ToString();
					try
					{
						string[] ary = brand.Split('/');
						tcData3.InnerText = ary[3].ToString(); // e.g. http:///blahblah fourth index will contain the brand
					}
					catch  //the url may not contain the brand for instance the top level site
					{
						tcData3 .InnerText = "na";
					}
					trData.Cells.Add(tcData3);

					//teamsite last modified date
					HtmlTableCell tcData4 = new HtmlTableCell();
					tcData4.InnerText = lastModified;	
					trData.Cells.Add(tcData4);

					reportHtmlTable.Rows.Add(trData);
				
					site.Dispose();
				}
			}
			catch (Exception ex)
			{
				FileErrorLogger _logger = new FileErrorLogger();
				_logger.LogError(ex.Message, ErrorLogSeverity.SeverityError,
					ErrorType.TypeApplication, "TeamSiteReports.UploadFile()");
				_logger = null;
			}
			finally
			{
				siteCollection.Dispose();
			}
			
			reportHtmlTable.RenderControl(txtWriter);

			txtWriter.Close(); 
			tf.UploadFile(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteLastModifiedReportLocation"], 
				System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteLastModifiedReportDestination"]);

		

		}
		
	}

}
