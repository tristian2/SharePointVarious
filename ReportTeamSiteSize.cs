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
	/// Class to produce report on sum of file size for all teamsite files
	/// </summary>
	public class ReportTeamSiteSize : Report
	{
		long totalFileSize;
		int fileCount;

		public override void GenerateReport()
		{
			TeamSiteFile tf = new TeamSiteFile();
			string fileName = System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSizeReportLocation"];
					
			StreamWriter fwriter = File.CreateText( fileName ); 
			Console.WriteLine("Creating file "+System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSizeReportLocation"]);

			HtmlTextWriter txtWriter=new HtmlTextWriter(fwriter);

			HtmlTable reportHtmlTable = new HtmlTable();
			reportHtmlTable.BgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.White);
			reportHtmlTable.Border = 1;
			reportHtmlTable.BorderColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.LightGray );
			reportHtmlTable.Style.Add("font-family", "Verdana");
			reportHtmlTable.Style.Add("font-size", "9pt");

			HtmlTableRow trMessage = new HtmlTableRow();
			HtmlTableCell tcMessage = new HtmlTableCell();
			tcMessage.ColSpan = 10;
			tcMessage.InnerText = @"Last run: " + System.DateTime.Now.ToString();
			tcMessage.Style.Add("font-style","italic");
			trMessage.Cells.Add(tcMessage);
			reportHtmlTable.Rows.Add(trMessage);

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

			//teamsite # docs
			HtmlTableCell tcHeader3 = new HtmlTableCell();
			tcHeader3.InnerText = "#docs";
			trHeader.Cells.Add(tcHeader3);

			//teamsite size Mbytes
			HtmlTableCell tcHeader4 = new HtmlTableCell();
			tcHeader4.InnerText = "filesize (MB)";
			trHeader.Cells.Add(tcHeader4);
			reportHtmlTable.Rows.Add(trHeader);

			Console.WriteLine("Connecting to site...");
			SPSite siteCollection = new SPSite(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteUrl"]);
			SPWebCollection sites = siteCollection.AllWebs;
			
			try
			{	
				foreach (SPWeb site in sites)
				{				
					try
					{

						totalFileSize = 0;
						fileCount = 0;

						SPFolderCollection folders = site.Folders;
						traverseFolders(folders);

						Console.WriteLine( "Summary: " + SPEncode.HtmlEncode(site.Name) + " Number: " + fileCount + 
							" Size: " + totalFileSize);

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

						//teamsite # docs 
						HtmlTableCell tcData3 = new HtmlTableCell();
						tcData3 .InnerText = fileCount.ToString();
						trData.Cells.Add(tcData3);

						//teamsite size Mbytes
						HtmlTableCell tcData4 = new HtmlTableCell();
						totalFileSize = totalFileSize / 1000000;
						tcData4.InnerText = totalFileSize.ToString();
						//tcData4.BgColor = roleStyle;
						trData.Cells.Add(tcData4);

						reportHtmlTable.Rows.Add(trData);

					}
					catch(Exception ex)
					{
						FileErrorLogger _logger = new FileErrorLogger();
						_logger.LogError(ex.Message, ErrorLogSeverity.SeverityError,
							ErrorType.TypeApplication, "Intranet.TeamSiteReports.UploadFile()");
						_logger = null;
					}
					finally
					{						
						site.Dispose();
					}				
				}

			}
			catch (Exception ex)
			{
				FileErrorLogger _logger = new FileErrorLogger();
				_logger.LogError(ex.Message, ErrorLogSeverity.SeverityError,
					ErrorType.TypeApplication, "Intranet.TeamSiteReports.UploadFile()");
				_logger = null;
			}
			finally
			{
				siteCollection.Dispose();
			}
			
			reportHtmlTable.RenderControl(txtWriter);

			txtWriter.Close(); 
			tf.UploadFile(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSizeReportLocation"], 
				System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSizeReportDestination"]);

		}
		private void traverseFolders(SPFolderCollection folders)
		{

			foreach (SPFolder folder in folders)
			{
				Console.WriteLine("Folder: "+folder.Name);
				SPFileCollection files = folder.Files;

				for (int i=0; i<files.Count; i++)        
				{
					totalFileSize += files[i].Length;
					Console.WriteLine("\tFile: "+files[i].Name.ToString());
					fileCount ++;
				}

				traverseFolders(folder.SubFolders);
				
			}
			
		}
	}

}
