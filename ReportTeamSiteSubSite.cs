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
	/// Class to produce report containing rows for all teamsite subsites 
	/// </summary>
	public class ReportTeamSiteSubSite : Report
	{
		public override void GenerateReport()
		{
			TeamSiteFile tf = new TeamSiteFile();
			string fileName = System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSubsiteReportLocation"];
			StreamWriter fwriter = File.CreateText( fileName ); 
			Console.WriteLine("Creating file "+System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSubsiteReportLocation"]);
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

			//teamsite brand
			HtmlTableCell tcHeader3 = new HtmlTableCell();
			tcHeader3.InnerText = "brand";
			trHeader.Cells.Add(tcHeader3);

			//teamsite subsite name
			HtmlTableCell tcHeader4 = new HtmlTableCell();
			tcHeader4.InnerText = "subsite name";
			trHeader.Cells.Add(tcHeader4);

			//teamsite subsite url
			HtmlTableCell tcHeader5 = new HtmlTableCell();
			tcHeader5.InnerText = "subsite url";
			trHeader.Cells.Add(tcHeader5);

			reportHtmlTable.Rows.Add(trHeader);

			SPSite siteCollection = new SPSite(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteUrl"]);
			SPWebCollection sites = siteCollection.AllWebs;

			try
			{
				foreach (SPWeb site in sites)
				{
					try
					{
						//ignore top level brand sites
						switch (site.Name)
						{
							case "": 
								break;
							case "doh":
								break;
							case "Re":
								break;
							case "Mi":
								break;
							default:
								SPWebCollection subsites = site.Webs;
								//if there are subsites log them
								if(subsites.Count > 0)
								{
									foreach (SPWeb subsite in subsites)
									{
										try
										{
											Console.WriteLine("Site:"+site.Name);
											Console.WriteLine("\tSubsite:"+subsite.Name);
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

											//subsite name
											HtmlTableCell tcData4 = new HtmlTableCell();
											tcData4.InnerText = subsite.Name;
											trData.Cells.Add(tcData4);

											//subsite url
											HtmlTableCell tcData5 = new HtmlTableCell();
											HtmlAnchor ha2 = new HtmlAnchor();
											ha2.InnerText=subsite.Url;		
											ha2.HRef=subsite.Url;
											tcData5.Controls.Add(ha2);	
											trData.Cells.Add(tcData5);

											reportHtmlTable.Rows.Add(trData);
									
										}
										catch(Exception ex)
										{
											FileErrorLogger _logger = new FileErrorLogger();
											_logger.LogError(ex.Message, ErrorLogSeverity.SeverityError,
												ErrorType.TypeApplication, "TeamSiteReports.UploadFile()");
											_logger = null;
										}
										finally
										{
											subsite.Dispose();
										}
									}
								
								}
								break;
						}//switch
					}//try
					catch(Exception ex)
					{
						FileErrorLogger _logger = new FileErrorLogger();
						_logger.LogError(ex.Message, ErrorLogSeverity.SeverityError,
							ErrorType.TypeApplication, "TeamSiteReports.UploadFile()");
						_logger = null;					
					}
					finally
					{
						site.Dispose();
					}
				}//foreach
			}//try
			catch(Exception ex)
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
			tf.UploadFile(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSubsiteReportLocation"], 
				System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteSubsiteReportDestination"]);

		}

	}
}
