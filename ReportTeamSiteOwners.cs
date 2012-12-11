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
	/// Class to produce the teamsite owners report containing rows for all teamsite owners
	/// </summary>
	public class ReportTeamSiteOwners : Report
	{
		public override void GenerateReport()
		{
			TeamSiteFile tf = new TeamSiteFile();
			string fileName = System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteOwnersReportLocation"];
			string roleStyle;
					
			StreamWriter fwriter = File.CreateText( fileName ); 
			Console.WriteLine("Creating file "+System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteOwnersReportLocation"]);

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

			//teamsite users fullname
			HtmlTableCell tcHeader4 = new HtmlTableCell();
			tcHeader4.InnerText = "teamsite owners";
			trHeader.Cells.Add(tcHeader4);

			//lanId
			HtmlTableCell tcHeader5 = new HtmlTableCell();
			tcHeader5.InnerText = "lanId" ;
			trHeader.Cells.Add(tcHeader5);

			//email address
			HtmlTableCell tcHeader6 = new HtmlTableCell();
			tcHeader6.InnerText = "email" ;
			trHeader.Cells.Add(tcHeader6);

			//teamsite request for access email address
			HtmlTableCell tcHeader8 = new HtmlTableCell();
			tcHeader8.InnerText = "request for access email";
			trHeader.Cells.Add(tcHeader8);

			//teamsite memebership count
			HtmlTableCell tcHeader9 = new HtmlTableCell();
			tcHeader9.InnerText = "TeamSite Membership";
			trHeader.Cells.Add(tcHeader9);	

			//subsites
			HtmlTableCell tcHeader10 = new HtmlTableCell();
			tcHeader10.InnerText = "Subsites";
			trHeader.Cells.Add(tcHeader10);

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
						SPWebCollection subSites = site.Webs;
						int subsitesCount = subSites.Count;
						int roleCount = 0;
						SPUserCollection users = site.Users;
						
						Console.WriteLine("Site: "+site.Name);
						foreach(SPUser user in users)
						{
							SPRoleCollection roles = user.Roles;

							//we need to count the roles here we count the number of teamsite owners..
							foreach(SPRole role in roles)
							{
								if (role.Name == System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.roleExclusiveInclude"].ToString())
								{
									roleCount ++;
								}
							}


						} //user
						if (roleCount > Convert.ToInt16(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteOwnersCount"]) )
						{
							//set the style to flag this on the cell	
							roleStyle = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.Orange);
						} 
						else 
						{
							roleStyle = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.White);
						}

						foreach(SPUser user in users)
						{
							SPRoleCollection roles = user.Roles;
							Console.WriteLine("\tUser: "+user.Name);
							//now build the main body of the report
							foreach(SPRole role in roles)
							{
								
								if (role.Name == System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.roleExclusiveInclude"].ToString())
								{							
									string sRole = role.Name.ToString();

									SPListCollection lists =  site.Lists;
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
										tcData3.InnerText = ary[3].ToString(); // e.g. http://ts/one fourth index will contain the brand
									}
									catch  //the url may not contain the brand for instance the top level site
									{
										tcData3 .InnerText = "na";
									}
									trData.Cells.Add(tcData3);

									//teamsite users fullname
									HtmlTableCell tcData4 = new HtmlTableCell();
									tcData4.InnerText = user.Name;
									tcData4.BgColor = roleStyle;
									trData.Cells.Add(tcData4);

									//teamsite user lanId
									HtmlTableCell tcData5 = new HtmlTableCell();
									tcData5.InnerText = user.LoginName;
									tcData5.BgColor = roleStyle;
									trData.Cells.Add(tcData5);

									//teamsite user email address
									HtmlTableCell tcData6 = new HtmlTableCell();
									HtmlAnchor haEmail = new HtmlAnchor();
									haEmail.InnerText="mailto:"+user.Email;
									haEmail.HRef=user.Email;
									tcData6.Controls.Add(haEmail);
									tcData6.InnerText = user.Email ;  //email
									trData.Cells.Add(tcData6);

									//teamsite request for access email address	
									HtmlTableCell tcData8 = new HtmlTableCell();
									try
									{

										SPPermissionCollection permsSite = site.Permissions;

										if (permsSite.RequestAccess)
										{
											tcData8.InnerText = permsSite.RequestAccessEmail.ToString();
										} 
										else 
										{
											tcData8.InnerText = "";
										}
									}
									catch //
									{
										tcData8.BgColor = "#FF0000";
										tcData8.InnerText = "permissions error";
									}

									trData.Cells.Add(tcData8);

									//teamsite memebrship count
									HtmlTableCell tcData9 = new HtmlTableCell();
									tcData9.InnerText = site.Users.Count.ToString();
									trData.Cells.Add(tcData9);	

									//subsites
									
									HtmlTableCell tcData10 = new HtmlTableCell();
									tcData10.InnerText = subsitesCount.ToString();
									if (subsitesCount>0)
										tcData10.BgColor = System.Drawing.ColorTranslator.ToHtml (System.Drawing.Color.Red);
									trData.Cells.Add(tcData10);

									reportHtmlTable.Rows.Add(trData);

								}
							}
							site.Dispose();
						}
					}
					catch(Exception ex)
					{
						FileErrorLogger _logger = new FileErrorLogger();
						_logger.LogError(ex.Message, ErrorLogSeverity.SeverityError,
							ErrorType.TypeApplication, "Centrica.Intranet.TeamSiteReports.UploadFile()");
						_logger = null;					}
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
					ErrorType.TypeApplication, "Centrica.Intranet.TeamSiteReports.UploadFile()");
				_logger = null;
			}
			finally 
			{
				siteCollection.Dispose();
			}
			
			reportHtmlTable.RenderControl(txtWriter);

			txtWriter.Close(); 
			tf.UploadFile(System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteOwnersReportLocation"], 
				System.Configuration.ConfigurationSettings.AppSettings["TeamSiteReports.TeamSiteOwnersReportDestination"]);
			Console.WriteLine("Completed");

		}

	}
}
