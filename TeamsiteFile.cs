#region Imported Namespaces
using System;
using System.IO; 
using Microsoft.SharePoint;
using EA.Logger;
#endregion
namespace TeamSiteReports
{
	/// <summary>
	/// Contains some methods used for loading files into teamsite
	/// </summary>
	public class TeamSiteFile
	{

		public void UploadFile(string srcUrl, string destUrl)
		{
			if (! File.Exists(srcUrl))
			{
				FileErrorLogger _logger = new FileErrorLogger();
				_logger.LogError(String.Format("{0} does not exist", srcUrl), ErrorLogSeverity.SeverityError,
					ErrorType.TypeApplication, "TeamSiteReports.UploadFile()");
				_logger = null;
			}
			

			SPWeb site = new SPSite(destUrl).OpenWeb();
			try 
			{
				FileStream fStream = File.OpenRead(srcUrl);
				byte[] contents = new byte[fStream.Length];
				fStream.Read(contents, 0, (int)fStream.Length);
				fStream.Close(); 
				site.Files.Add(destUrl, contents,true);
			}
			catch(SPException spe)
			{
				FileErrorLogger _logger = new FileErrorLogger();
				_logger.LogError(spe.Message+srcUrl, ErrorLogSeverity.SeverityError,
					ErrorType.TypeApplication, "TeamSiteReports.UploadFile()");
				_logger = null;
				

			}
			catch(Exception e)
			{
				Object thisLock = new Object();
				lock (thisLock)
				{
					FileErrorLogger _logger = new FileErrorLogger();
					_logger.LogError(e.Message+srcUrl, ErrorLogSeverity.SeverityError,
						ErrorType.TypeApplication, "TeamSiteReports.UploadFile()");
					_logger = null;
				}
			}
			finally
			{
				if (site !=null)
					site.Dispose();

			}


		}
		private string EnsureParentFolder(SPWeb parentSite, string destinUrl)
		{
			destinUrl = parentSite.GetFile(destinUrl).Url;

			int index = destinUrl.LastIndexOf("/");
			string parentFolderUrl = string.Empty;

			if (index > -1)
			{
				parentFolderUrl = destinUrl.Substring(0, index);

				SPFolder parentFolder = parentSite.GetFolder(parentFolderUrl);

				if (! parentFolder.Exists)
				{
					SPFolder currentFolder = parentSite.RootFolder;

					foreach(string folder in parentFolderUrl.Split('/'))
					{
						currentFolder = currentFolder.SubFolders.Add(folder);
					}
				}
			}

			return parentFolderUrl;
		}



	}
}
