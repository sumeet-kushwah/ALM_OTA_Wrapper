using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALM_Wrapper
{
    public class Release
    {
        private TDAPIOLELib.TDConnection tDConnection;
        private ReleaseFolders releaseFolders;

        public Release(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
            releaseFolders = new ReleaseFolders(tDConnection);
        }

        /// <summary>
        /// Creates a release in ALM, you must pass values for all required fields
        /// </summary>
        /// <param name="releaseDetails">dictionary object with release database field names and values</param>
        /// <param name="releaseFolderPath">Folder path to create the release</param>
        /// <returns>TDAPIOLELib.Release Object</returns>
        public TDAPIOLELib.Release Create(Dictionary<String, String> releaseDetails, String releaseFolderPath)
        {
            TDAPIOLELib.ReleaseFolder releaseFolder = releaseFolders.GetNodeObject(releaseFolderPath);
            return Create(releaseDetails, releaseFolder);
        }

        /// <summary>
        /// Creates a release in ALM, you must pass values for all required fields
        /// </summary>
        /// <param name="releaseDetails">dictionary object with release field names and values</param>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <returns>TDAPIOLELib.Release Object</returns>
        public TDAPIOLELib.Release Create(Dictionary<String, String> releaseDetails, TDAPIOLELib.ReleaseFolder releaseFolder)
        {
            TDAPIOLELib.Release release;
            TDAPIOLELib.ReleaseFactory releaseFactory = releaseFolder.ReleaseFactory;
            release = releaseFactory.AddItem(System.DBNull.Value);

            foreach (KeyValuePair<String, String> kvp in releaseDetails)
            {
                release[kvp.Key] = kvp.Value;
            }
            release.Post();

            return release;
        }

        /// <summary>
        /// Updates a field value for release
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <param name="fieldName">Database field name</param>
        /// <param name="newValue">New field value</param>
        /// <param name="Post">Save changes</param>
        /// <returns>True if successfull</returns>
        public Boolean UpdateFieldValue(TDAPIOLELib.Release release, String fieldName, String newValue, Boolean Post = true)
        {
            release[fieldName.ToUpper()] = newValue;
            if (Post)
            release.Post();
            return true;
        }

        /// <summary>
        /// Delete a release
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <returns>True if successfull</returns>
        public Boolean Delete(TDAPIOLELib.Release release)
        {
            TDAPIOLELib.ReleaseFactory releaseFactory = tDConnection.ReleaseFactory;
            releaseFactory.RemoveItem(release.ID);
            return true;
        }

        /// <summary>
        /// Get the list of cycles for the release
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <returns>True if Successfull</returns>
        public TDAPIOLELib.List GetCycles(TDAPIOLELib.Release release)
        {
           return release.CycleFactory.NewList("");
        }

        /// <summary>
        /// Get Cycle object with Cycle Name
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <param name="CycleName">Name of the Cycle</param>
        /// <returns>TDAPIOLELib.Cycle Object</returns>
        public TDAPIOLELib.Cycle GetCycleWithName(TDAPIOLELib.Release release, String CycleName)
        {
            foreach (TDAPIOLELib.Cycle CYC in release.CycleFactory.NewList(""))
            {
                if (CYC.Name == CycleName)
                {
                    return CYC;
                }
            }
            return null;
        }

        /// <summary>
        /// Get the Folder path for a release
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <returns>Path of the release</returns>
        public String GetPath(TDAPIOLELib.Release release)
        {
            return releaseFolders.GetPath(release.Parent);
        }

        /// <summary>
        /// Get the release object using release ID
        /// </summary>
        /// <param name="releaseID"></param>
        /// <returns></returns>
        public TDAPIOLELib.Release GetObjectWithID(int releaseID)
        {
            TDAPIOLELib.ReleaseFactory releaseFactory = tDConnection.ReleaseFactory;
            return releaseFactory[releaseID];
        }

        /*public TDAPIOLELib.Release GetObjectWithID(int releaseID)
        {
            TDAPIOLELib.ReleaseFactory releaseFactory = tDConnection.ReleaseFactory;
            TDAPIOLELib.TDFilter OTDFilter = releaseFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List releaseList;

            try
            {
                OTDFilter["REL_ID"] = Convert.ToString(releaseID);
                releaseList = releaseFactory.NewList(OTDFilter.Text);

                if (releaseList != null && releaseList.Count == 1)
                {
                    return releaseList[1];
                }
                else
                {
                    throw (new Exception("Unable to find release with ID : " + releaseID));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }*/

        /// <summary>
        /// Download Release attachments
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <param name="attachmentDownloadPath">Path to download attachments</param>
        /// <returns>True if successfull</returns>
        public Boolean DownloadAttachments(TDAPIOLELib.Release release, String attachmentDownloadPath)
        {
            return Utilities.DownloadAttachments(release.Attachments, attachmentDownloadPath);
        }

        /// <summary>
        /// Add attachment to Release
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <param name="attachmentsPath">List of Attachment path</param>
        /// <returns>True if Successfull</returns>
        public Boolean AddAttachment(TDAPIOLELib.Release release, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(release.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Delete all release attachments
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.Release release)
        {
            return Utilities.DeleteAllAttachments(release.Attachments);
        }

        /// <summary>
        /// Delete release attachement by name
        /// </summary>
        /// <param name="release">TDAPIOLELib.Release Object</param>
        /// <param name="attachmentName">Name of attachment</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.Release release, String attachmentName)
        {
            return Utilities.DeleteAttachmentByName(release.Attachments, attachmentName);
        }
    }
}
