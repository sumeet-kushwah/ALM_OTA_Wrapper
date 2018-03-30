using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALM_Wrapper
{
    public class ReleaseFolders
    {
        private TDAPIOLELib.TDConnection tDConnection;

        public ReleaseFolders(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        /// <summary>
        /// Creates a release folder
        /// </summary>
        /// <param name="parentFolderPath"></param>
        /// <param name="newFolderName"></param>
        /// <returns></returns>
        public TDAPIOLELib.ReleaseFolder Create(String parentFolderPath, String newFolderName)
        {
            TDAPIOLELib.ReleaseFolder releaseFolder = GetNodeObject(parentFolderPath);

            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = releaseFolder.ReleaseFolderFactory;
            releaseFolder = releaseFolderFactory.AddItem(System.DBNull.Value);
            releaseFolder.Name = newFolderName;
            releaseFolder.Post();
            return releaseFolder;
        }

        /// <summary>
        /// Creates a release folder
        /// </summary>
        /// <param name="parentFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="newFolderName">New Folder Name</param>
        /// <returns>TDAPIOLELib.ReleaseFolder Object</returns>
        public TDAPIOLELib.ReleaseFolder Create(TDAPIOLELib.ReleaseFolder parentFolder, String newFolderName)
        {
            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = parentFolder.ReleaseFolderFactory;
            parentFolder = releaseFolderFactory.AddItem(System.DBNull.Value);
            parentFolder.Name = newFolderName;
            parentFolder.Post();
            return parentFolder;
        }

        /// <summary>
        /// Deletes a release folder
        /// </summary>
        /// <param name="parentFolderPath">Release folder path</param>
        /// <param name="folderToDelete"></param>
        /// <returns></returns>
        public Boolean Delete(String parentFolderPath, String folderToDelete)
        {
            TDAPIOLELib.ReleaseFolder releaseFolder = GetNodeObject(parentFolderPath + "\\" + folderToDelete);
            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = tDConnection.ReleaseFolderFactory;
            releaseFolderFactory.RemoveItem(releaseFolder.ID);
            return true;
        }


        /// <summary>
        /// Converts folder path string to TDAPIOLELib.ReleaseFolder
        /// <para/> retruns TDAPIOLELib.ReleaseFolder object
        /// </summary>
        /// <param name="folderPath">Releases folder path</param>
        /// <returns>TDAPIOLELib.ReleaseFolder object</returns>
        public TDAPIOLELib.ReleaseFolder GetNodeObject(String folderPath)
        {
            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = tDConnection.ReleaseFolderFactory;
            TDAPIOLELib.ReleaseFolder releaseFolder = releaseFolderFactory.Root;

            foreach (String Folder in folderPath.Split('\\'))
            {
                if (!(Folder.ToUpper() == "RELEASES"))
                {
                    releaseFolder = GetChildFolderWithName(releaseFolder, Folder);
                }
            }

            return releaseFolder;

        }

        /// <summary>
        /// Gets the list of releases under a release folder
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetReleases(TDAPIOLELib.ReleaseFolder releaseFolder)
        {
            TDAPIOLELib.List relNames = new TDAPIOLELib.List();
            TDAPIOLELib.ReleaseFactory releaseFactory = releaseFolder.ReleaseFactory;
            foreach (TDAPIOLELib.Release ORel in releaseFactory.NewList(""))
            {
                relNames.Add(ORel);
            }

            return relNames;
        }

        /// <summary>
        /// Gets the release object by release name
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="releaseName">Name of the release</param>
        /// <returns></returns>
        public TDAPIOLELib.Release GetReleaseByName(TDAPIOLELib.ReleaseFolder releaseFolder, String releaseName)
        {
            TDAPIOLELib.ReleaseFactory releaseFactory = releaseFolder.ReleaseFactory;
            TDAPIOLELib.TDFilter tDFilter = releaseFactory.Filter;
            tDFilter["REL_NAME"] = releaseName;
            TDAPIOLELib.List list = tDFilter.NewList();

            if (list.Count == 1)
                return list[1] as TDAPIOLELib.Release;

            return null;
        }

        /// <summary>
        /// Get the path of the release folder
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <returns>Path of the release folder</returns>
        public String GetPath(TDAPIOLELib.ReleaseFolder releaseFolder)
        {
            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = tDConnection.ReleaseFolderFactory;
            
            String Path = "";

            while (releaseFolder.ID != releaseFolderFactory.Root.ID)
            {
                Path = Path + releaseFolder.Name + "\\";
                releaseFolder = releaseFolder.Parent;
            }

            return reversePath(Path + "Releases");
        }

        /// <summary>
        /// To reverse the path calculated by GetPath Function
        /// </summary>
        /// <param name="Path">path of the release folder in reverse</param>
        /// <returns>Reversed Path</returns>
        private String reversePath(String Path)
        {
            String[] P = Path.Split('\\');
            Path = "";
            for (int counter = P.Length - 1; counter >= 0; counter--)
            {
                if (counter != 0)
                    Path = Path + P[counter] + "\\";
                else
                    Path = Path + P[counter];
            }

            return Path;
        }

        /// <summary>
        /// Get the Child Folder names for the release
        /// </summary>
        /// <param name="folderPath">Path of the folder to search for the release</param>
        /// <returns>List of release folder names</returns>
        public List<String> GetChildFolderNames(String folderPath)
        {
            List<String> relFolderNames = new List<string>();
            TDAPIOLELib.ReleaseFolder releaseFolder = GetNodeObject(folderPath);
            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = releaseFolder.ReleaseFolderFactory;
            foreach (TDAPIOLELib.ReleaseFolder rf in releaseFolderFactory.NewList(""))
            {
                relFolderNames.Add(rf.Name);
            }

            return relFolderNames;
        }

        /// <summary>
        /// Gets the TDAPIOLELib.ReleaseFolder object for each folder under Release Folder
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <returns>List of TDAPIOLELib.ReleaseFolder Objects</returns>
        public TDAPIOLELib.List GetChildFolders(TDAPIOLELib.ReleaseFolder releaseFolder)
        {
            TDAPIOLELib.List relFolderNames = new TDAPIOLELib.List();
            foreach (TDAPIOLELib.ReleaseFolder rf in releaseFolder.ReleaseFolderFactory.NewList(""))
            {
                relFolderNames.Add(rf);
            }

            return relFolderNames;
        }

        /// <summary>
        /// Finds a Release Folder with name
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="Name">Name of the folder to be serached</param>
        /// <returns>TDAPIOLELib.ReleaseFolder Object</returns>
        public TDAPIOLELib.ReleaseFolder GetChildFolderWithName(TDAPIOLELib.ReleaseFolder releaseFolder, String Name)
        {
            TDAPIOLELib.ReleaseFolderFactory releaseFolderFactory = releaseFolder.ReleaseFolderFactory;
            TDAPIOLELib.TDFilter tDFilter = releaseFolderFactory.Filter;
            tDFilter["RF_NAME"] = Name;

            TDAPIOLELib.List list = tDFilter.NewList();
            if (list.Count > 0)
                return list[1];
            else
                return null;
        }

        /// <summary>
        /// Renames a release folder
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="newName">New folder name</param>
        /// <returns>True if successfull</returns>
        public Boolean Rename(TDAPIOLELib.ReleaseFolder releaseFolder, String newName)
        {
            releaseFolder.Name = newName;
            releaseFolder.Post();
            return true;
        }

        /// <summary>
        /// Creates a new release folder path. 
        /// </summary>
        /// <param name="folderPath">Folder path to be created</param>
        /// <returns>Final folders TDAPIOLELib.ReleaseFolder Object</returns>
        public TDAPIOLELib.ReleaseFolder CreateNewFolderPath(String folderPath)
        {
            TDAPIOLELib.ReleaseFolder releaseFolder = tDConnection.ReleaseFolderFactory.Root;
            TDAPIOLELib.ReleaseFolder newReleaseFolder;

            foreach (String Folder in folderPath.Split('\\'))
            {
                if (!(Folder.ToUpper() == "RELEASES"))
                {
                    newReleaseFolder = GetChildFolderWithName(releaseFolder, Folder);
                    if (newReleaseFolder != null)
                    {
                        releaseFolder = newReleaseFolder;

                    }
                    else
                    {
                        releaseFolder = Create(releaseFolder, Folder);
                    }
                }
            }
            return releaseFolder;
        }

        /// <summary>
        /// Download all release folder attachments
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="attachmentDownloadPath">Path to download the attachments</param>
        /// <returns></returns>
        public Boolean DownloadAttachments(TDAPIOLELib.ReleaseFolder releaseFolder, String attachmentDownloadPath)
        {
            return Utilities.DownloadAttachments(releaseFolder.Attachments, attachmentDownloadPath);
        }

        /// <summary>
        /// Add Attachments to the release folder
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="attachmentsPath">List of Attachment paths</param>
        /// <returns>True if successfull</returns>
        public Boolean AddAttachment(TDAPIOLELib.ReleaseFolder releaseFolder, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(releaseFolder.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Delete all attachments from release folder
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <returns></returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.ReleaseFolder releaseFolder)
        {
            return Utilities.DeleteAllAttachments(releaseFolder.Attachments);
        }

        /// <summary>
        /// Deletes attachments by name
        /// </summary>
        /// <param name="releaseFolder">TDAPIOLELib.ReleaseFolder Object</param>
        /// <param name="attachmentName">attachment name</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.ReleaseFolder releaseFolder, String attachmentName)
        {
            return Utilities.DeleteAttachmentByName(releaseFolder.Attachments, attachmentName);
        }

       

    }
}
