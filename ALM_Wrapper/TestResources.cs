using System;
using TDAPIOLELib;

namespace ALM_Wrapper
{
    public class TestResources
    {
        private TDAPIOLELib.TDConnection tDConnection;
        public enum ResourceUploadType { File = 0, Folder = 1 };
        public enum ResourceType {ApplicationArea = 0,  DataTable = 1, EnvironmentVariables = 2, FunctionLibrary = 3, RecoveryScenario = 4, SharedObjectRepository = 5, TestResource = 6, TestingActivity = 7};

        public TestResources(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        public TDAPIOLELib.QCResourceFolder CreateFolder(TDAPIOLELib.QCResourceFolder ParentFolder, String NewFolderName)
        {
            TDAPIOLELib.QCResourceFolderFactory qCResourceFolderFactory = ParentFolder.QCResourceFolderFactory;
            TDAPIOLELib.QCResourceFolder qCResourceFolder = qCResourceFolderFactory.AddItem(System.DBNull.Value);
            qCResourceFolder.Name = NewFolderName;
            qCResourceFolder.Post();

            return qCResourceFolder;
        }

        public TDAPIOLELib.QCResourceFolder CreateFolderPath(String FolderPath)
        {
            TDAPIOLELib.TreeManager OTManager = tDConnection.TreeManager;
            String PathChecked = "Resources";
            TDAPIOLELib.QCResourceFolder QCFolder = GetFolderFromPath("Resources");

            foreach (String Folder in FolderPath.Split('\\'))
            {
                if (!(Folder.ToUpper() == "RESOURCES"))
                {
                    try
                    {
                        QCFolder = GetFolderFromPath(PathChecked + "\\" + Folder);
                        PathChecked = PathChecked + "\\" + Folder;
                    }
                    catch
                    {
                        QCFolder = GetFolderFromPath(PathChecked);
                        CreateFolder(QCFolder, Folder);
                        PathChecked = PathChecked + "\\" + Folder;
                    }
                }
            }
            return QCFolder;
        }

        public Boolean RenameFolder(TDAPIOLELib.QCResourceFolder qCResourceFolder, String NewName)
        {
            qCResourceFolder.Name = NewName;
            qCResourceFolder.Post();
            return true;
        }

        public Boolean DeleteFolder(TDAPIOLELib.QCResourceFolder qCResourceFolder)
        {
            tDConnection.QCResourceFolderFactory.RemoveItem(qCResourceFolder.ID);
            return true;
        }

        public TDAPIOLELib.List GetChildFolderList(TDAPIOLELib.QCResourceFolder qCResourceFolder)
        {
            TDAPIOLELib.QCResourceFolderFactory qCResourceFolderFactory = qCResourceFolder.QCResourceFolderFactory;
            return qCResourceFolderFactory.NewList("");
        }

        public System.Collections.Generic.List<String> GetChildFolderNames(TDAPIOLELib.QCResourceFolder qCResourceFolder)
        {
            TDAPIOLELib.QCResourceFolderFactory qCResourceFolderFactory = qCResourceFolder.QCResourceFolderFactory;
            System.Collections.Generic.List<String> list = new System.Collections.Generic.List<string>();
            foreach (TDAPIOLELib.QCResourceFolder folder in qCResourceFolderFactory.NewList(""))
            {
                list.Add(folder.Name);
            }
            return list;
        }

        public TDAPIOLELib.List GetResources(TDAPIOLELib.QCResourceFolder qCResourceFolder)
        {
            TDAPIOLELib.QCResourceFactory qCResourceFactory = qCResourceFolder.QCResourceFactory;
            return qCResourceFactory.NewList("");
        }

        public System.Collections.Generic.List<String> GetResourceNames(TDAPIOLELib.QCResourceFolder qCResourceFolder)
        {
            TDAPIOLELib.QCResourceFactory qCResourceFactory = qCResourceFolder.QCResourceFactory;
            System.Collections.Generic.List<String> list = new System.Collections.Generic.List<string>();
            foreach (TDAPIOLELib.QCResourceFolder folder in qCResourceFactory.NewList(""))
            {
                list.Add(folder.Name);
            }
            return list;
        }

        public TDAPIOLELib.QCResourceFolder GetFolderFromPath(String FolderPath)
        {
            TDAPIOLELib.QCResourceFolderFactory qCResourceFolderFactory = tDConnection.QCResourceFolderFactory;
            TDAPIOLELib.TDFilter tDFilter = qCResourceFolderFactory.Filter;
            TDAPIOLELib.QCResourceFolder qCResourceFolder = tDConnection.QCResourceFolderFactory[1];

            foreach (String OFolder in FolderPath.Split('\\'))
            {
                tDFilter = qCResourceFolderFactory.Filter;
                tDFilter["RFO_NAME"] = OFolder;
                qCResourceFolder = tDFilter.NewList()[1];
                qCResourceFolderFactory = qCResourceFolder.QCResourceFolderFactory;
            }

            return qCResourceFolder;
        }

        public TDAPIOLELib.QCResource CreateResource(TDAPIOLELib.QCResourceFolder qCResourceFolder,String Name, String ResourceFilePath, ResourceUploadType resourceUploadType, ResourceType resourceType)
        {

            //Check if Resource type is folder and should be uploaded or not
            if ((resourceUploadType == ResourceUploadType.Folder) && (resourceType!=TestResources.ResourceType.TestResource))
            {
                throw (new Exception("Folder cant be uploaded to " + GetResourceType(resourceType) + " resource type"));
            }

            TDAPIOLELib.QCResourceFactory qCResourceFactory = qCResourceFolder.QCResourceFactory;
            TDAPIOLELib.QCResource qCResource = qCResourceFactory.AddItem(System.DBNull.Value);
            qCResource.Name = Name;
            qCResource.ResourceType = GetResourceType(resourceType);


            String ResourceType = GetResourceType(resourceType);
            
            if (resourceUploadType == ResourceUploadType.File)
            {
                UploadFileToResource(qCResource, ResourceFilePath);
               
            }
            else
            {
                UploadFolderToResource(qCResource, ResourceFilePath);
               
            }
            

            return qCResource;

        }

        private String GetResourceType(ResourceType resourceType)
        {
            String Result = "";
            switch (resourceType)
            {
                case ResourceType.ApplicationArea:
                    Result = "Application Area";
                    break;
                case ResourceType.DataTable:
                    Result = "Data Table";
                    break;
                case ResourceType.EnvironmentVariables:
                    Result = "Environment Variables";
                    break;
                case ResourceType.FunctionLibrary:
                    Result = "Function Library";
                    break;
                case ResourceType.RecoveryScenario:
                    Result = "Recovery Scenario";
                    break;
                case ResourceType.SharedObjectRepository:
                    Result = "Shared Object Repository";
                    break;
                case ResourceType.TestingActivity:
                    Result = "Testing Activity";
                    break;
                case ResourceType.TestResource:
                    Result = "Test Resource";
                    break;
            }
            return Result;
        }

        public Boolean UploadFileToResource(TDAPIOLELib.QCResource qCResource, String FilePath)
        {
            qCResource.FileName = System.IO.Path.GetFileName(FilePath);
            qCResource.Post();

            TDAPIOLELib.IResourceStorage resourceStorage = (IResourceStorage)qCResource;
            resourceStorage.UploadResource(System.IO.Path.GetDirectoryName(FilePath), true);
            return true;
        }

        public Boolean UploadFolderToResource(TDAPIOLELib.QCResource qCResource, String FolderPath)
        {
            qCResource.FileName = new System.IO.DirectoryInfo(FolderPath).Name;
            qCResource.Post();

            TDAPIOLELib.IResourceStorage resourceStorage = (IResourceStorage)qCResource;
            resourceStorage.UploadResource(System.IO.Path.GetDirectoryName(FolderPath), true);
            return true;
        }

        public TDAPIOLELib.QCResource CreateResource(TDAPIOLELib.QCResourceFolder qCResourceFolder, String Name)
        {
            TDAPIOLELib.QCResourceFactory qCResourceFactory = qCResourceFolder.QCResourceFactory;
            TDAPIOLELib.QCResource qCResource = qCResourceFactory.AddItem(System.DBNull.Value);
            qCResource.Name = Name;
            qCResource.Post();
            return qCResource;

        }

        public Boolean DownloadResources(TDAPIOLELib.QCResource qCResource, String FolderPath)
        {
            TDAPIOLELib.IResourceStorage resourceStorage = (IResourceStorage)qCResource;
            resourceStorage.DownloadResource(FolderPath, true);
            return true;
        }
    }
}
