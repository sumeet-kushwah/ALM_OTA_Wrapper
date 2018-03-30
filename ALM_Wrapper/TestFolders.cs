using System;
using System.Collections.Generic;

namespace ALM_Wrapper
{
    public class TestFolders
    {
        private TDAPIOLELib.TDConnection tDConnection;
        private Test test;

        public TestFolders(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
            this.test = new Test(tDConnection);
        }

        /// <summary>
        /// Finds All Folder names under a test plan folder.
        /// <para/>returns List of string with folder names
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns>List of string with folder names</returns>
        public List<String> GetChildFolderNames(String folderPath)
        {
            List<String> OFNames = new List<string>();
            TDAPIOLELib.SysTreeNode OSysTreeNode = GetNodeObject(folderPath);
            for (int Counter = 1; Counter <= OSysTreeNode.Count; Counter++)
            {
                OFNames.Add(OSysTreeNode.Child[Counter].Name);
            }
            return OFNames;
        }

        /// <summary>
        /// Finds All Folder names under a test plan folder.
        /// <para/>returns List of string with folder names
        /// </summary>
        /// <param name="testfolder">TDAPIOLELib.TestFolder Object</param>
        /// <returns>List of string with folder names</returns>
        public List<String> GetChildFolderNames(TDAPIOLELib.TestFolder testFolder)
        {
            List<String> List = new List<string>();
            foreach (TDAPIOLELib.TestFolder oFolder in testFolder.TestFolderFactory.NewList(""))
            {
                List.Add(oFolder["AL_DESCRIPTION"]);
            }
            return List;

        }

        /// <summary>
        /// Finds All Folders under a test plan folder.
        /// <para/>returns List of string with folder names
        /// </summary>
        /// <param name="testfolder">TDAPIOLELib.TestFolder Object</param>
        /// <returns>List of string with folder names</returns>
        public TDAPIOLELib.List GetChildFolders(TDAPIOLELib.TestFolder testFolder)
        {
            TDAPIOLELib.List List = new TDAPIOLELib.List();
            foreach (TDAPIOLELib.TestFolder oFolder in testFolder.TestFolderFactory.NewList(""))
            {
                List.Add(oFolder);
            }
            return List;
        }

        /// <summary>
        /// Creates a new test plan folder.
        /// <para/> retruns true if successfull
        /// </summary>
        /// <param name="parentFolderPath">path of the test plan folder where new folder should be created</param>
        /// <param name="newFolderName">name of the new folder</param>
        /// <returns>TDAPIOLELib.TestFolder Object for new folder</returns>
        public TDAPIOLELib.TestFolder Create(String parentFolderPath, String newFolderName)
        {
            return Create(GetFolderObject(parentFolderPath), newFolderName);
        }


        /// <summary>
        /// Creates new test folder
        /// </summary>
        /// <param name="testFolder">TDAPIOLELib.TestFolder Object</param>
        /// <param name="newFolderName">New Folder name</param>
        /// <returns>TDAPIOLELib.TestFolder Object for new folder</returns>
        public TDAPIOLELib.TestFolder Create(TDAPIOLELib.TestFolder testFolder, String newFolderName)
        {
            TDAPIOLELib.TestFolderFactory testFolderFactory = testFolder.TestFolderFactory;
            TDAPIOLELib.TestFolder testFolder1 = testFolderFactory.AddItem(System.DBNull.Value);
            testFolder1["AL_DESCRIPTION"] = newFolderName;
            testFolder1.Post();
            return testFolder1;
        }

        /// <summary>
        /// Finds tests under test plan folder.
        /// <para/> returns TDAPIOLELib.List. Each item of this list can be converted to TDAPIOLELib.Test Object
        /// </summary>
        /// <param name="folderPath">Test plan folder path</param>
        /// <returns></returns>
        public TDAPIOLELib.List GetTests(String folderPath)
        {
            ///Old SysTreeNode Object
            //TDAPIOLELib.TreeManager OTManager = tDConnection.TreeManager;
            //var OTFolder = OTManager.get_NodeByPath(folderPath);
            //TDAPIOLELib.TestFactory OTFactory = OTFolder.TestFactory;
            //return OTFactory.NewList("");
            return GetTests(GetFolderObject(folderPath));
        }

        /// <summary>
        /// Finds tests under test plan folder.
        /// <para/> returns TDAPIOLELib.List. Each item of this list can be converted to TDAPIOLELib.Test Object
        /// </summary>
        /// <param name="testFolder">TDAPIOLELib.TestFolder</param>
        /// <returns></returns>
        public TDAPIOLELib.List GetTests(TDAPIOLELib.TestFolder testFolder)
        {
            TDAPIOLELib.TestFactory testFactory = testFolder.TestFactory;
            return testFactory.NewList("");
        }

        /// <summary>
        /// Finds all tests recursively under a test folder
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public TDAPIOLELib.List GetTestsRecurcively(String folderPath)
        {
            TDAPIOLELib.TreeManager OTManager = tDConnection.TreeManager;
            TDAPIOLELib.TestFactory OTFactory = tDConnection.TestFactory;
            TDAPIOLELib.TDFilter tDFilter = OTFactory.Filter;
            tDFilter["TS_SUBJECT"] = "^\\" + folderPath + "^";

            return OTFactory.NewList(tDFilter.Text);

        }

        /// <summary>
        /// Delete a test plan folder.
        /// <para/> retruns true if successfull
        /// </summary>
        /// <param name="parentFolderPath">Path of the test plan to seach for the folder</param>
        /// <param name="MovetoUnattached">Move tests under the folder to be deleted to unattached folder</param>
        /// <returns>true if successfull</returns>
        public Boolean Delete(String parentFolderPath, Boolean MovetoUnattached = false)
        {
            ///Old SysTreeNode Functionality
            //TDAPIOLELib.SysTreeNode OSysTreeNode = GetNodeObject(parentFolderPath);
            //TDAPIOLELib.SysTreeNode NodeToBeDeleted = OSysTreeNode.FindChildNode(deleteFolderName) as TDAPIOLELib.SysTreeNode;

            //if (NodeToBeDeleted.NodeID > 0)
            //{
            //    OSysTreeNode.RemoveNode(NodeToBeDeleted);
            //    return true;
            //}
            //else
            //{
            //    return false;
            //}

            return Delete(GetFolderObject(parentFolderPath), MovetoUnattached);
            
        }

        /// <summary>
        /// Delete a test plan folder.
        /// <para/> retruns true if successfull
        /// </summary>
        /// <param name="testFolder">TDAPIOLELib.TestFolder Object for the folder to be deleted</param>
        /// <param name="MovetoUnattached">Move tests under the folder to be deleted to unattached folder</param>
        /// <returns>true if successfull</returns>
        public Boolean Delete(TDAPIOLELib.TestFolder testFolder, Boolean MovetoUnattached = false)
        {
            TDAPIOLELib.TestFolderFactory testFolderFactory = tDConnection.TestFolderFactory;
            if (MovetoUnattached == false)
                testFolderFactory.RemoveItem(testFolder.ID);
            else
                testFolderFactory.RemoveItemAndMoveTestsToUnattached(testFolder.ID);

            return true;
        }

        /// <summary>
        /// Converts folder path string to TDAPIOLELib.SysTreeNode
        /// <para/> retruns TDAPIOLELib.SysTreeNode object
        /// </summary>
        /// <param name="folderPath">test plan folder path</param>
        /// <returns>TDAPIOLELib.SysTreeNode object</returns>
        public TDAPIOLELib.SysTreeNode GetNodeObject(String folderPath)
        {
            TDAPIOLELib.TreeManager OTManager = tDConnection.TreeManager;
            return OTManager.NodeByPath[folderPath];
             
        }

        /// <summary>
        /// Get the TDAPIOLELib.TestFolder object from the folder path
        /// </summary>
        /// <param name="folderPath">Path of the test folder</param>
        /// <returns>TDAPIOLELib.TestFolder Object</returns>
        public TDAPIOLELib.TestFolder GetFolderObject(String folderPath)
        {
            TDAPIOLELib.TestFolderFactory testFolderFactory = tDConnection.TestFolderFactory;
            TDAPIOLELib.TestFolder testFolder = testFolderFactory.Root;
            TDAPIOLELib.TDFilter tDFilter;


            foreach (String foldername in folderPath.Split('\\'))
            {
                if (foldername.ToUpper() != "SUBJECT")
                {
                    testFolderFactory = testFolder.TestFolderFactory;
                    tDFilter = testFolderFactory.Filter;
                    tDFilter["AL_DESCRIPTION"] = foldername;
                    testFolder = tDFilter.NewList()[1];
                }
            }


            return testFolder;
        }

        

        /// <summary>
        /// Creates new folder path. e.g. if path is passed as Subject\Dummy1\Dummy2\Dummy3 and Only subject folder exists in ALM then this function will create Dummy1, Dummy2 and Dummy3 folders.
        /// <para/> true if successfull
        /// </summary>
        /// <param name="folderPath">folder path</param>
        /// <returns>true if successfull</returns>
        public TDAPIOLELib.TestFolder CreateNewFolderPath(String folderPath)
        {
            TDAPIOLELib.TreeManager OTManager = tDConnection.TreeManager;
            String PathChecked = "Subject";
            TDAPIOLELib.SysTreeNode oTestFolder = GetNodeObject("subject");

            foreach (String Folder in folderPath.Split('\\'))
            {
                if (!(Folder.ToUpper() == "SUBJECT"))
                {
                    try
                    {
                        oTestFolder = GetNodeObject(PathChecked + "\\" + Folder);
                        PathChecked = PathChecked + "\\" + Folder;
                    }
                    catch
                    {
                        oTestFolder = GetNodeObject(PathChecked);
                        oTestFolder.AddNode(Folder);
                        PathChecked = PathChecked + "\\" + Folder;
                    }
                }
            }

            TDAPIOLELib.TestFolder testFolder = GetFolderObject(PathChecked);
            return testFolder;
        }

        
        /// <summary>
        /// Rename test plan folder
        /// <para/> retruns true if successfull
        /// </summary>
        /// <param name="parentFolderPath">path of the test plan to seach for the folder</param>
        /// <param name="currentFolderName">folder to be renamed</param>
        /// <param name="newFolderName">new folder name</param>
        /// <returns></returns>
        public Boolean Rename(String parentFolderPath, String newFolderName)
        {
            return Rename(GetFolderObject(parentFolderPath), newFolderName);
            ////SysTreeNode Functionality
            //TDAPIOLELib.SysTreeNode OParentFolder = GetNodeObject(parentFolderPath);
            //TDAPIOLELib.SysTreeNode OCurrentFolder = OParentFolder.FindChildNode(currentFolderName) as TDAPIOLELib.SysTreeNode;
            //OCurrentFolder.Name = newFolderName;
            //OCurrentFolder.Post();
            
        }

        /// <summary>
        /// Rename test plan folder
        /// </summary>
        /// <param name="testFolder">TDAPIOLELib.TestFolder Object</param>
        /// <param name="NewName">New Folder name</param>
        /// <returns>True if successfull</returns>
        public Boolean Rename(TDAPIOLELib.TestFolder testFolder, String NewName)
        {
            testFolder["AL_DESCRIPTION"] = NewName;
            testFolder.Post();
            return true;
        }
    }
}
