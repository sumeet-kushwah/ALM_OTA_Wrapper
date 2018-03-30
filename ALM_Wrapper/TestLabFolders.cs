using System;
using System.Collections.Generic;

namespace ALM_Wrapper
{
    public class TestLabFolders
    {
        private TDAPIOLELib.TDConnection tDConnection;

        public TestLabFolders(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        /// <summary>
        /// Find test sets under TestLab Folder
        /// <para/>return TDAPIOLELib.List. Each item from the list can be converted to TDAPIOLELib.TestSet
        /// </summary>
        /// <param name="folderPath">TestLab folder path</param>
        /// <returns></returns>
        public TDAPIOLELib.List FindTestSets(String folderPath)
        {
            TDAPIOLELib.TestSetTreeManager testSetTreeManager = tDConnection.TestSetTreeManager;
            TDAPIOLELib.TestSetFolder testSetFolder = testSetTreeManager.NodeByPath[folderPath];

            TDAPIOLELib.TestSetFactory testSetFactory = testSetFolder.TestSetFactory;
            TDAPIOLELib.List OTestSetList = testSetFactory.NewList("");

            return OTestSetList;
        }

        /// <summary>
        /// Find test sets under a test set folder
        /// </summary>
        /// <param name="testSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <returns>TDAPIOLELib.List Object containing test set objects</returns>
        public TDAPIOLELib.List FindTestSets(TDAPIOLELib.TestSetFolder testSetFolder)
        {
            TDAPIOLELib.TestSetFactory testSetFactory = testSetFolder.TestSetFactory;
            return testSetFactory.NewList("");

        }

        /// <summary>
        /// Find Cycle for the test set folder
        /// </summary>
        /// <param name="folderPath">path of the test set folder</param>
        /// <returns>TDAPIOLELib.Cycle Object</returns>
        public TDAPIOLELib.Cycle GetCycle(String folderPath)
        {
            TDAPIOLELib.TestSetFolder testSetFolder = GetTestSetFolder(folderPath);
            return testSetFolder.TargetCycle as TDAPIOLELib.Cycle;
        }

        /// <summary>
        /// Find Cycle for the test set folder
        /// </summary>
        /// <param name="testSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <returns>TDAPIOLELib.Cycle Object</returns>
        public TDAPIOLELib.Cycle GetCycle(TDAPIOLELib.TestSetFolder testSetFolder)
        {
            return testSetFolder.TargetCycle as TDAPIOLELib.Cycle;
        }

        /// <summary>
        /// Returns the path for the test set folder
        /// </summary>
        /// <param name="testSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <returns>Path of the test set folder</returns>
        public String GetPath(TDAPIOLELib.TestSetFolder testSetFolder)
        {
            return testSetFolder.Path;
        }

        /// <summary>
        /// Get TestSetFolder object from test set folder path
        /// <para/>returns TDAPIOLELib.TestSetFolder Object
        /// </summary>
        /// <param name="testSetFolderPath">test set folder path from test lab</param>
        /// <returns>TDAPIOLELib.TestSetFolder Object</returns>
        public TDAPIOLELib.TestSetFolder GetTestSetFolder(String testSetFolderPath)
        {
            TDAPIOLELib.TestSetTreeManager testSetTreeManager = tDConnection.TestSetTreeManager;
            return testSetTreeManager.NodeByPath[testSetFolderPath];
        }

        /// <summary>
        /// Find folders under a test lab folder
        /// <para/>returns list if string with folder names
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns>returns list if string with folder names</returns>
        public List<String> FindChildFolders(String folderPath)
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
        /// Find child folders under a test plan folder
        /// <para>returns TDAPIOLELib.List Object containing the list of Test Set folder object</para>
        /// </summary>
        /// <param name="testSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <returns>TDAPIOLELib.List Object containing the list of Test Set folder object</returns>
        public TDAPIOLELib.List FindChildFolders(TDAPIOLELib.TestSetFolder testSetFolder)
        {
            TDAPIOLELib.List List = new TDAPIOLELib.List();
            
            for (int Counter = 1; Counter <= testSetFolder.Count; Counter++)
            {
                List.Add(testSetFolder.Child[Counter]);
            }
            return List;
        }

        public TDAPIOLELib.TestSetFolder FindChildFolderByName(TDAPIOLELib.TestSetFolder parentFolder, String ChildFolderToBeSearched)
        {
            return (TDAPIOLELib.TestSetFolder)parentFolder.FindChildNode(ChildFolderToBeSearched);
        }


        /// <summary>
        /// Get SysTreeNode Object from Test lab folder path
        /// <para/>return TDAPIOLELib.SysTreeNode Object
        /// </summary>
        /// <param name="FolderPath"></param>
        /// <returns>TDAPIOLELib.SysTreeNode Object</returns>
        public TDAPIOLELib.SysTreeNode GetNodeObject(String FolderPath)
        {
            TDAPIOLELib.TestSetTreeManager OTSTManager = tDConnection.TestSetTreeManager;
            return OTSTManager.NodeByPath[FolderPath];

        }

        /// <summary>
        /// Create a new folder under test lab
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="ParentFolderPath">Parent folder name</param>
        /// <param name="NewFolderName">New folder name</param>
        /// <returns>true if successfull</returns>
        public Boolean Create(String ParentFolderPath, String NewFolderName)
        {
            TDAPIOLELib.SysTreeNode OSysTreeNode = GetNodeObject(ParentFolderPath);
            OSysTreeNode.AddNode(NewFolderName);
            return true;
        }

        public TDAPIOLELib.TestSetFolder Create(TDAPIOLELib.TestSetFolder testSetFolder, String NewFolderName)
        {
            return (TDAPIOLELib.TestSetFolder)testSetFolder.AddNode(NewFolderName); 
        }

        /// <summary>
        /// Deletes a folder in test lab
        /// </summary>
        /// <param name="ParentFolderPath">Parent folder path</param>
        /// <param name="DeleteFolderName">folder to be deleted</param>
        /// <returns></returns>
        public Boolean Delete(String ParentFolderPath, String DeleteFolderName)
        {
            TDAPIOLELib.SysTreeNode OSysTreeNode = GetNodeObject(ParentFolderPath);
            TDAPIOLELib.SysTreeNode NodeToBeDeleted = OSysTreeNode.FindChildNode(DeleteFolderName) as TDAPIOLELib.SysTreeNode;

            if (NodeToBeDeleted.NodeID > 0)
            {
                OSysTreeNode.RemoveNode(NodeToBeDeleted);
            }
            return true;
        }

        public Boolean Delete(TDAPIOLELib.TestSetFolder testSetFolder)
        {
            TDAPIOLELib.TestSetFolder testSetFolderFather = (TDAPIOLELib.TestSetFolder)testSetFolder.Father;
            testSetFolderFather.RemoveNode(testSetFolder);
            return true;
        }

        /// <summary>
        /// Create a new folder path in test lab
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="FolderPath">Folder path to be created. All missing folder will be created</param>
        /// <returns>true if successfull</returns>
        public Boolean CreateNewFolderPath(String FolderPath)
        {
            String PathChecked = "Root";
            TDAPIOLELib.SysTreeNode oTestFolder = GetNodeObject("root");
            foreach (String Folder in FolderPath.Split('\\'))
            {
                if (!(Folder.ToUpper() == "ROOT"))
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

            return true;
        }

        /// <summary>
        /// Renames a Folder
        /// <para/>Returns true if successfull
        /// </summary>
        /// <param name="ParentFolderPath">Parent folder path</param>
        /// <param name="CurrentFolderName">Current folder name</param>
        /// <param name="NewFolderName">New folder name</param>
        /// <returns>true if successfull</returns>
        public Boolean Rename(String ParentFolderPath, String CurrentFolderName, String NewFolderName)
        {
            TDAPIOLELib.SysTreeNode OParentFolder = GetNodeObject(ParentFolderPath);
            TDAPIOLELib.SysTreeNode OCurrentFolder = OParentFolder.FindChildNode(CurrentFolderName) as TDAPIOLELib.SysTreeNode;
            OCurrentFolder.Name = NewFolderName;
            OCurrentFolder.Post();
            return true;
        }

        public Boolean Rename(TDAPIOLELib.TestSetFolder testSetFolder, String NewFolderName)
        {
            testSetFolder.Name = NewFolderName;
            testSetFolder.Post();
            return true;
        }

        /// <summary>
        /// Assign Cycle to test lab folder
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="testLabFolderPath">testplan folder path</param>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <returns>true if successfull</returns>
        public Boolean AssignCycle(String testLabFolderPath, TDAPIOLELib.Cycle cycle)
        {
            return AssignCycle(GetTestSetFolder(testLabFolderPath), cycle);
            //TDAPIOLELib.TestSetFolder testSetFolder = GetTestSetFolder(testLabFolderPath);
            //testSetFolder.TargetCycle = cycle.ID;
            //return true;
        }

        /// <summary>
        /// Assign Cycle to test lab folder
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="testSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <returns>true if successfull</returns>
        public Boolean AssignCycle(TDAPIOLELib.TestSetFolder testSetFolder, TDAPIOLELib.Cycle cycle)
        {
            testSetFolder.TargetCycle = cycle.ID;
            return true;
        }

        /// <summary>
        /// Download attachments from a TestSet Folder
        /// </summary>
        /// <param name="TestSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <param name="attachmentDownloadPath">Path to download attachments</param>
        /// <returns>True if successfull</returns>
        public Boolean DownloadAttachments(TDAPIOLELib.TestSetFolder TestSetFolder, String attachmentDownloadPath)
        {
            return Utilities.DownloadAttachments(TestSetFolder.Attachments, attachmentDownloadPath);
        }

        /// <summary>
        /// Add attachments to TestSet Folder
        /// </summary>
        /// <param name="TestSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <param name="attachmentsPath">List of attachment path</param>
        /// <returns>true if Successfull</returns>
        public Boolean AddAttachment(TDAPIOLELib.TestSetFolder TestSetFolder, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(TestSetFolder.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Deletes all attachments from TestSet Folder
        /// </summary>
        /// <param name="TestSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.TestSetFolder TestSetFolder)
        {
            return Utilities.DeleteAllAttachments(TestSetFolder.Attachments);
        }

        /// <summary>
        /// Delete attachment by name
        /// </summary>
        /// <param name="TestSetFolder">TDAPIOLELib.TestSetFolder Object</param>
        /// <param name="attachmentName">Attachment name to be deleted</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.TestSetFolder TestSetFolder, String attachmentName)
        {
            return Utilities.DeleteAttachmentByName(TestSetFolder.Attachments, attachmentName);
        }
    }
}
