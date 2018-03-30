using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALM_Wrapper
{
    public class TestSet
    {
        private TDAPIOLELib.TDConnection tDConnection;
        private TestLabFolders testLabFolders;

        public TestSet(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
            testLabFolders = new TestLabFolders(tDConnection);
        }

        /// <summary>
        /// Create a new test set in the test plan
        /// <para/> returns TDAPIOLELib.TestSet object
        /// </summary>
        /// <param name="testSetDetails">Test set database fields with values</param>
        /// <param name="testSetFolderPath">Test lab Folder path for test set</param>
        /// <returns></returns>
        public TDAPIOLELib.TestSet Create(Dictionary<String, String> testSetDetails, String testSetFolderPath)
        {
            TDAPIOLELib.TestSetFactory testSetFactory = testLabFolders.GetTestSetFolder(testSetFolderPath).TestSetFactory;
            TDAPIOLELib.TestSet testSet = testSetFactory.AddItem(System.DBNull.Value);

            foreach (KeyValuePair<string, string> kvp in testSetDetails)
            {
                testSet[kvp.Key.ToUpper()] = kvp.Value;
            }

            //Post the test to ALM
            testSet.Post();
            return testSet;
        }

        /// <summary>
        /// Add Attachment to test set
        /// <para/>return true if successfull
        /// </summary>
        /// <param name="testSet">test set object</param>
        /// <param name="attachMentsPath">list of all attachments</param>
        /// <returns>true if successfull</returns>
        public Boolean AddAttachments(TDAPIOLELib.TestSet testSet, List<String> attachMentsPath)
        {
            return Utilities.AddAttachment(testSet.Attachments, attachMentsPath);
        }

        /// <summary>
        /// Delete all attachmentsfrom test set
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet">test set object</param>
        /// <returns>returns true if successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.TestSet testSet)
        {
            return Utilities.DeleteAllAttachments(testSet.Attachments);
        }

        /// <summary>
        /// Download all testset attachment
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <param name="downloadPath">Path to dowsnload attachments</param>
        /// <returns>True if Successfull</returns>
        public Boolean DownloadAttachements(TDAPIOLELib.TestSet testSet, String downloadPath)
        {
            return Utilities.DownloadAttachments(testSet.Attachments, downloadPath);
        }

        /// <summary>
        /// Delete test set attachment by name
        /// <para/> retruns true if successfull
        /// </summary>
        /// <param name="testSet">test set object</param>
        /// <param name="attachmentName">attachemnt name</param>
        /// <returns>true if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.TestSet testSet, String attachmentName)
        {
            return Utilities.DeleteAttachmentByName(testSet.Attachments, attachmentName);
        }

        /// <summary>
        /// Get test instances from testset
        /// <para/>returns TDAPIOLELib.List. Each object from this list can be converted to TDAPIOLELib.TSTest
        /// </summary>
        /// <param name="testSet">test set object</param>
        /// <returns>TDAPIOLELib.List. Each object from this list can be converted to TDAPIOLELib.TSTest</returns>
        public TDAPIOLELib.List GetAllTestInstance(TDAPIOLELib.TestSet testSet)
        {
            TDAPIOLELib.TSTestFactory tSTestFactory = testSet.TSTestFactory;
            return tSTestFactory.NewList("");
        }

        /// <summary>
        /// Get last run object for TSTest object
        /// <para/>TDAPIOLELib.Run object
        /// </summary>
        /// <param name="tSTest">TSTest Object</param>
        /// <returns>TDAPIOLELib.Run object</returns>
        public TDAPIOLELib.Run GetLastRunDetails(TDAPIOLELib.TSTest tSTest)
        {
            return tSTest.LastRun;
        }

        /// <summary>
        /// pass all tests under a testset
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet">testset object</param>
        /// <returns>returns true if successfull</returns>
        public Boolean PassAllTests(TDAPIOLELib.TestSet testSet)
        {
            foreach (TDAPIOLELib.TSTest tsTest in GetAllTestInstance(testSet))
            {
                RunTestSetTest(tsTest.RunFactory, "Passed");
            }

            return true;
        }

        /// <summary>
        /// Fail all tests under a testset
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet">testset object</param>
        /// <returns>returns true if successfull</returns>
        public Boolean FailAllTests(TDAPIOLELib.TestSet testSet)
        {
            foreach (TDAPIOLELib.TSTest tsTest in GetAllTestInstance(testSet))
            {
                RunTestSetTest(tsTest.RunFactory, "Failed");
            }

            return true;
        }

        /// <summary>
        /// Execute tests under a testset
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="executionDetails">TSTest.ID and Status must be added to this dictionary object</param>
        /// <returns>true if successfull</returns>
        public Boolean ExecuteTests(Dictionary<int, String> executionDetails)
        {
            TDAPIOLELib.TSTest tSTest;
            foreach (KeyValuePair<int, String> kvp in executionDetails)
            {
                tSTest = GetTSTestWithID(kvp.Key);
                RunTestSetTest(tSTest.RunFactory, kvp.Value);
            }

            return true;
        }

        /// <summary>
        /// Get TSTest object with ID
        ///<para/> returns TDAPIOLELib.TSTest object
        /// </summary>
        /// <param name="id">ID of the TSTest</param>
        /// <returns>TDAPIOLELib.TSTest object</returns>
        public TDAPIOLELib.TSTest GetTSTestWithID(int id)
        {
            TDAPIOLELib.TSTestFactory tSTestFactory = tDConnection.TSTestFactory as TDAPIOLELib.TSTestFactory;
            TDAPIOLELib.TDFilter tDFilter = tSTestFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List testSetList;

            TDAPIOLELib.TSTest tSTest;

            try
            {
                tDFilter["TC_TESTCYCL_ID"] = Convert.ToString(id);
                testSetList = tSTestFactory.NewList(tDFilter.Text);

                if (testSetList != null && testSetList.Count == 1)
                {
                    tSTest = testSetList[1];
                    return tSTest;
                }
                else
                {
                    throw (new Exception("Unable to find test Set instance with ID : " + id));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Execute a single TSTest
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="tSTest">TDAPIOLELib.TSTest Object</param>
        /// <param name="Status">Execution Status</param>
        /// <returns>true if successfull</returns>
        public Boolean ExecuteSingleTest(TDAPIOLELib.TSTest tSTest, String Status)
        {
            return RunTestSetTest(tSTest.RunFactory, Status);
        }

        public TDAPIOLELib.TSTest AddTest(TDAPIOLELib.TestSet testSet, TDAPIOLELib.Test test)
        {
            TDAPIOLELib.TSTestFactory tSTestFactory = testSet.TSTestFactory;
            TDAPIOLELib.TSTest tSTest = tSTestFactory.AddItem(test.ID);
            tSTest.Post();
            return tSTest;
        }

        /// <summary>
        /// Find tests under a test plan folder and them to testset 
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <param name="testPlanFolderPath">Test plan folder path</param>
        /// <returns></returns>
        public Boolean AddAllTestsFromTestFolder(TDAPIOLELib.TestSet testSet, String testPlanFolderPath)
        {
            TestFolders testFolders = new TestFolders(tDConnection);
            foreach (TDAPIOLELib.Test OTest in testFolders.GetTests(testPlanFolderPath))
            {
                AddTest(testSet, OTest);
            }
            return true;
        }

        /// <summary>
        /// Removes a single test from test set
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <param name="tSTest">TDAPIOLELib.TSTest Object</param>
        /// <returns></returns>
        public Boolean RemoveTest(TDAPIOLELib.TestSet testSet, TDAPIOLELib.TSTest tSTest)
        {
            TDAPIOLELib.TSTestFactory tSTestFactory = testSet.TSTestFactory;
            tSTestFactory.RemoveItem(tSTest.ID);
            testSet.Post();
            return true;
        }

        /// <summary>
        /// Removes all tests from testset
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <returns>true if successfull</returns>
        public Boolean RemoveAllTests(TDAPIOLELib.TestSet testSet)
        {
            TDAPIOLELib.List list = testSet.TSTestFactory.NewList("");
            foreach (TDAPIOLELib.TSTest tSTest in list)
            {
                testSet.TSTestFactory.RemoveItem(tSTest.ID);
                testSet.Post();
            }
            return true;
        }

        /// <summary>
        /// Filters tests from test plan using the filter text. Filter text can be colpied directly from ALM and passed to this function. 
        /// <para/> returns TDAPIOLELib.List Object. Each item from this list can be coverted to TDAPIOLELib.Test object.
        /// </summary>
        /// <param name="filterString">Copy this from ALM</param>
        /// <returns>TDAPIOLELib.List Object. Each item from this list can be coverted to TDAPIOLELib.Test object.</returns>
        public TDAPIOLELib.List Filter(String filterString)
        {
            TDAPIOLELib.TestSetFactory testSetFactory = tDConnection.TestSetFactory as TDAPIOLELib.TestSetFactory;
            TDAPIOLELib.TDFilter tDFilter = testSetFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List testList;

            try
            {
                tDFilter.Text = filterString;
                testList = testSetFactory.NewList(tDFilter.Text);

                return testList;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Get Test Plan Test Object from Test Set Test Object
        /// <para/>returns TDAPIOLELib.Test object
        /// </summary>
        /// <param name="tSTest"></param>
        /// <returns></returns>
        public TDAPIOLELib.Test GetTestObjectFromTestSetTest(TDAPIOLELib.TSTest tSTest)
        {
            return tSTest.Test;
        }

        /// <summary>
        /// Updates the field value for testset dabase fields
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <param name="fieldName">Database field name</param>
        /// <param name="newValue">new value for the test set field</param>
        /// <returns>true if successfull</returns>
        public Boolean UpdateFieldValue(TDAPIOLELib.TestSet testSet, String fieldName, String newValue, Boolean Post = true)
        {
            testSet[fieldName.ToUpper()] = newValue;
            if (Post)
            testSet.Post();
            return true;
        }

        /// <summary>
        /// Finds testset with testset ID
        /// <para/>returns TDAPIOLELib.TestSet Object
        /// </summary>
        /// <param name="id">ID of the testset</param>
        /// <returns>TDAPIOLELib.TestSet Object</returns>
        public TDAPIOLELib.TestSet GetObjectWithID(int id)
        {
            TDAPIOLELib.TestSetFactory testSetFactory = tDConnection.TestSetFactory as TDAPIOLELib.TestSetFactory;
            TDAPIOLELib.TDFilter tDFilter = testSetFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List testSetList;
            TDAPIOLELib.TestSet testSet;

            try
            {
                tDFilter["CY_CYCLE_ID"] = Convert.ToString(id);
                testSetList = testSetFactory.NewList(tDFilter.Text);

                if (testSetList != null && testSetList.Count == 1)
                {
                    testSet = testSetList[1];
                    return testSet;
                }
                else
                {
                    throw (new Exception("Unable to find test Set with ID : " + id));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Deletes a test set
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="testSet"></param>
        /// <returns>true if successfull</returns>
        public Boolean Delete(TDAPIOLELib.TestSet testSet)
        {
            try
            {
                TDAPIOLELib.TestSetFactory testSetFactory = tDConnection.TestSetFactory;
                testSetFactory.RemoveItem(testSet.ID);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Get Testlab folder path for test set
        /// <para/>returns String folder path
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <returns>String folder path</returns>
        public String GetFolderPath(TDAPIOLELib.TestSet testSet)
        {
            TDAPIOLELib.TestSetFolder testSetFolder = testSet.TestSetFolder;
            return testSetFolder.Path.ToString();
        }

        /// <summary>
        /// Run TSTest
        /// <para/>returns true if successfull
        /// </summary>
        /// <param name="runFactory">RunFactory Object for the testSetTest</param>
        /// <param name="Status">Run status</param>
        /// <returns>true if successfull</returns>
        private Boolean RunTestSetTest(TDAPIOLELib.RunFactory runFactory, String Status)
        {
            TDAPIOLELib.Run run;
            TDAPIOLELib.StepFactory stepFactory;

            run = runFactory.AddItem(DateTime.Now.ToString().Replace("/", "_").Replace(":", "_").Replace(" ", "_"));

            run.Status = Status;
            run.Post();
            run.CopyDesignSteps();
            run.Post();
            run.Refresh();

            stepFactory = run.StepFactory;

            foreach (TDAPIOLELib.Step step in stepFactory.NewList(""))
            {
                step.Status = Status;
                step.Post();
                step.Refresh();
            }
            return true;
        }

        /// <summary>
        /// Finds unattached testSets
        /// </summary>
        /// <returns>List of unattached testsets</returns>
        public TDAPIOLELib.List FindUnattachedTestSets()
        {
            TDAPIOLELib.List list = new TDAPIOLELib.List();
            TDAPIOLELib.Recordset recordset = Utilities.ExecuteQuery("SELECT CY_CYCLE_ID FROM CYCLE Where CY_FOLDER_ID < 0 and CY_CYCLE_ID > 0", tDConnection);
            recordset.First();
            for (int Counter = 1; Counter <= recordset.RecordCount; Counter++)
            {
                list.Add(GetObjectWithID(Convert.ToInt32(recordset["CY_CYCLE_ID"])));
                recordset.Next();
            }
            return list;
        }
    }
}
