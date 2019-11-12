using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Configuration;
using ALM_Wrapper;

namespace UnitTestProject1
{
    [TestClass]
    public class ALM_Wrapper_Test
    {
        static ALM_CORE aLM_CORE = new ALM_CORE();

        [TestInitialize]
        public void SetUpALM()
        {
            aLM_CORE.LoginALM(ConfigurationManager.AppSettings["ALMServer"], ConfigurationManager.AppSettings["ALMUser"], ConfigurationManager.AppSettings["ALMPassword"], ConfigurationManager.AppSettings["ALMDomain"], ConfigurationManager.AppSettings["ALMProject"]);
            Console.WriteLine("Logged in to ALM");
        }

        [TestMethod]
        public void UpdateAutomationIndicator()
        {
            ////Get all tests from ALM

            TDAPIOLELib.Recordset recordset = aLM_CORE.ExecuteQuery("Select TS_TEST_ID from test where TS_USER_TEMPLATE_06 is not null");
            Console.WriteLine(recordset.RecordCount);
            recordset.First();

            TDAPIOLELib.Test test;


            for (int Counter = 0; Counter < recordset.RecordCount; Counter++)
            {

                test = aLM_CORE.TestPlan.Test.GetObjectWithID(Convert.ToInt32(recordset[0]));

                if (test["TS_USER_TEMPLATE_06"] == "Y" && (test["TS_STATUS"] != "Ready For Automation" && test["TS_STATUS"] != "Automated"))
                {
                    Console.WriteLine(test.ID);
                    test["TS_STATUS"] = "Automated";
                    test.Post();
                }
                else if (test["TS_USER_TEMPLATE_06"] == "N" && (test["TS_STATUS"] != "Ready For Automation" && test["TS_STATUS"].ToString().ToUpper() != "Cannot be Automated".ToUpper()))
                {
                    Console.WriteLine(test.ID);
                    test["TS_STATUS"] = "Cannot be automated";
                    test.Post();
                }

                recordset.Next();
            }



        }

        [Ignore]
        [TestMethod]
        public void Verify_Requirements()
        {
            
            Dictionary<String, String> requirementDetails = new Dictionary<String, String>();
            requirementDetails.Add("RQ_REQ_NAME", "Dummy1");
            aLM_CORE.Requirements.Create(requirementDetails, "Requirements", TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES.REQ_TYPE_FOLDER);

            TDAPIOLELib.Req req = aLM_CORE.Requirements.CreateFolder("Dummy2", "Requirements\\Dummy1");
            aLM_CORE.Requirements.Delete(req);

            req = aLM_CORE.Requirements.GetReqByPath("Requirements\\Dummy1");
            Console.WriteLine(req.Path);

            aLM_CORE.Requirements.UpdateFieldValue(req, "RQ_REQ_NAME", "Dummy6", true);

            Console.WriteLine("Requirement Name after change : " + req.Name);


            List<String> Attach = new List<String>();

            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC1.txt");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC2.docx");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC3.xlsx");

            aLM_CORE.Requirements.AddAttachements(req, Attach);
            aLM_CORE.Requirements.DeleteAttachementByName(req, "DOC2.docx");
            aLM_CORE.Requirements.DeleteAllAttachments(req);

            aLM_CORE.TestPlan.TestFolders.Create("Subject", "Dummy1");

            Dictionary<String, String> TestN = new Dictionary<String, String>();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST");
            TestN.Add("TS_STATUS", "Ready");

            aLM_CORE.TestPlan.TestFolders.Create("Subject\\Dummy1", "Dummy2");



            TDAPIOLELib.Test test = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1");
            TDAPIOLELib.Test test1 = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1\\Dummy2");

            TestN.Clear();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST2");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test2 = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1\\Dummy2");



            //Create a requirement of type undefined
            TDAPIOLELib.Req req1 = aLM_CORE.Requirements.Create(requirementDetails, "Requirements\\Dummy6", TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES.REQ_TYPE_UNDEFINED);

            aLM_CORE.Requirements.AddTestToRequirementsCoverage(test, req1);
            aLM_CORE.Requirements.AddTestPlanFolderToRequirementsCoverage(aLM_CORE.TestPlan.TestFolders.GetNodeObject("Subject\\Dummy1\\Dummy2").NodeID, req1);

            requirementDetails.Clear();
            requirementDetails.Add("RQ_REQ_NAME", "Dummy2");

            TDAPIOLELib.Req req2 = aLM_CORE.Requirements.Create(requirementDetails, "Requirements\\Dummy6", TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES.REQ_TYPE_UNDEFINED);

            TDAPIOLELib.Trace trace1 = aLM_CORE.Requirements.AddRequirementTraceability_TraceFrom(req1, req2);
            TDAPIOLELib.Trace trace2 = aLM_CORE.Requirements.AddRequirementTraceability_TraceTo(req1, req2);

            foreach (TDAPIOLELib.Trace oTrace in aLM_CORE.Requirements.GetRequirementTraceability_TraceFrom(req1))
            {
                Console.Write("Requirement Tracebility Trace from requirements : " + oTrace.FromReq.Name);
            }

            foreach (TDAPIOLELib.Trace oTrace in aLM_CORE.Requirements.GetRequirementTraceability_TraceTo(req1))
            {
                Console.Write("Requirement Tracebility Trace from requirements : " + oTrace.ToReq.Name);
            }

            TDAPIOLELib.Recordset recordset = aLM_CORE.ExecuteQuery("Select REL_ID From Releases where rownum < 10");
            recordset.First();

            TDAPIOLELib.List releaseIDList = new TDAPIOLELib.List();
            for (int Counter = 0; Counter < recordset.RecordCount; Counter++)
            {
                Console.WriteLine("Release ID : " + recordset[0]);
                releaseIDList.Add(recordset[0]);
                recordset.Next();
            }

            aLM_CORE.Requirements.SetTargetReleases(req1, releaseIDList);

            foreach (TDAPIOLELib.Release release in aLM_CORE.Requirements.GetTargetReleases(req1))
            {
                Console.WriteLine("Target Release ID : " + release.ID);
            }

            aLM_CORE.Requirements.DeleteTargetReleases(req1);

            recordset = aLM_CORE.ExecuteQuery("Select RCYC_ID From RELEASE_CYCLES where rownum < 10");
            recordset.First();

            TDAPIOLELib.List cycleIDList = new TDAPIOLELib.List();
            for (int Counter = 0; Counter < recordset.RecordCount; Counter++)
            {
                Console.WriteLine("Cycle ID : " + recordset[0]);
                cycleIDList.Add(recordset[0]);
                recordset.Next();
            }

            aLM_CORE.Requirements.SetTargetCycles(req1, cycleIDList);

            foreach (TDAPIOLELib.Cycle cycle in aLM_CORE.Requirements.GetTargetCycles(req1))
            {
                Console.WriteLine("Target Cycle ID : " + cycle.ID);
            }

            aLM_CORE.Requirements.DeleteTargetCycles(req1);

            aLM_CORE.Requirements.DeleteRequirementTraceability_TraceFrom(req1, trace1);
            aLM_CORE.Requirements.DeleteRequirementTraceability_TraceTo(req1, trace2);

            foreach (TDAPIOLELib.Req oReq in aLM_CORE.Requirements.GetChildRequirements(aLM_CORE.Requirements.GetReqByPath("Requirements\\Dummy6")))
            {
                Console.WriteLine("Child requirement Found : " + oReq.Name);
            }

            Console.WriteLine("Parent Requirement Name : " + aLM_CORE.Requirements.GetParentRequirement(aLM_CORE.Requirements.GetReqByPath("Requirements\\Dummy6\\Dummy1")).Name);

            aLM_CORE.Requirements.Delete(req1);
            aLM_CORE.Requirements.Delete(req2);

            aLM_CORE.Requirements.Delete(req);


            aLM_CORE.TestPlan.Test.Delete(test);
            aLM_CORE.TestPlan.Test.Delete(test1);
            aLM_CORE.TestPlan.Test.Delete(test2);

            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1");
        }

        [Ignore]
        [TestMethod]
        public void Verify_Cycles()
        {
            Dictionary<String, String> releaseDetails = new Dictionary<string, string>();
            TDAPIOLELib.ReleaseFolder releaseFolder = aLM_CORE.Releases.ReleaseFolders.Create("Releases", "Dummy1");
            releaseDetails.Add("REL_NAME", "Release1");
            releaseDetails.Add("REL_START_DATE", DateTime.Now.ToShortDateString());
            releaseDetails.Add("REL_END_DATE", DateTime.Now.AddDays(10).ToShortDateString());
            TDAPIOLELib.Release release = aLM_CORE.Releases.Release.Create(releaseDetails, aLM_CORE.Releases.ReleaseFolders.GetPath(releaseFolder));

            TDAPIOLELib.Cycle cycle = aLM_CORE.Releases.Cycle.Create("Cycle1", release.StartDate, release.EndDate, release);

            List<String> Attach = new List<String>();

            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC1.txt");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC2.docx");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC3.xlsx");

            aLM_CORE.Releases.Cycle.AddAttachment(cycle, Attach);

            aLM_CORE.Releases.Cycle.DownloadAttachments(cycle, "C:\\Temp\\ALMDOWNLOAD");

            aLM_CORE.Releases.Cycle.DeleteAttachmentByName(cycle, "DOC1.txt");
            aLM_CORE.Releases.Cycle.DeleteAllAttachments(cycle);

            TDAPIOLELib.Recordset ORec = aLM_CORE.Releases.Cycle.GetAllDetails(cycle);
            for (int i = 0; i < ORec.RecordCount; i++)
            {
                for (int j = 0; j < ORec.ColCount; j++)
                {
                    Console.WriteLine(ORec.ColName[j] + "--" + ORec[j]);
                }
                ORec.Next();
            }

            release = aLM_CORE.Releases.Cycle.GetRelease(cycle);
            Console.WriteLine(release.Name);

            aLM_CORE.Releases.Cycle.Delete(cycle);
            aLM_CORE.Releases.Release.Delete(release);
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases", "Dummy1");

        }

        [Ignore]
        [TestMethod]
        public void Verify_Releases()
        {
            Dictionary<String, String> releaseDetails = new Dictionary<string, string>();
            TDAPIOLELib.ReleaseFolder releaseFolder = aLM_CORE.Releases.ReleaseFolders.Create("Releases", "Dummy1");
            releaseDetails.Add("REL_NAME", "Release1");
            releaseDetails.Add("REL_START_DATE", DateTime.Now.ToShortDateString());
            releaseDetails.Add("REL_END_DATE", DateTime.Now.AddDays(10).ToShortDateString());
            TDAPIOLELib.Release release = aLM_CORE.Releases.Release.Create(releaseDetails, aLM_CORE.Releases.ReleaseFolders.GetPath(releaseFolder));
            List<String> Attach = new List<String>();

            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC1.txt");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC2.docx");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC3.xlsx");
            aLM_CORE.Releases.Release.AddAttachment(release, Attach);

            aLM_CORE.Releases.Release.DownloadAttachments(release, "C:\\Temp\\ALMDOWNLOAD");
            aLM_CORE.Releases.Release.DeleteAttachmentByName(release, "DOC3.xlsx");
            aLM_CORE.Releases.Release.DeleteAllAttachments(release);
            Console.WriteLine("Path of release : " + aLM_CORE.Releases.Release.GetPath(release));

            foreach (TDAPIOLELib.Cycle cycle in aLM_CORE.Releases.Release.GetCycles(aLM_CORE.Releases.ReleaseFolders.GetReleaseByName(aLM_CORE.Releases.ReleaseFolders.GetNodeObject("Releases\\Dummy"), "Test123")))
            {
                Console.WriteLine("Cycle Name : " + cycle.Name);
            }

            aLM_CORE.Releases.Release.Delete(release);
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases", "Dummy1");

        }

        [Ignore]
        [TestMethod]
        public void verify_ReleaseFolders()
        {
            aLM_CORE.Releases.ReleaseFolders.Create("Releases", "Dummy1");
            aLM_CORE.Releases.ReleaseFolders.Create("Releases\\Dummy1", "Dummy2");
            TDAPIOLELib.ReleaseFolder releaseFolder = aLM_CORE.Releases.ReleaseFolders.CreateNewFolderPath("Releases\\Dummy1\\Dummy2\\Dummy3\\Dummy4\\Dummy5");
            TDAPIOLELib.ReleaseFolder releaseFolder1 = aLM_CORE.Releases.ReleaseFolders.CreateNewFolderPath("Releases\\Dummy1\\Dummy2\\Dummy3\\Dummy4\\Dummy6");

            Console.WriteLine(aLM_CORE.Releases.ReleaseFolders.GetPath(releaseFolder));

            List<String> Attach = new List<String>();

            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC1.txt");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC2.docx");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC3.xlsx");

            aLM_CORE.Releases.ReleaseFolders.AddAttachment(releaseFolder, Attach);
            aLM_CORE.Releases.ReleaseFolders.DeleteAttachmentByName(releaseFolder, "DOC1.txt");
            aLM_CORE.Releases.ReleaseFolders.DeleteAllAttachments(releaseFolder);

            foreach (String folderNames in aLM_CORE.Releases.ReleaseFolders.GetChildFolderNames("Releases"))
            {
                Console.WriteLine(folderNames);
            }

            foreach (TDAPIOLELib.ReleaseFolder releasef in aLM_CORE.Releases.ReleaseFolders.GetChildFolders(aLM_CORE.Releases.ReleaseFolders.GetNodeObject("Releases\\Dummy1\\Dummy2\\Dummy3\\Dummy4")))
            {
                Console.WriteLine(releasef.Name);
            }

            releaseFolder1 = aLM_CORE.Releases.ReleaseFolders.GetChildFolderWithName(aLM_CORE.Releases.ReleaseFolders.GetNodeObject("Releases\\Dummy1\\Dummy2\\Dummy3\\Dummy4"), "Dummy6");

            Console.WriteLine(releaseFolder1.Modified);

            Console.WriteLine("Path of the releaseFolder1 : " + aLM_CORE.Releases.ReleaseFolders.GetPath(releaseFolder1));
            aLM_CORE.Releases.ReleaseFolders.Rename(releaseFolder1, "Dummy Renamed");

            Console.WriteLine("Dummy6 after rename : " + releaseFolder1.Name);

            foreach (TDAPIOLELib.Release release in aLM_CORE.Releases.ReleaseFolders.GetReleases(aLM_CORE.Releases.ReleaseFolders.GetNodeObject("Releases\\Dummy")))
            {
                Console.WriteLine("Release Name : " + release.Name);
            }

            aLM_CORE.Releases.ReleaseFolders.Delete("Releases\\Dummy1\\Dummy2\\Dummy3\\Dummy4", "Dummy5");
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases\\Dummy1\\Dummy2\\Dummy3\\Dummy4", "\"Dummy Renamed\"");
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases\\Dummy1\\Dummy2\\Dummy3", "Dummy4");
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases\\Dummy1\\Dummy2", "Dummy3");
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases\\Dummy1", "Dummy2");
            aLM_CORE.Releases.ReleaseFolders.Delete("Releases", "Dummy1");
        }

        [Ignore]
        [TestMethod]
        public void Verify_Defects()
        {
            Dictionary<String, String> defectDetails = new Dictionary<string, string>();

            defectDetails.Add("BG_SUMMARY", "Defect created using Automation");
            defectDetails.Add("BG_USER_TEMPLATE_01", "TEST");
            defectDetails.Add("BG_DETECTED_IN_RCYC", "1014");
            defectDetails.Add("BG_DETECTION_DATE", DateTime.Now.ToShortDateString());
            defectDetails.Add("BG_SEVERITY", "Sev-3");
            defectDetails.Add("BG_DETECTED_BY", "Sumeet.Kushwah");

            TDAPIOLELib.Bug bug = aLM_CORE.Defect.Create(defectDetails);
            Console.WriteLine("Total Defects in Project : " + aLM_CORE.Defect.CountAll());

            TDAPIOLELib.Recordset ORec = aLM_CORE.Defect.GetAllDetails(bug);

            Console.WriteLine("Writing all Database field names and values...");

            for (int i = 0; i < ORec.RecordCount; i++)
            {
                for (int j = 0; j < ORec.ColCount; j++)
                {
                    Console.WriteLine(ORec.ColName[j] + "--" + ORec[j]);
                }
                ORec.Next();
            }

            //Create a test and Link defect to it
            // Create a test folder
            aLM_CORE.TestPlan.TestFolders.Create("Subject", "Dummy1");

            //create a test here 
            Dictionary<String, String> TestN = new Dictionary<String, String>();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1");
            Console.WriteLine("Test Created Under Folder Subject\\Dummy1" + test.Name);
            aLM_CORE.Defect.LinkDefectToTest(test, bug, "Linking defect to test");

            //Create a test set and Link defects to it
            aLM_CORE.TestLab.TestLabFolders.Create("Root", "Dummy1");
            Dictionary<String, String> testSetDetails = new Dictionary<string, string>();
            testSetDetails.Add("CY_CYCLE", "Dummy Test Set");
            TDAPIOLELib.TestSet testSet = aLM_CORE.TestLab.TestSet.Create(testSetDetails, "Root\\Dummy1");

            aLM_CORE.Defect.LinkDefectToTestSet(testSet, bug, "Test Set to Bug Linked");

            TDAPIOLELib.TSTest tSTest = aLM_CORE.TestLab.TestSet.AddTest(testSet, test);
            aLM_CORE.Defect.LinkDefectToTestSetTest(tSTest, bug, "Test Set Test to Bug Linked");


            TDAPIOLELib.List list = aLM_CORE.Defect.GetLinkedDefectsToTest(test);

            foreach (TDAPIOLELib.Bug bug1 in list)
            {
                Console.WriteLine("Defect Attached to test is : " + bug1.Summary);
            }

            list = aLM_CORE.Defect.GetLinkedDefectsToTestSet(testSet);

            foreach (TDAPIOLELib.Bug bug1 in list)
            {
                Console.WriteLine("Defect Attached to testset is : " + bug1.Summary);
            }

            list = aLM_CORE.Defect.GetLinkedDefectsToTestSetTest(tSTest);

            foreach (TDAPIOLELib.Bug bug1 in list)
            {
                Console.WriteLine("Defect Attached to testset test is : " + bug1.Summary);
            }

            List<String> Attach = new List<String>();

            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC1.txt");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC2.docx");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC3.xlsx");

            aLM_CORE.Defect.AddAttachment(bug, Attach);

            aLM_CORE.Defect.DownloadAttachments(bug, "C:\\Temp\\ALMDOWNLOAD");

            aLM_CORE.Defect.DeleteAttachmentByName(bug, "DOC2.docx");
            aLM_CORE.Defect.DeleteAllAttachments(bug);

            try
            {
                aLM_CORE.Defect.UpdateFieldValue(bug, "BG_STATUS", "Closed");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
            }

            bug = aLM_CORE.Defect.GetObjectWithID(Convert.ToInt32(bug.ID));

            aLM_CORE.TestPlan.Test.Delete(test);
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1");
            aLM_CORE.TestLab.TestLabFolders.Delete("Root", "Dummy1");
            aLM_CORE.TestLab.TestSet.Delete(testSet);

            aLM_CORE.Defect.Delete(bug);

            Console.WriteLine("Done");
        }

        [Ignore]
        [TestMethod]
        public void verify_TestSets()
        {
            aLM_CORE.TestLab.TestLabFolders.Create("Root", "Dummy1");

            Dictionary<String, String> testSetDetails = new Dictionary<string, string>();

            testSetDetails.Add("CY_CYCLE", "Dummy Test Set");
            TDAPIOLELib.TestSet testSet = aLM_CORE.TestLab.TestSet.Create(testSetDetails, "Root\\Dummy1");

            Console.WriteLine("Test Set created : " + testSet.Name);

            //Create a test folder and a test inside it. We will add this to the testSet we created above
            aLM_CORE.TestPlan.TestFolders.Create("Subject", "Dummy1");

            Console.WriteLine("Test Folder Created : Subject\\Dummy1");

            //create a test here 
            Dictionary<String, String> TestN = new Dictionary<String, String>();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1");

            Console.WriteLine("Test Created : " + test.Name);

            //Addign test to testset
            aLM_CORE.TestLab.TestSet.AddTest(testSet, test);

            Console.WriteLine("Test added to test set");

            //Create a new folder in test plan and add three new tests to it
            //Create a test folder and a test inside it. We will add this to the testSet we created above
            aLM_CORE.TestPlan.TestFolders.Create("Subject", "Dummy2");
            Console.WriteLine("Test Folder Created : Subject\\Dummy2");
            TestN.Clear();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST 1");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test1 = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy2");

            TestN.Clear();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST 2");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test2 = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy2");

            TestN.Clear();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST 3");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test3 = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy2");

            Console.WriteLine("Created three new tests and added it to the folder Subject\\Dummy2");

            aLM_CORE.TestLab.TestSet.AddAllTestsFromTestFolder(testSet, "Subject\\Dummy2");

            Console.WriteLine("Three tests from Subject\\Dummy2 are added to testSet : " + testSet.Name);

            List<String> Attach = new List<String>();

            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC1.txt");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC2.docx");
            Attach.Add("C:\\Temp\\ALMUPLOAD\\DOC3.xlsx");

            aLM_CORE.TestLab.TestSet.AddAttachments(testSet, Attach);

            Console.WriteLine("Added Attachments to testset");

            aLM_CORE.TestLab.TestSet.PassAllTests(testSet);

            Console.WriteLine("Executed all tests from test set and updated status as passed");

            aLM_CORE.TestLab.TestSet.FailAllTests(testSet);

            Console.WriteLine("Executed all tests from test set and updated status as Failed");

            TDAPIOLELib.List TSTestList = aLM_CORE.TestLab.TestSet.GetAllTestInstance(testSet);

            aLM_CORE.TestLab.TestSet.ExecuteSingleTest(TSTestList[1], "Not Completed");

            Console.WriteLine("Executed First test from test set and marked it as Not Completed");

            Dictionary<int, String> executionList = new Dictionary<int, String>();

            //Mark all tests as No Run
            //Create the dictionary object for the test set
            foreach (TDAPIOLELib.TSTest tSTest1 in TSTestList)
            {
                executionList.Add(Convert.ToInt32(tSTest1.ID), "No Run");
            }

            aLM_CORE.TestLab.TestSet.ExecuteTests(executionList);
            Console.WriteLine("Changed status of all tests to No Run");

            testSet = aLM_CORE.TestLab.TestSet.GetObjectWithID(Convert.ToInt32(testSet.ID));
            Console.WriteLine("Got test set with ID");

            Console.WriteLine("TestLab Folder Path for the test set is : " + aLM_CORE.TestLab.TestSet.GetFolderPath(testSet));
            //Console.WriteLine("Test Plan Folder path for the First test in the test set is : " + aLM_CORE.TestPlan.Test.GetPath((aLM_CORE.TestLab.TestSet.GetTestObjectFromTestSetTest((TSTestList[0] as TDAPIOLELib.TSTest).ID) as TDAPIOLELib.Test)));

            TDAPIOLELib.TSTest tSTest = aLM_CORE.TestLab.TestSet.GetTSTestWithID(Convert.ToInt32(TSTestList[2].ID));
            Console.WriteLine("Found Second Test Object from test set : " + tSTest.TestName);

            TDAPIOLELib.Run run = aLM_CORE.TestLab.TestSet.GetLastRunDetails(TSTestList[1]);
            Console.WriteLine("Got last run details for test set test : " + run.Name + " : " + run.ID);

            aLM_CORE.TestLab.TestSet.RemoveTest(testSet, TSTestList[1]);
            Console.WriteLine("Removed first test from testset");

            aLM_CORE.TestLab.TestSet.RemoveAllTests(testSet);
            Console.WriteLine("Removed all tests from testset");

            aLM_CORE.TestLab.TestSet.DownloadAttachements(testSet, "C:\\Temp\\ALMDOWNLOAD");
            Console.WriteLine("Downloaded attachements");

            aLM_CORE.TestLab.TestSet.DeleteAttachmentByName(testSet, "DOC1.txt");
            Console.WriteLine("Deleted attachment by name : DOC1.txt");

            aLM_CORE.TestLab.TestSet.DeleteAllAttachments(testSet);
            Console.WriteLine("Deleted all attachments");

            Console.WriteLine("Printing testsets under unattached...");

            TDAPIOLELib.List unattachedTestSetsList = aLM_CORE.TestLab.TestSet.FindUnattachedTestSets();
            foreach (TDAPIOLELib.TestSet unattachedTestSets in unattachedTestSetsList)
            {
                Console.WriteLine(unattachedTestSets.Name);
            }

            Console.WriteLine("Done");

            aLM_CORE.TestLab.TestSet.UpdateFieldValue(testSet, "CY_STATUS", "Closed");

            //CleanUp
            aLM_CORE.TestLab.TestSet.Delete(testSet);
            aLM_CORE.TestLab.TestLabFolders.Delete("Root", "Dummy1");
            aLM_CORE.TestPlan.Test.Delete(test);
            aLM_CORE.TestPlan.Test.Delete(test1);
            aLM_CORE.TestPlan.Test.Delete(test2);
            aLM_CORE.TestPlan.Test.Delete(test3);
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1");
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy2");
        }

        [Ignore]
        [TestMethod]
        public void Verify_TestSetFolders()
        {
            aLM_CORE.TestLab.TestLabFolders.Create("Root", "Dummy1");
            aLM_CORE.TestLab.TestLabFolders.CreateNewFolderPath("Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4");

            Console.WriteLine("Created TestLab Folders");

            //Find Folders under a testlab folder
            List<String> list = aLM_CORE.TestLab.TestLabFolders.FindChildFolders("Root\\Dummy1");
            foreach (String folderName in list)
            {
                Console.WriteLine("Folder Under Root\\Dummy\\ : " + folderName);
            }

            //Create two testsets
            Dictionary<String, String> testSet_New = new Dictionary<String, String>();
            testSet_New.Add("CY_CYCLE", "Dummy test set name");

            aLM_CORE.TestLab.TestSet.Create(testSet_New, "Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4");

            testSet_New.Clear();
            testSet_New.Add("CY_CYCLE", "Dummy test set name 2");
            aLM_CORE.TestLab.TestSet.Create(testSet_New, "Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4");

            TDAPIOLELib.List tdList = aLM_CORE.TestLab.TestLabFolders.FindTestSets("Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4");
            foreach (TDAPIOLELib.TestSet testSet in tdList)
            {
                Console.WriteLine("TestSets under Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4 : " + testSet.Name);
            }

            TDAPIOLELib.SysTreeNode sysTreeNode = aLM_CORE.TestLab.TestLabFolders.GetNodeObject("Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4");
            Console.WriteLine("sysTreeNode Object : " + sysTreeNode.Path);

            TDAPIOLELib.TestSetFolder testSetFolder = aLM_CORE.TestLab.TestLabFolders.GetTestSetFolder("Root\\Dummy1\\Dummy2\\Dummy3\\Dummy4");
            Console.WriteLine("Parent Folder name : " + testSetFolder.Father.Name.ToString());

            aLM_CORE.TestLab.TestLabFolders.Rename("Root\\Dummy1\\Dummy2\\Dummy3", "Dummy4", "Renamed Dummy4");

            Console.WriteLine("Renamed a Folder");

            aLM_CORE.TestLab.TestLabFolders.Delete("Root\\Dummy1\\Dummy2\\Dummy3", "Renamed Dummy4");
            aLM_CORE.TestLab.TestLabFolders.Delete("Root\\Dummy1\\Dummy2", "Dummy3");
            aLM_CORE.TestLab.TestLabFolders.Delete("Root\\Dummy1", "Dummy2");
            aLM_CORE.TestLab.TestLabFolders.Delete("Root", "Dummy1");

            Console.WriteLine("Deleted all folders");

            Console.WriteLine("Done");

        }

        [Ignore]
        [TestMethod]
        public void Verify_TestScripts()
        {
            //Create a test folder
            aLM_CORE.TestPlan.TestFolders.Create("Subject", "Dummy1");

            //create a test here 
            Dictionary<String, String> TestN = new Dictionary<String, String>();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST");
            TestN.Add("TS_STATUS", "Ready");
            TDAPIOLELib.Test test = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1");

            Console.WriteLine("Test Created Under Folder Subject\\Dummy1" + test.Name);

            //Add Design steps
            //Expected Values
            List<String> Expected = new List<String>();
            Expected.Add("Step dummy 1");
            Expected.Add("Step dummy 2");

            //Description values
            List<String> Desc = new List<String>();
            Desc.Add("Step dummy 1 Desc");
            Desc.Add("Step dummy 2 Desc");

            //Attachments for the design steps
            List<String> Attach = new List<String>();
            Attach.Add(@"C:\Temp\ALMUPLOAD\DOC2.docx");
            Attach.Add(@"C:\Temp\ALMUPLOAD\DOC3.xlsx");

            aLM_CORE.TestPlan.Test.AddDesignSteps(test, Desc, Expected, Attach);

            Console.WriteLine("Added Bulk design steps with attachments to test");

            //clear Attach
            Attach.Clear();

            //Upload two attachments to design step
            Attach.Add(@"C:\Temp\ALMUPLOAD\DOC1.txt");
            Attach.Add(@"C:\Temp\ALMUPLOAD\DOC3.xlsx");

            TDAPIOLELib.DesignStep designStep = aLM_CORE.TestPlan.Test.AddSingleDeignStep(test, "Single Step", "Nothing is expected");
            aLM_CORE.TestPlan.Test.AddAttachmentToDesignStep(designStep, Attach);

            Console.WriteLine("Added Single design step with two attachments to test");

            aLM_CORE.TestPlan.Test.AddSingleParameter(test, "BrowserName", "Chrome, IE, Firefox", "Chrome");

            Console.WriteLine("Added Single parameter");

            //Add test parameters
            List<String> ParamName = new List<String>();
            ParamName.Add("Browser");
            ParamName.Add("URL");

            List<String> ParamDescription = new List<String>();
            ParamDescription.Add("Browser for executions");
            ParamDescription.Add("URL to be used");

            List<String> ParamDefaultVal = new List<String>();
            ParamDefaultVal.Add("IE");
            ParamDefaultVal.Add("Http://www.google.com");

            aLM_CORE.TestPlan.Test.AddParameters(test, ParamName, ParamDescription, ParamDefaultVal);

            Console.WriteLine("Added Bulk Parameters to test");

            Console.WriteLine("Number of tests in test plan : " + aLM_CORE.TestPlan.Test.CountAllTests().ToString());

            Console.WriteLine("Number of tests under Subject\\Dummy1 : " + aLM_CORE.TestPlan.Test.CountTestUnderFolder("Subject\\Dummy1"));

            //Search test with ID
            test = aLM_CORE.TestPlan.Test.GetObjectWithID(Convert.ToInt32(test.ID));

            TDAPIOLELib.List tdList = aLM_CORE.TestPlan.Test.GetDesignStepList(test);

            Console.WriteLine("Printing design steps for test..");

            foreach (TDAPIOLELib.DesignStep ODes in tdList)
            {
                Console.WriteLine(ODes.StepDescription + " : " + ODes.StepExpectedResult);
            }

            //Add Attachments to test
            List<String> AttachmentsPath = new List<string>();
            AttachmentsPath.Add(@"C:\Temp\ALMUPLOAD\DOC1.txt");
            AttachmentsPath.Add(@"C:\Temp\ALMUPLOAD\DOC3.xlsx");
            AttachmentsPath.Add(@"C:\Temp\ALMUPLOAD\DOC2.docx");

            aLM_CORE.TestPlan.Test.AddAttachment(test, AttachmentsPath);

            Console.WriteLine("Added Attachments to test");

            aLM_CORE.TestPlan.Test.DownloadAttachments(test, "C:\\Temp\\ALMDOWNLOAD");

            Console.WriteLine("Downloaded test attachments to C:\\Temp\\ALMDOWNLOAD");

            aLM_CORE.TestPlan.Test.DeleteAttachmentByName(test, "DOC3.xlsx");

            Console.WriteLine("Deleted Attachment by name");

            aLM_CORE.TestPlan.Test.DeleteAllAttachments(test);

            Console.WriteLine("Deleted all attachment");

            Console.WriteLine("Path of test case : " + aLM_CORE.TestPlan.Test.GetPath(test));

            aLM_CORE.TestPlan.Test.MarkAsTemplate(test);

            Console.WriteLine("Marked test as template test");

            aLM_CORE.TestPlan.Test.UpdateFieldValue(test, "TS_STATUS", "Design");

            Console.WriteLine("Updated Test Field value");

            aLM_CORE.TestPlan.Test.Delete(test);

            Console.WriteLine("Deleted test");

            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1");

            Console.WriteLine("Deleted Folder Dummy1");

            Console.WriteLine("Reading tests from unattached...");

            TDAPIOLELib.List Tdlist = aLM_CORE.TestPlan.Test.FindUnattachedTests();
            foreach (TDAPIOLELib.Test Otest in Tdlist)
            {
                Console.WriteLine("Tests from unattached : " + Otest.Name);// + ", Path : " + aLM_CORE.TestPlan.Test.GetPath(test));
            }

            //Delete All Unattached tests
            foreach (TDAPIOLELib.Test Otest in Tdlist)
            {
                aLM_CORE.TestPlan.Test.Delete(Otest);
            }

            Console.WriteLine("Deleted all tests from Unattached");
            Console.WriteLine("Done");
        }

        [Ignore]
        [TestMethod]
        public void Verify_TestFolder()
        {
            /////Test TestFolders.Cs
            TDAPIOLELib.TestFolder parent = aLM_CORE.TestPlan.TestFolders.Create("Subject", "Dummy1");
            aLM_CORE.TestPlan.TestFolders.CreateNewFolderPath("Subject\\Dummy1\\Dummy2\\Dummy3");
            aLM_CORE.TestPlan.TestFolders.CreateNewFolderPath("Subject\\Dummy1\\Dummy4\\Dummy5");
            aLM_CORE.TestPlan.TestFolders.CreateNewFolderPath("Subject\\Dummy1\\Dummy6\\Dummy7");

            Console.WriteLine("Folders Created Successfully in ALM");

            List<String> list = aLM_CORE.TestPlan.TestFolders.GetChildFolderNames("Subject\\Dummy1");

            foreach (String folderName in list)
            {
                Console.WriteLine("Folder Found under Subject\\Dummy1 : " + folderName);
            }

            list = aLM_CORE.TestPlan.TestFolders.GetChildFolderNames(parent);
            foreach (String folderName in list)
            {
                Console.WriteLine("Folder Found under Subject\\Dummy1 : " + folderName);
            }

            TDAPIOLELib.SysTreeNode sysTreeNode = aLM_CORE.TestPlan.TestFolders.GetNodeObject("Subject\\Dummy1\\Dummy6");
            Console.WriteLine("Count of folders under this folder is : " + sysTreeNode.Count);

            parent = aLM_CORE.TestPlan.TestFolders.GetFolderObject("Subject\\Dummy1\\Dummy6");
            Console.WriteLine("Count of folders under this folder is : " + parent.TestFolderFactory.NewList("").Count);

            //Create first test under folder
            Dictionary<String, String> TestN = new Dictionary<String, String>();
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST");
            TestN.Add("TS_STATUS", "Ready");
            aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1\\Dummy6");

            TestN.Clear();

            //Create second test inder folder
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST1");
            TestN.Add("TS_STATUS", "Design");
            aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1\\Dummy6");

            //Get the List of folders under test folder
            TDAPIOLELib.List Tdlist = aLM_CORE.TestPlan.TestFolders.GetTests("Subject\\Dummy1\\Dummy6");

            foreach (TDAPIOLELib.Test test in Tdlist)
            {
                Console.WriteLine("Test Found under Subject\\Dummy1\\Dummy6 : " + test.Name);
            }

            //Create test under Dummy 7, which is under dummy 6
            TestN.Clear();

            //Create second test inder folder
            TestN.Add("TS_NAME", "THIS IS DUMMUY TEST2");
            TestN.Add("TS_STATUS", "Design");
            TDAPIOLELib.Test test12 = aLM_CORE.TestPlan.Test.Create(TestN, "Subject\\Dummy1\\Dummy6\\Dummy7");

            Tdlist = aLM_CORE.TestPlan.TestFolders.GetTestsRecurcively("Subject\\Dummy1\\Dummy6");
            foreach (TDAPIOLELib.Test test in Tdlist)
            {
                Console.WriteLine("Test Found Name : " + test.Name + ", Path : " + aLM_CORE.TestPlan.Test.GetPath(test));
            }

            //Rename
            aLM_CORE.TestPlan.TestFolders.Rename("Subject\\Dummy1\\Dummy6\\Dummy7", "Renamed7");

            //Now read the tests again
            Tdlist = aLM_CORE.TestPlan.TestFolders.GetTestsRecurcively("Subject\\Dummy1\\Dummy6");
            foreach (TDAPIOLELib.Test test in Tdlist)
            {
                Console.WriteLine("After Renaming Folder Test Found Name : " + test.Name + ", Path : " + aLM_CORE.TestPlan.Test.GetPath(test));
            }

            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1\\Dummy2\\Dummy3");
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1\\Dummy4\\Dummy5");
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1\\Dummy6\\Renamed7", true);

            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1\\Dummy2");
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1\\Dummy4");
            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1\\Dummy6");

            aLM_CORE.TestPlan.TestFolders.Delete("Subject\\Dummy1");

            aLM_CORE.TestPlan.Test.Delete(test12);

            Console.WriteLine("Deleted all folders");


            Console.WriteLine("Done");
        }

        [Ignore]
        [TestMethod]
        public void Verify_TestResources()
        {
            TDAPIOLELib.QCResourceFolder qCResourceFolder = aLM_CORE.TestResources.GetFolderFromPath("Resources\\TESTING");
            qCResourceFolder = aLM_CORE.TestResources.CreateFolder(qCResourceFolder, "NEW RESOURCE");
            aLM_CORE.TestResources.DeleteFolder(qCResourceFolder);

            TDAPIOLELib.QCResourceFolder qCResourceFolder1 = aLM_CORE.TestResources.GetFolderFromPath("Resources");
            qCResourceFolder = aLM_CORE.TestResources.CreateFolder(qCResourceFolder1, "NEW RESOURCE");
            TDAPIOLELib.QCResourceFolder qCResourceFolder2 = aLM_CORE.TestResources.CreateFolder(qCResourceFolder1, "NEW RESOURCE1");
            foreach (TDAPIOLELib.QCResourceFolder OFolder in aLM_CORE.TestResources.GetChildFolderList(qCResourceFolder1))
            {
                Console.WriteLine(OFolder.Name);
            }

            //Upload File
            TDAPIOLELib.QCResource qCResource = aLM_CORE.TestResources.CreateResource(qCResourceFolder, "TESTRES", "C:\\Temp\\ALMUPLOAD\\DOC2.docx", ALM_Wrapper.TestResources.ResourceUploadType.File, ALM_Wrapper.TestResources.ResourceType.TestResource);

            //Upload Folder
            qCResource = aLM_CORE.TestResources.CreateResource(qCResourceFolder, "TESTRES1", "C:\\Temp\\ALMUPLOAD", ALM_Wrapper.TestResources.ResourceUploadType.Folder, ALM_Wrapper.TestResources.ResourceType.TestResource);

            qCResource = aLM_CORE.TestResources.CreateResource(qCResourceFolder, "TESTRES2", "C:\\Temp\\ALMUPLOAD\\DOC3.xlsx", ALM_Wrapper.TestResources.ResourceUploadType.File, ALM_Wrapper.TestResources.ResourceType.DataTable);

            aLM_CORE.TestResources.DeleteFolder(qCResourceFolder);
            aLM_CORE.TestResources.DeleteFolder(qCResourceFolder2);

            aLM_CORE.TestResources.CreateFolderPath("Resources\\Test\\Test1\\Test2");
            aLM_CORE.TestResources.DeleteFolder(aLM_CORE.TestResources.GetFolderFromPath("Resources\\Test"));
        }

        
        [TestMethod]
        public void Test_AnalysisAndDashboardScripts()
        {
            ////Analysis Scripts
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = aLM_CORE.Analysis.FindPrivateFolder();
            Console.WriteLine("ID of private folder : " + analysisItemFolder.ID);

            

            analysisItemFolder = aLM_CORE.Analysis.FindPublicFolder();
            Console.WriteLine("ID of Public folder : " + analysisItemFolder.ID);

            analysisItemFolder = aLM_CORE.Analysis.CreateFolder("Private", "TestFolder1");
            analysisItemFolder = aLM_CORE.Analysis.CreateFolderPath("Private\\TestFolder1\\Dummy2\\Dummy3");

            TDAPIOLELib.AnalysisItem analysisItem = aLM_CORE.Analysis.CreateExcelReport(analysisItemFolder, "FindAllBugs", "Select * from Bug");
            
            analysisItem = aLM_CORE.Analysis.CreateDefectSummaryGraph(analysisItemFolder, "DefectFirstGraph", "BG_STATUS", ALM_Wrapper.Analysis.DefectSummaryGraphSumOF.ActualFixTime, "BG_DETECTED_IN_REL", "");

            Console.WriteLine(analysisItem.LayoutData.ToString());

            TDAPIOLELib.AnalysisItemFileFactory analysisItemFileFactory = analysisItem.AnalysisItemFileFactory;

            foreach (TDAPIOLELib.AnalysisItemFile aif in analysisItemFileFactory.NewList(""))
            {
                aif.SetFilePath("C:\\Temp");
                aif.Download();
            }

            //TDAPIOLELib.Gra


            TDAPIOLELib.AnalysisItem analysisItem1 = aLM_CORE.Analysis.CreateDefectAgeGraph(analysisItemFolder, "Defect First Age Graph", "BG_RESPONSIBLE", ALM_Wrapper.Analysis.DefectSummaryGraphSumOF.None, ALM_Wrapper.Analysis.DefectAgeGrouping.NoGrouping, "");

            aLM_CORE.Analysis.RenameFolder(analysisItemFolder, "TestFolder2");
            Console.WriteLine("New Folder Name : " + analysisItemFolder.Name);

            ///Dashboard Scripts
            TDAPIOLELib.DashboardFolder dashboardFolderParent = aLM_CORE.Dashboard.CreateFolder(aLM_CORE.Dashboard.FindPrivateFolder(), "TestFolder1");

            TDAPIOLELib.DashboardFolder dashboardFolder = aLM_CORE.Dashboard.CreateFolderPath("Private\\TestFolder1\\Dummy1\\Dummy2");
            TDAPIOLELib.DashboardPage dashboardPage = aLM_CORE.Dashboard.CreatePage(dashboardFolder, "TESTPAGE1");

            TDAPIOLELib.DashboardPageItem dashboardPage1 = aLM_CORE.Dashboard.AddAnalysisItemToDashboard(dashboardPage, analysisItem, 0, 0);
            TDAPIOLELib.DashboardPageItem dashboardPage2 = aLM_CORE.Dashboard.AddAnalysisItemToDashboard(dashboardPage, analysisItem1, 0, 1);

            foreach(TDAPIOLELib.DashboardFolder dashboardFolder1 in aLM_CORE.Dashboard.FindChildFolders(aLM_CORE.Dashboard.GetFolderObject("Private\\TestFolder1")))
            {
                Console.WriteLine("Folder found under Private\\TestFolder1 : " + dashboardFolder1.Name);
            }

            foreach (TDAPIOLELib.DashboardPage dp in aLM_CORE.Dashboard.FindChildPages(dashboardFolder))
            {
                Console.WriteLine("Dashboard page found under " + dashboardFolder.Name + " - " + dp.Name);
            }

            aLM_CORE.Dashboard.RenameFolder(dashboardFolder, "TestFolder2");
            Console.WriteLine("New Dashboard Folder name is : " + dashboardFolder.Name);

            aLM_CORE.Dashboard.RenamePage(dashboardPage, "PageName1");
            Console.WriteLine("New Dashboard page name is : " + dashboardPage.Name);


            aLM_CORE.Dashboard.DeletePageItem(dashboardPage1);
            aLM_CORE.Dashboard.DeletePageItem(dashboardPage2);

            aLM_CORE.Dashboard.DeletePage(dashboardPage);
            aLM_CORE.Dashboard.DeleteFolder(dashboardFolder);

            aLM_CORE.Dashboard.DeleteFolder(dashboardFolderParent);

            aLM_CORE.Analysis.DeleteFolder(analysisItemFolder);
            aLM_CORE.Analysis.DeleteFolder(aLM_CORE.Analysis.GetFolderObject("Private\\TestFolder1"));

            Console.WriteLine("Done");
        }
    }
}
