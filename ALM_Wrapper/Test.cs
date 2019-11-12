using System;
using System.Collections.Generic;

namespace ALM_Wrapper
{
    /// <summary>
    /// This class provides helper methods for working on ALM Tests in Test Plan
    /// </summary>
    public class Test
    {
        /// <summary>
        /// TDAPIOLELib.TDConnection Object for the current ALM Connection
        /// </summary>
        private TDAPIOLELib.TDConnection tDConnection;

        /// <summary>
        /// Creates Helper Test Class Object
        /// </summary>
        /// <param name="OALMConnection">Pass TDConnection object to create the Test Object.</param>
        public Test(TDAPIOLELib.TDConnection OALMConnection)
        {
            this.tDConnection = OALMConnection;
        }


        public TDAPIOLELib.List GetTestsWithFilters(TDAPIOLELib.TDFilter tDFilter)
        {
            TDAPIOLELib.TestFactory OTestFactory = tDConnection.TestFactory as TDAPIOLELib.TestFactory;
            return OTestFactory.NewList(tDFilter.Text);
        }

        /// <summary>
        /// Finds Test plan Test case using test case ID. 
        /// <para/> returns TDAPIOLELib.Test Object
        /// </summary>
        /// <param name="id">Test ID of test case from test plan</param>
        /// <returns>TDAPIOLELib.Test Object</returns>
        public TDAPIOLELib.Test GetObjectWithID(int id)
        {
            TDAPIOLELib.TestFactory OTestFactory = tDConnection.TestFactory as TDAPIOLELib.TestFactory;
            TDAPIOLELib.TDFilter OTDFilter = OTestFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List OTestList;

            TDAPIOLELib.Test OTest;

            try
            {

                OTDFilter["TS_TEST_ID"] = Convert.ToString(id);
                OTestList = OTestFactory.NewList(OTDFilter.Text);

                if (OTestList != null && OTestList.Count == 1)
                {
                    OTest = OTestList[1];
                    return OTest;
                }
                else
                {
                    throw (new Exception("Unable to find test with ID : " + id));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Count all tests in TestPlan. 
        /// <para/> returns Count of tests from test plan
        /// </summary>
        /// <returns>Count of tests from test plan</returns>
        public int CountAllTests()
        {
            TDAPIOLELib.Recordset ORecordSet = Utilities.ExecuteQuery("Select Count(*) from Test", tDConnection);
            ORecordSet.First();
            return Convert.ToInt32(ORecordSet[0]);
        }

        /// <summary>
        /// Count tests under a test plan folder. 
        /// <para/> returns Count of tests under test folder
        /// </summary>
        /// <param name="folderPath">Path of test plan folder. Path should start from "Subject\"</param>
        /// <returns>Count of tests under test folder</returns>
        public int CountTestUnderFolder(String folderPath)
        {
            TDAPIOLELib.TreeManager OTManager = tDConnection.TreeManager;
            var OTFolder = OTManager.get_NodeByPath(folderPath);
            TDAPIOLELib.TestFactory OTFactory = OTFolder.TestFactory;
            return OTFactory.NewList("").Count;
        }

        /// <summary>
        /// Counts tests under a test folder
        /// </summary>
        /// <param name="testFolder">Test Folder</param>
        /// <returns>Count of tests inside a test folder</returns>
        public int CountTestUnderFolder(TDAPIOLELib.TestFolder testFolder)
        {
            TDAPIOLELib.TestFactory testFactory = testFolder.TestFactory;
            TDAPIOLELib.TDFilter tDFilter = testFactory.Filter;
            return tDFilter.NewList().Count;
        }

        /// <summary>
        /// Delete a test from test plan. 
        /// <para/> returns true if test is deleted from test plan
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test object</param>
        /// <returns>true if test is deleted from test plan</returns>
        public Boolean Delete(TDAPIOLELib.Test test)
        {
            try
            {
                TDAPIOLELib.TestFactory OTFactory = tDConnection.TestFactory;
                OTFactory.RemoveItem(test.ID);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Download attachments of a test case. 
        /// <para/> returns true if the download is successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test test Object</param>
        /// <param name="attachmentDownloadPath">Folder path to download the attachments</param>
        /// <returns>true if the download is successfull</returns>
        public Boolean DownloadAttachments(TDAPIOLELib.Test test, String attachmentDownloadPath)
        {
            return Utilities.DownloadAttachments(test.Attachments, attachmentDownloadPath);

        }

        /// <summary>
        /// Get the test plan folder path of a test case. 
        /// <para/> returns test plan path of the test case
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <returns>test plan path of the test case</returns>
        public String GetPath(TDAPIOLELib.Test test)
        {
            try
            {
                return test["TS_SUBJECT"].Path;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Add design steps to test. description[0], expectedResults[0] will become first step and attachmentsPath[0] will be attached to it. 
        /// <para/> returns true if design steps are added
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="description">List of all descriptions to be added to test. For each description in the list one design step will be added to the test case</param>
        /// <param name="expectedResults">List of all expected results to be added to the test.</param>
        /// <param name="attachmentsPath">List of all attachemnts to be added. </param>
        /// <returns>true if design steps are added</returns>
        public Boolean AddDesignSteps(TDAPIOLELib.Test test, List<String> description, List<String> expectedResults, List<String> attachmentsPath)
        {
            for (int Counter = 0; Counter < description.Count; Counter++)
            {
                AddSingleDeignStep(test, description[Counter], expectedResults[Counter], attachmentsPath[Counter]);
            }

            return true;
        }

        /// <summary>
        /// Add a single design step to the test. Added design step will be the last design step. 
        /// <para/> returns TDAPIOLELib.DesignStep Object
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="description">Description of the design step</param>
        /// <param name="expectedResult">Expected result of the design step</param>
        /// <param name="attachmentPath">attachment path of the design step</param>
        /// <returns>TDAPIOLELib.DesignStep Object</returns>
        public TDAPIOLELib.DesignStep AddSingleDeignStep(TDAPIOLELib.Test test, String description, String expectedResult, String attachmentPath = "")
        {
            TDAPIOLELib.DesignStep ODesignStep;
            TDAPIOLELib.DesignStepFactory ODesignStepFactory;

            ODesignStepFactory = test.DesignStepFactory;

            int StepCounter = ODesignStepFactory.NewList("").Count + 1;
      
            ODesignStep = ODesignStepFactory.AddItem(System.DBNull.Value);
            ODesignStep.StepName = "Step " + StepCounter;
            ODesignStep.StepDescription = description;
            ODesignStep.StepExpectedResult = expectedResult;
            ODesignStep.Post();

            if (attachmentPath != "")
            {
                TDAPIOLELib.AttachmentFactory OAttachmentFactory;
                TDAPIOLELib.Attachment OAttachment;

                OAttachmentFactory = ODesignStep.Attachments;
                if (System.IO.File.Exists(attachmentPath))
                {
                    OAttachment = OAttachmentFactory.AddItem(System.DBNull.Value);
                    OAttachment.FileName = attachmentPath;
                    OAttachment.Type = Convert.ToInt16(TDAPIOLELib.tagTDAPI_ATTACH_TYPE.TDATT_FILE);
                    OAttachment.Post();
                }
                else
                {
                    throw (new Exception("File Not Found : " + attachmentPath));
                }
            }

            return ODesignStep;
        }

        /// <summary>
        /// Adds attachment to the test. 
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="attachmentsPath">List of all attachments to be added to test</param>
        /// <returns>true if successfull</returns>
        public Boolean AddAttachment(TDAPIOLELib.Test test, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(test.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Adds attachment to the designStep. 
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.DesignStep Object</param>
        /// <param name="attachmentsPath">List of all attachments to be added to test</param>
        /// <returns>true if successfull</returns>
        public Boolean AddAttachmentToDesignStep(TDAPIOLELib.DesignStep designStep, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(designStep.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Adds parameters to the test. 
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="paramName">List of parameter names to be added to tests</param>
        /// <param name="paramDescription">List of parameter descriptions for the parameter names</param>
        /// <param name="paramDefaultValue">List of default values for the parameter names</param>
        /// <returns>true if successfull</returns>
        public Boolean AddParameters(TDAPIOLELib.Test test, List<String> paramName, List<String> paramDescription, List<String> paramDefaultValue)
        {
            TDAPIOLELib.TestParameterFactory OTestParamFactory;
            TDAPIOLELib.ISupportTestParameters OISupportTestParams;
            TDAPIOLELib.TestParameter OTestParameter;

            OISupportTestParams = test as TDAPIOLELib.ISupportTestParameters;
            OTestParamFactory = OISupportTestParams.TestParameterFactory as TDAPIOLELib.TestParameterFactory;

            for (int Counter = 0; Counter < paramName.Count; Counter++)
            {
                OTestParameter = OTestParamFactory.AddItem(System.DBNull.Value);
                OTestParameter.Name = paramName[Counter];
                OTestParameter.Description = paramDescription[Counter];
                OTestParameter.DefaultValue = paramDefaultValue[Counter];
                OTestParameter.Post();
            }

            return true;
        }

        /// <summary>
        /// Adds a single parameter to test. 
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="paramName">parameter name</param>
        /// <param name="paramDescription">parameter descriptio</param>
        /// <param name="ParamDefaultValue">paramater default value</param>
        /// <returns>true if successfull</returns>
        public Boolean AddSingleParameter(TDAPIOLELib.Test test, String paramName, String paramDescription, String paramDefaultValue = "")
        {
            TDAPIOLELib.TestParameterFactory OTestParamFactory;
            TDAPIOLELib.ISupportTestParameters OISupportTestParams;
            TDAPIOLELib.TestParameter OTestParameter;

            OISupportTestParams = test as TDAPIOLELib.ISupportTestParameters;
            OTestParamFactory = OISupportTestParams.TestParameterFactory as TDAPIOLELib.TestParameterFactory;
            
            OTestParameter = OTestParamFactory.AddItem(System.DBNull.Value);
            OTestParameter.Name = paramName;
            OTestParameter.Description = paramDescription;
            OTestParameter.DefaultValue = paramDefaultValue;
            OTestParameter.Post();
            
            return true;
        }

        /// <summary>
        /// Creats a new test case. 
        /// <para/> returns TDAPIOLELib.Test object
        /// </summary>
        /// <param name="testDetails">Dictionary object with the values for each field. Field names must be the data base field names. e.g. TS_NAME = 'Example Test'</param>
        /// <param name="path">test plan path, test will be created under this folder</param>
        /// <returns>TDAPIOLELib.Test object</returns>
        public TDAPIOLELib.Test Create(Dictionary<String, String> testDetails, String path)
        {
            TDAPIOLELib.TreeManager OTreeManager = tDConnection.TreeManager;
            TDAPIOLELib.Test OTest;
            TDAPIOLELib.TestFactory OTestFactory;
            var testFolder = OTreeManager.NodeByPath["subject"];

            if (path != "")
            {
                testFolder = OTreeManager.get_NodeByPath(path);
            }
            else
            {
                throw (new Exception("Path is required for creating tests"));
            }
            
            //Create test here
            OTestFactory = testFolder.TestFactory;
            OTest = OTestFactory.AddItem(System.DBNull.Value);
            foreach (KeyValuePair<string, string> kvp in testDetails)
            {
                OTest[kvp.Key.ToUpper()] = kvp.Value;
            }
            OTest.Post();
            return OTest;
        }

        /// <summary>
        /// Filters tests from test plan using the filter text. Filter text can be colpied directly from ALM and passed to this function. 
        /// <para/> returns TDAPIOLELib.List Object. Each item from this list can be coverted to TDAPIOLELib.Test object.
        /// </summary>
        /// <param name="FilterString">Copy this from ALM</param>
        /// <returns>TDAPIOLELib.List Object. Each item from this list can be coverted to TDAPIOLELib.Test object.</returns>
        public TDAPIOLELib.List Filter(String FilterString)
        {
            TDAPIOLELib.TestFactory OTestFactory = tDConnection.TestFactory as TDAPIOLELib.TestFactory;
            TDAPIOLELib.TDFilter OTDFilter = OTestFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List OTestList;

            try
            {
                OTDFilter.Text = FilterString;
                OTestList = OTestFactory.NewList(OTDFilter.Text);

                return OTestList;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Deletes all test attachments
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <returns>true if successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.Test test)
        {
            return Utilities.DeleteAllAttachments(test.Attachments);            
        }

        /// <summary>
        /// Search and delete test attachment using attachment name
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="attachmentName">Name of the attachment</param>
        /// <returns>true if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.Test test, String attachmentName)
        {
           return Utilities.DeleteAttachmentByName(test.Attachments, attachmentName);
        }

        /// <summary>
        /// Update the field values for the test.
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="fieldName">database field name</param>
        /// <param name="newValue">new value for the field</param>
        /// <returns></returns>
        public Boolean UpdateFieldValue(TDAPIOLELib.Test test, String fieldName, String newValue, Boolean Post = true)
        {
            test[fieldName.ToUpper()] = newValue;
            if (Post)
            test.Post();
            return true;
        }

        /// <summary>
        /// Assign requirement to the test case
        /// <para/> returns true if successfull
        /// </summary>
        /// <param name="test"></param>
        /// <param name="requirement"></param>
        /// <returns>true if successfull</returns>
        public Boolean AssignRequirement(TDAPIOLELib.Test test, TDAPIOLELib.Req requirement)
        {
            TDAPIOLELib.ICoverableReq OICoverable = requirement as TDAPIOLELib.ICoverableReq;
            OICoverable.AddTestToCoverage(test.ID);
            return true;
        }

        /// <summary>
        /// Gets the list of design steps for the test object
        /// <para/> returns TDAPIOLELib.List. Each list item can be converted to the TDAPIOLELib.DesignStep
        /// </summary>
        /// <param name="test"></param>
        /// <returns>TDAPIOLELib.List. Each list item can be converted to the TDAPIOLELib.DesignStep</returns>
        public TDAPIOLELib.List GetDesignStepList(TDAPIOLELib.Test test)
        {
            return test.DesignStepFactory.NewList("");
        }

        /// <summary>
        /// Mark the test as Template test
        /// <para/> returns true is successfully
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <returns>true is successfully</returns>
        public Boolean MarkAsTemplate(TDAPIOLELib.Test test)
        {
            test.TemplateTest = true;
            test.Post();
            return true;
        }

        /// <summary>
        /// Finds the list of tests under unattached folder
        /// <para/>returns TDAPIOLELib.List Object. Each item from this list can be converted to TDAPIOLELib.Test Object
        /// </summary>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List FindUnattachedTests()
        {
            TDAPIOLELib.List list = new TDAPIOLELib.List();
            TDAPIOLELib.Recordset recordset = Utilities.ExecuteQuery("Select * from test where TS_SUBJECT = -2", tDConnection);
            recordset.First();
            for (int Counter = 1; Counter <= recordset.RecordCount; Counter++)
            {
                list.Add(GetObjectWithID(Convert.ToInt32(recordset["TS_TEST_ID"])));
                recordset.Next();
            }
            return list;
        }

        /// <summary>
        /// Download design step attachments 
        /// <para/> return true if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="attachmentDownloadPath">Path</param>
        /// <returns>true if successfull</returns>
        public Boolean DownloadDesignStepsAttachments(TDAPIOLELib.Test test, String attachmentDownloadPath)
        {
            TDAPIOLELib.TestFactory testFactory = tDConnection.TestFactory as TDAPIOLELib.TestFactory;

            foreach (TDAPIOLELib.DesignStep ODesignStep in GetDesignStepList(test))
            {
                if (ODesignStep.HasAttachment)
                {
                    Utilities.DownloadAttachments(ODesignStep.Attachments, attachmentDownloadPath);
                }
            }
                
            return true;
        }

        /// <summary>
        /// Add Test Configurations for a test plan test
        /// <para/> returns TDAPIOLELib.TestConfig Object if successfull
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="testConfigName">Name for the new test config</param>
        /// <param name="testConfigDesc">Description for the new test config</param>
        /// <returns>return TDAPIOLELib.TestConfig Object if successfull</returns>
        public TDAPIOLELib.TestConfig AddTestConfiguration(TDAPIOLELib.Test test, String testConfigName, String testConfigDesc = "")
        {
            TDAPIOLELib.ITestConfigFactory testConfigFactory = test.TestConfigFactory;
            TDAPIOLELib.TestConfig testConfig = testConfigFactory.AddItem(System.DBNull.Value);

            testConfig["TSC_NAME"] = testConfigName;
            testConfig["TSC_DESC"] = testConfigDesc;

            testConfig.Post();
            return testConfig;
        }

        public Boolean AddParamValueForTestConfig(TDAPIOLELib.TestConfig testConfig, String ParamName, String ParamValue)
        {
            TDAPIOLELib.ISupportParameterValues supportParameterValues;
            TDAPIOLELib.ParameterValueFactory parameterValueFactory;
            

            supportParameterValues = testConfig as TDAPIOLELib.ISupportParameterValues;
            parameterValueFactory = supportParameterValues.ParameterValueFactory;

            foreach (TDAPIOLELib.ParameterValue parameterValue in parameterValueFactory.NewList(""))
            {
                if (parameterValue.Name == ParamName)
                {
                    parameterValue.ActualValue = ParamValue;
                    parameterValue.Post();
                }
            }

            return true;
        }

    }
}
