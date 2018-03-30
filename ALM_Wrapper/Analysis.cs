using System;
using System.Collections.Generic;
using System.Xml;

namespace ALM_Wrapper
{
    public class Analysis
    {
        private TDAPIOLELib.TDConnection tDConnection;

        
        public enum DefectSummaryGraphSumOF { ActualFixTime = 0, EstimatedFixTime = 1 , None = 2};

        public enum DefectAgeGrouping { NoGrouping = 0, OneWeek = 1, OneMonth = 2, SixMonths = 3, OneYear = 4 };

        //Constructor
        public Analysis(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        /// <summary>
        /// Gets the folder object for the public folder
        /// </summary>
        /// <returns>TDAPIOLELib.AnalysisItemFolder Object</returns>
        public TDAPIOLELib.AnalysisItemFolder FindPublicFolder()
        {
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = tDConnection.AnalysisItemFolderFactory;

            TDAPIOLELib.TDFilter tDFilter = analysisItemFolderFactory.Filter;
            tDFilter["AIF_NAME"] = "Public";
            tDFilter["AIF_OWNER"] = "__default__";

            TDAPIOLELib.List list = analysisItemFolderFactory.NewList(tDFilter.Text);

            foreach (TDAPIOLELib.AnalysisItemFolder OAF in list)
            {
                if (OAF.Name.ToUpper() == "PUBLIC")
                {
                    return OAF;
                }
            }

            return null;
        }

        /// <summary>
        /// Gets the folder object for the private folder
        /// </summary>
        /// <returns>TDAPIOLELib.AnalysisItemFolder Object</returns>
        public TDAPIOLELib.AnalysisItemFolder FindPrivateFolder()
        {
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = tDConnection.AnalysisItemFolderFactory;

            TDAPIOLELib.TDFilter tDFilter = analysisItemFolderFactory.Filter;
            tDFilter["AIF_NAME"] = "Private";
            tDFilter["AIF_OWNER"] = "__default__";

            TDAPIOLELib.List list = analysisItemFolderFactory.NewList(tDFilter.Text);

            foreach (TDAPIOLELib.AnalysisItemFolder OAF in list)
            {
                if (OAF.Name.ToUpper() == "PRIVATE")
                {
                    return OAF;
                }
            }

            return null;
        }

        /// <summary>
        /// Get the Child Folder names for Analysis Folder
        /// </summary>
        /// <param name="folderPath">Path of the folder to search for the Analysis</param>
        /// <returns>List of Analysis folder names</returns>
        public List<String> GetChildFolderNames(String folderPath)
        {
            List<String> anaFolderNames = new List<string>();
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = GetFolderObject(folderPath);
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = analysisItemFolder.AnalysisItemFolderFactory;
            foreach (TDAPIOLELib.AnalysisItemFolder aif in analysisItemFolderFactory.NewList(""))
            {
                anaFolderNames.Add(aif.Name);
            }

            return anaFolderNames;
        }

        /// <summary>
        /// Get the Child Folder names for Analysis Folder
        /// </summary>
        /// <param name="folderPath">TDAPIOLELib.AnalysisItemFolder Object</param>
        /// <returns>List of Analysis folder names</returns>
        public List<String> GetChildFolderNames(TDAPIOLELib.AnalysisItemFolder analysisItemFolder)
        {
            List<String> anaFolderNames = new List<string>();
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = analysisItemFolder.AnalysisItemFolderFactory;
            foreach (TDAPIOLELib.AnalysisItemFolder aif in analysisItemFolderFactory.NewList(""))
            {
                anaFolderNames.Add(aif.Name);
            }

            return anaFolderNames;
        }

        /// <summary>
        /// Renames a folder
        /// </summary>
        /// <param name="analysisItemFolder">Analysis Folder Object</param>
        /// <param name="newFolderName">New folder name</param>
        /// <returns></returns>
        public Boolean RenameFolder(TDAPIOLELib.AnalysisItemFolder analysisItemFolder, String newFolderName)
        {
            analysisItemFolder.Name = newFolderName;
            analysisItemFolder.Post();
            return true;
        }

        /// <summary>
        /// Renames a report or graph
        /// </summary>
        /// <param name="analysisItem">Analysis Item object</param>
        /// <param name="newItemName">new item name</param>
        /// <returns></returns>
        public Boolean RenameReportOrGraph(TDAPIOLELib.AnalysisItem analysisItem, String newItemName)
        {
            analysisItem.Name = newItemName;
            analysisItem.Post();
            return true;
        }

        /// <summary>
        /// Get the Child Folders for Analysis Folder
        /// </summary>
        /// <param name="folderPath">Path of the folder to search for the release</param>
        /// <returns>List of Analysis folder names</returns>
        public TDAPIOLELib.List GetChildFolders(String folderPath)
        {
            TDAPIOLELib.List anaFolderNames = new TDAPIOLELib.List();
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = GetFolderObject(folderPath);
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = analysisItemFolder.AnalysisItemFolderFactory;
            foreach (TDAPIOLELib.AnalysisItemFolder aif in analysisItemFolderFactory.NewList(""))
            {
                anaFolderNames.Add(aif);
            }

            return anaFolderNames;
        }

        /// <summary>
        /// Returns a list of reports or graphs under analysis folder
        /// </summary>
        /// <param name="analysisItemFolder">Analysis folder object</param>
        /// <returns></returns>
        public TDAPIOLELib.List GetChildGraphs(TDAPIOLELib.AnalysisItemFolder analysisItemFolder)
        {
            TDAPIOLELib.List anaFolderNames = new TDAPIOLELib.List();
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = analysisItemFolder.AnalysisItemFactory;
            foreach (TDAPIOLELib.AnalysisItemFolder aif in analysisItemFactory.NewList(""))
            {
                anaFolderNames.Add(aif);
            }

            return anaFolderNames;
        }

        /// <summary>
        /// Returns list of reports and graphs under analysis folder
        /// </summary>
        /// <param name="folderPath">Analysis folder path</param>
        /// <returns></returns>
        public TDAPIOLELib.List GetChildGraphs(String folderPath)
        {
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = GetFolderObject(folderPath);
            return GetChildGraphs(analysisItemFolder);
        }

        /// <summary>
        /// Returns list of graph and report names under analysis folder
        /// </summary>
        /// <param name="analysisItemFolder">Analysis folder object</param>
        /// <returns></returns>
        public List<String> GetChildGraphNames(TDAPIOLELib.AnalysisItemFolder analysisItemFolder)
        {
            List<String> list = new List<String>();
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = analysisItemFolder.AnalysisItemFactory;
            foreach (TDAPIOLELib.AnalysisItemFolder aif in analysisItemFactory.NewList(""))
            {
                list.Add(aif.Name);
            }

            return list;
        }

        /// <summary>
        /// Returns list of graph and report names under analysis folder
        /// </summary>
        /// <param name="folderPath">String path of the folder</param>
        /// <returns></returns>
        public List<String> GetChildGraphNames(String folderPath)
        {
            List<String> list = new List<String>();
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = GetFolderObject(folderPath);
            return GetChildGraphNames(analysisItemFolder);
        }

        /// <summary>
        /// Creates a new Analysis Folder
        /// </summary>
        /// <param name="parentFolderPath">String path for parent folder</param>
        /// <param name="newFolderName">new folder name</param>
        /// <returns>TDAPIOLELib.AnalysisItemFolder Object</returns>
        public TDAPIOLELib.AnalysisItemFolder CreateFolder(String parentFolderPath, String newFolderName)
        {
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = tDConnection.AnalysisItemFolderFactory;
            TDAPIOLELib.TDFilter tdFilter;
            TDAPIOLELib.List list;
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder;

            foreach (String folder in parentFolderPath.Split('\\'))
            {
                tdFilter = analysisItemFolderFactory.Filter;
                tdFilter["AIF_NAME"] = folder;
                list = analysisItemFolderFactory.NewList(tdFilter.Text);
                analysisItemFolder = list[1];
                analysisItemFolderFactory = analysisItemFolder.AnalysisItemFolderFactory;
            }

            analysisItemFolder = analysisItemFolderFactory.AddItem(System.DBNull.Value);
            analysisItemFolder.Name = newFolderName;
            analysisItemFolder.Post();
            
            return analysisItemFolder;
        }

        /// <summary>
        /// Creates a new Analysis Folder
        /// </summary>
        /// <param name="parentFolder">AnalysisItemFolder object for parent folder</param>
        /// <param name="newFolderName">new folder name</param>
        /// <returns>TDAPIOLELib.AnalysisItemFolder Object</returns>
        public TDAPIOLELib.AnalysisItemFolder CreateFolder(TDAPIOLELib.AnalysisItemFolder parentFolder, String newFolderName)
        {
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = parentFolder.AnalysisItemFolderFactory.AddItem(System.DBNull.Value);
            analysisItemFolder.Name = newFolderName;
            analysisItemFolder.Post();
            return analysisItemFolder;
        }

        /// <summary>
        /// Creates a folder path. All missing folders in the folder path will be created
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns>TDAPIOLELib.AnalysisItemFolder Object</returns>
        public TDAPIOLELib.AnalysisItemFolder CreateFolderPath(string folderPath)
        {
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = tDConnection.AnalysisItemFolderFactory;
            TDAPIOLELib.TDFilter tdFilter;
            TDAPIOLELib.List list;
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = null;

            foreach (String folder in folderPath.Split('\\'))
            {
                tdFilter = analysisItemFolderFactory.Filter;
                tdFilter["AIF_NAME"] = folder;
                list = analysisItemFolderFactory.NewList(tdFilter.Text);
                if (list.Count <= 0)
                    analysisItemFolder = CreateFolder(analysisItemFolder, folder);
                else
                    analysisItemFolder = list[1];

                analysisItemFolderFactory = analysisItemFolder.AnalysisItemFolderFactory;
            }

            return analysisItemFolder;
        }

        /// <summary>
        /// Get TDAPIOLELib.AnalysisItemFolder Object from string path
        /// </summary>
        /// <param name="folderPath">Path of the folder to be created</param>
        /// <returns>TDAPIOLELib.AnalysisItemFolder Object</returns>
        public TDAPIOLELib.AnalysisItemFolder GetFolderObject(String folderPath)
        {
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = tDConnection.AnalysisItemFolderFactory;
            TDAPIOLELib.TDFilter tdFilter;
            TDAPIOLELib.List list;
            TDAPIOLELib.AnalysisItemFolder analysisItemFolder = null;

            foreach (String folder in folderPath.Split('\\'))
            {
                tdFilter = analysisItemFolderFactory.Filter;
                tdFilter["AIF_NAME"] = folder;
                list = analysisItemFolderFactory.NewList(tdFilter.Text);
                if (list.Count <= 0)
                    throw (new Exception("Analysis Folder Not Found : " + folder));
                else
                    analysisItemFolder = list[1];

                analysisItemFolderFactory = analysisItemFolder.AnalysisItemFolderFactory;
            }

            return analysisItemFolder;
        }

        /// <summary>
        /// Delete a report or graph
        /// </summary>
        /// <param name="ID">ID of the report or graph to be deleted</param>
        /// <returns></returns>
        public Boolean DeleteReportOrGraph(int ID)
        {
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = tDConnection.AnalysisItemFactory;
            analysisItemFactory.RemoveItem(ID);
            return true;
        }

        /// <summary>
        /// Deletes report or graph
        /// </summary>
        /// <param name="analysisItem">analysisItem Object for graph or report</param>
        /// <returns></returns>
        public Boolean DeleteReportOrGraph(TDAPIOLELib.AnalysisItem analysisItem)
        {
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = tDConnection.AnalysisItemFactory;
            analysisItemFactory.RemoveItem(analysisItem.ID);
            return true;
        }

        /// <summary>
        /// Deletes Analysis Folder
        /// </summary>
        /// <param name="analysisItemFolder">AnalysisItemFolder Object</param>
        /// <returns></returns>
        public Boolean DeleteFolder(TDAPIOLELib.AnalysisItemFolder analysisItemFolder)
        {
            TDAPIOLELib.AnalysisItemFolderFactory analysisItemFolderFactory = tDConnection.AnalysisItemFolderFactory;
            analysisItemFolderFactory.RemoveItem(analysisItemFolder.ID);
            return true;
        }

        /// <summary>
        /// Creates a Defect Age graph under analysis folder
        /// </summary>
        /// <param name="analysisItemFolder">AnalysisItemFolder Object</param>
        /// <param name="graphName">Name of the graph</param>
        /// <param name="groupByField">Group by field name</param>
        /// <param name="sumOfField">Sum of field, if this is emplty then count of will be automatically selected</param>
        /// <param name="ageGrouping">Age Grouping Enum object.</param>
        /// <param name="filterString">Filter defects. You can get this using TDFilter Object</param>
        /// <returns></returns>
        public TDAPIOLELib.AnalysisItem CreateDefectAgeGraph(TDAPIOLELib.AnalysisItemFolder analysisItemFolder, String graphName, String groupByField, DefectSummaryGraphSumOF sumOfField, DefectAgeGrouping ageGrouping, String filterString)
        {
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = analysisItemFolder.AnalysisItemFactory;
            TDAPIOLELib.AnalysisItem analysisItem = analysisItemFactory.AddItem(System.DBNull.Value);
            analysisItem.Name = graphName;
            analysisItem.Type = "Graph";

            String SOF = "";

            if (sumOfField.Equals(DefectSummaryGraphSumOF.ActualFixTime))
                SOF = "BG_ACTUAL_FIX_TIME";
            else if (sumOfField.Equals(DefectSummaryGraphSumOF.EstimatedFixTime))
                SOF = "BG_ESTIMATED_FIX_TIME";
            else if (sumOfField.Equals(DefectSummaryGraphSumOF.None))
                SOF = "";
            else
                throw (new Exception("Invalid Defect SumOf Fields"));

            String AG = "";

            if (ageGrouping == DefectAgeGrouping.NoGrouping)
                AG = "All";
            else if (ageGrouping == DefectAgeGrouping.OneMonth)
                AG = "OneMonth";
            else if (ageGrouping == DefectAgeGrouping.OneWeek)
                AG = "OneWeek";
            else if (ageGrouping == DefectAgeGrouping.SixMonths)
                AG = "SixMonth";
            else if (ageGrouping == DefectAgeGrouping.OneYear)
                AG = "OneYear";

            analysisItem.FilterData = GetDefectAgeGraphFilterData(groupByField, SOF, AG, filterString);
            analysisItem.SubType = "AgeGraph";
            analysisItem.Module = "defect";
            analysisItem.Post();
            return analysisItem;
        }

        private String GetDefectAgeGraphFilterData(String GroupByField, String SumOFField, String AgeGrouping, String FilterString)
        {
            System.IO.StringWriter stringWriter = new System.IO.StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(stringWriter);
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("AnalysisDefinition");
            xmlWriter.WriteAttributeString("Version", "3.0");
            xmlWriter.WriteAttributeString("GraphProviderId", "QC.Graph.Provider");
            xmlWriter.WriteAttributeString("GroupByField", GroupByField);
            xmlWriter.WriteAttributeString("ForceRefresh", "False");
            xmlWriter.WriteAttributeString("SelectedProjects", "CURRENT-PROJECT-UID");
            xmlWriter.WriteAttributeString("SumOfField", SumOFField);
            xmlWriter.WriteAttributeString("AgeGrouping", AgeGrouping);
            xmlWriter.WriteAttributeString("MaxAge", "0");

            xmlWriter.WriteStartElement("Filter");
            xmlWriter.WriteAttributeString("FilterState", "Custom");
            xmlWriter.WriteAttributeString("FilterFormat", "Frec");

            if (FilterString.Length > 0)
                xmlWriter.WriteString("<![CDATA[" + FilterString + "]]>");
            else
                xmlWriter.WriteString("<![CDATA[]]>");

            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndDocument();
            xmlWriter.Close();

            return stringWriter.ToString();
        }

        /// <summary>
        /// Creates a new excel report under analysis folder
        /// </summary>
        /// <param name="analysisItemFolder">AnalysisItemFolder Object</param>
        /// <param name="name">Name of the Excel Report</param>
        /// <param name="query">Query to be added</param>
        /// <returns></returns>
        public TDAPIOLELib.AnalysisItem CreateExcelReport(TDAPIOLELib.AnalysisItemFolder analysisItemFolder, String name, String query)
        {
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = analysisItemFolder.AnalysisItemFactory;
            TDAPIOLELib.AnalysisItem analysisItem = analysisItemFactory.AddItem(System.DBNull.Value);
            analysisItem.Name = name;
            analysisItem.Type = "ExcelReport";
            analysisItem.FilterData = GetExcelReportFilterData(query);
            analysisItem.SubType = "ExcelReport";
            analysisItem.Module = "UnspecifiedEntity";
            analysisItem.Post();
            return analysisItem;
        }

        /// <summary>
        /// Creates a defect summary report
        /// </summary>
        /// <param name="analysisItemFolder">Folder under which graph should be added</param>
        /// <param name="graphName">Name of the Graph</param>
        /// <param name="groupByField">Group by field</param>
        /// <param name="sumOfField">Sum of field, If this field is left blank then Count will be automatically selected</param>
        /// <param name="xAxisField">XAxis field name</param>
        /// <param name="filterString">filter string </param>
        /// <returns></returns>
        public TDAPIOLELib.AnalysisItem CreateDefectSummaryGraph(TDAPIOLELib.AnalysisItemFolder analysisItemFolder, String graphName, String groupByField, DefectSummaryGraphSumOF sumOfField, String xAxisField, String filterString)
        {
            if (sumOfField.Equals(DefectSummaryGraphSumOF.ActualFixTime))
                return CreateSummaryGraph(analysisItemFolder, graphName, groupByField, "BG_ACTUAL_FIX_TIME", xAxisField, filterString, "defect");
            else if (sumOfField.Equals(DefectSummaryGraphSumOF.EstimatedFixTime))
                return CreateSummaryGraph(analysisItemFolder, graphName, groupByField, "BG_ESTIMATED_FIX_TIME", xAxisField, filterString, "defect");
            else if (sumOfField.Equals(DefectSummaryGraphSumOF.None))
                return CreateSummaryGraph(analysisItemFolder, graphName, groupByField, "", xAxisField, filterString, "defect");
            else
                throw (new Exception("Invalid Defect SumOf Fields"));
        }

        private TDAPIOLELib.AnalysisItem CreateSummaryGraph(TDAPIOLELib.AnalysisItemFolder analysisItemFolder, String GraphName, String GroupByField, String SumOfField, String XAxisField, String FilterString, String ReportModule)
        {
            TDAPIOLELib.AnalysisItemFactory analysisItemFactory = analysisItemFolder.AnalysisItemFactory;
            TDAPIOLELib.AnalysisItem analysisItem = analysisItemFactory.AddItem(System.DBNull.Value);
            analysisItem.Name = GraphName;
            analysisItem.Type = "Graph";
            analysisItem.FilterData = GetDefectSummaryGraphFilterData(GroupByField, SumOfField, XAxisField, FilterString);
            analysisItem.SubType = "SummaryGraph";
            analysisItem.Module = ReportModule;
            analysisItem.Post();
            return analysisItem;
        }

        private String GetDefectSummaryGraphFilterData(String GroupByField, String SumOFField, String XAxisField, String FilterString)
        {
            System.IO.StringWriter stringWriter = new System.IO.StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(stringWriter);
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("AnalysisDefinition");
            xmlWriter.WriteAttributeString("Version", "3.0");
            xmlWriter.WriteAttributeString("GraphProviderId", "QC.Graph.Provider");
            xmlWriter.WriteAttributeString("GroupByField", GroupByField);
            xmlWriter.WriteAttributeString("ForceRefresh", "False");
            xmlWriter.WriteAttributeString("SelectedProjects", "CURRENT-PROJECT-UID");
            xmlWriter.WriteAttributeString("SumOfField", SumOFField);
            xmlWriter.WriteAttributeString("XAxisField", XAxisField);
            xmlWriter.WriteAttributeString("ShowFullPath", "False");

                xmlWriter.WriteStartElement("Filter");
                xmlWriter.WriteAttributeString("FilterState", "Custom");
                xmlWriter.WriteAttributeString("FilterFormat", "Frec");

                if (FilterString.Length > 0)
                    xmlWriter.WriteString("<![CDATA["+ FilterString +"]]>");
                else
                    xmlWriter.WriteString("<![CDATA[]]>");

                xmlWriter.WriteEndElement();

                xmlWriter.WriteStartElement("Categories");
                xmlWriter.WriteEndElement();

            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndDocument();
            xmlWriter.Close();

            return stringWriter.ToString();
        }

        private String GetExcelReportFilterData(String Query)
        {
            System.IO.StringWriter stringWriter = new System.IO.StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(stringWriter);
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("AnalysisDefinition");
            xmlWriter.WriteAttributeString("Version", "3.0");
            xmlWriter.WriteAttributeString("GraphProviderId", "QC.Graph.Provider");

                xmlWriter.WriteStartElement("ExcelReportXml");

                    xmlWriter.WriteStartElement("Report");

                        xmlWriter.WriteStartElement("QueryBuilder");
                            xmlWriter.WriteAttributeString("MaxRecords", "8");

                            xmlWriter.WriteStartElement("Sheet");
                                xmlWriter.WriteAttributeString("SheetName", "Query1");
                                xmlWriter.WriteAttributeString("DataFetchScript", Query);
                                xmlWriter.WriteAttributeString("DataFetchMethod", "StandardSQL");
                                xmlWriter.WriteAttributeString("RealSql", Query);
                            xmlWriter.WriteEndElement();

                            xmlWriter.WriteStartElement("GlobalParameters");
                            xmlWriter.WriteEndElement();
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("PostProcessBuilder");
                        xmlWriter.WriteAttributeString("Script", "Sub QC_PostProcessing()" + System.Environment.NewLine  + "Dim MainWorksheet As Worksheet" + System.Environment.NewLine + "' Make sure your worksheet name matches!" + System.Environment.NewLine + "Set MainWorksheet = ActiveWorkbook.Worksheets(\"Query1\")" + System.Environment.NewLine + "Dim DataRange As Range" + System.Environment.NewLine + "Set DataRange = MainWorksheet.UsedRange" + System.Environment.NewLine + "' Now that you have the data in DataRange you can process it." + System.Environment.NewLine + "End Sub");
                        xmlWriter.WriteAttributeString("RunPostProcessing", "False");
                        xmlWriter.WriteEndElement();

                    xmlWriter.WriteStartElement("ReportGenerator");
                    xmlWriter.WriteAttributeString("Open", "true");
                    xmlWriter.WriteAttributeString("SavedFile", "Query1.xls");
                    xmlWriter.WriteAttributeString("OutputFormat", "Excel");
                    xmlWriter.WriteAttributeString("Status", "0");
                    xmlWriter.WriteEndElement();


            
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();

            xmlWriter.WriteEndDocument();
            xmlWriter.Close();

            return stringWriter.ToString();
        }

        //public Boolean GenerateExcelReport(TDAPIOLELib.AnalysisItem analysisItem)
        //{
        //    TDAPIOLELib.ExcelReportManager excelReportManager = tDConnection.ExcelReportManager;
        //    Object Obj = excelReportManager.GetReportQueryExecutor(analysisItem.ID, "Query1");
        //    //Console.WriteLine(Obj.GetType().FullName);
        //    TDAPIOLELib.ExcelReportQuery reportManager = (TDAPIOLELib.ExcelReportQuery)Obj;
        //    //reportManager.StartGenerateReport(analysisItem.ID);
        //    //reportManager.DownloadReport("C:\\temp\temp.xls");
        //    //reportManager.
        //    analysisItem.
        //    return true;
        //    //analysisItem.AnalysisItemFileFactory;
        //}

        //public Boolean CreateFilterData()
        //{
        //    OTAXMLLib.Convert convert = new OTAXMLLib.Convert();
        //    convert.Connection = tDConnection;

        //    int count;

        //    String Config = "<?xml version='1.0' encoding='UTF-8'?><config> <report_config td_type='bug' attach='y' history='y' view='detailed' attr_flag='0'></report_config></config>";
        //    Console.WriteLine(convert.ReadObject("", Config, out count));

        //    OTAREPORTLib.Reporter reporter = new OTAREPORTLib.Reporter();
        //    reporter.Connection = tDConnection;
        //    OTAREPORTLib.ReportConfig reportConfig = new OTAREPORTLib.ReportConfig();

        //    reportConfig.Type = "";

        //    return true;
        //}


    }
}
