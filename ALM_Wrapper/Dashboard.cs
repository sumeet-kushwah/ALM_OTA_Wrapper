using System;


namespace ALM_Wrapper
{
    public class Dashboard
    {
        private TDAPIOLELib.TDConnection tDConnection;

        //Constructor
        public Dashboard(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        public TDAPIOLELib.List FindChildPages(TDAPIOLELib.DashboardFolder dashboardFolder)
        {
            TDAPIOLELib.DashboardPageFactory dashboardPageFactory = dashboardFolder.DashboardPageFactory;
            TDAPIOLELib.List list = new TDAPIOLELib.List();
            foreach (TDAPIOLELib.DashboardPage DP in dashboardPageFactory.NewList(""))
            {
                list.Add(DP);
            }

            return list;
        }

        public TDAPIOLELib.List FindChildFolders(TDAPIOLELib.DashboardFolder dashboardFolder)
        {
            TDAPIOLELib.DashboardFolderFactory dashboardFolderFactory = dashboardFolder.DashboardFolderFactory;
            TDAPIOLELib.List list = new TDAPIOLELib.List();
            foreach (TDAPIOLELib.DashboardFolder DF in dashboardFolderFactory.NewList(""))
            {
                list.Add(DF);
            }

            return list;
        }

        public TDAPIOLELib.DashboardFolder CreateFolderPath(String FolderPath)
        {
            TDAPIOLELib.DashboardFolderFactory dashboardFolderFactory = tDConnection.DashboardFolderFactory;
            TDAPIOLELib.TDFilter tdFilter;
            TDAPIOLELib.List list;
            TDAPIOLELib.DashboardFolder dashboardFolder = null;

            foreach (String folder in FolderPath.Split('\\'))
            {
                tdFilter = dashboardFolderFactory.Filter;
                tdFilter["DF_NAME"] = folder;
                list = dashboardFolderFactory.NewList(tdFilter.Text);
                if (list.Count <= 0)
                    dashboardFolder = CreateFolder(dashboardFolder, folder);
                else
                    dashboardFolder = list[1];

                dashboardFolderFactory = dashboardFolder.DashboardFolderFactory;
            }

            return dashboardFolder;
        }

        //public TDAPIOLELib.DashboardFolder CreateFolder(String FolderPath)
        //{
        //    TDAPIOLELib.DashboardFolderFactory dashboardFolderFactory = tDConnection.DashboardFolderFactory;
        //    TDAPIOLELib.TDFilter tdFilter;
        //    TDAPIOLELib.List list;
        //    TDAPIOLELib.DashboardFolder dashboardFolder = null;

        //    foreach (String folder in FolderPath.Split('\\'))
        //    {
        //        tdFilter = dashboardFolderFactory.Filter;
        //        tdFilter["DF_NAME"] = folder;
        //        list = dashboardFolderFactory.NewList(tdFilter.Text);
        //        dashboardFolder = list[1];

        //        dashboardFolderFactory = dashboardFolder.DashboardFolderFactory;
        //    }

        //    return dashboardFolder;
        //}



        public TDAPIOLELib.DashboardFolder CreateFolder(TDAPIOLELib.DashboardFolder parentFolder, String newFolderName)
        {
            TDAPIOLELib.DashboardFolder dashboardFolder = parentFolder.DashboardFolderFactory.AddItem(System.DBNull.Value);
            dashboardFolder.Name = newFolderName;
            dashboardFolder.Post();
            return dashboardFolder;
        }

        public Boolean RenameFolder(TDAPIOLELib.DashboardFolder dashboardFolder, String newFolderName)
        {
            dashboardFolder.Name = newFolderName;
            dashboardFolder.Post();
            return true;
        }

        public Boolean DeleteFolder(TDAPIOLELib.DashboardFolder dashboardFolder)
        {
            TDAPIOLELib.DashboardFolderFactory dashboardFolderFactory = tDConnection.DashboardFolderFactory;
            dashboardFolderFactory.RemoveItem(dashboardFolder.ID);
            return true;
        }

        public TDAPIOLELib.DashboardPage CreatePage(TDAPIOLELib.DashboardFolder parentFolder, String PageName)
        {
            TDAPIOLELib.DashboardPageFactory dashboardPageFactory = parentFolder.DashboardPageFactory;
            TDAPIOLELib.DashboardPage dashboardPage = dashboardPageFactory.AddItem(System.DBNull.Value);
            dashboardPage.Name = PageName;
            dashboardPage.Post();
            return dashboardPage;
        }

        public Boolean RenamePage(TDAPIOLELib.DashboardPage dashboardPage, String newName)
        {
            dashboardPage.Name = newName;
            dashboardPage.Post();
            return true;
        }

        public Boolean DeletePage(TDAPIOLELib.DashboardPage dashboardPage)
        {
            TDAPIOLELib.DashboardPageFactory dashboardPageFactory = tDConnection.DashboardPageFactory;
            dashboardPageFactory.RemoveItem(dashboardPage.ID);
            return true;
        }

        public TDAPIOLELib.DashboardPageItem AddAnalysisItemToDashboard(TDAPIOLELib.DashboardPage dashboardPage, TDAPIOLELib.AnalysisItem analysisItem, int Row, int Column)
        {
            TDAPIOLELib.DashboardPageItemFactory dashboardPageItemFactory = dashboardPage.DashboardPageItemFactory;
            TDAPIOLELib.DashboardPageItem dashboardPageItem = dashboardPageItemFactory.AddItem(System.DBNull.Value);
            dashboardPageItem.AnalysisItemId = analysisItem.ID;
            dashboardPageItem.Column = Column;
            dashboardPageItem.Row = Row;
            dashboardPageItem.Post();
            return dashboardPageItem;
        }

        public Boolean DeletePageItem(TDAPIOLELib.DashboardPageItem dashboardPageItem)
        {
            TDAPIOLELib.DashboardPageItemFactory dashboardPageItemFactory = tDConnection.DashboardPageItemFactory;
            dashboardPageItemFactory.RemoveItem(dashboardPageItem.ID);
            return true;
        }

    }
}
