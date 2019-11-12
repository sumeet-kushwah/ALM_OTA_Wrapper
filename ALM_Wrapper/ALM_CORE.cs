using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ALM_Wrapper
{
    
    public class ALM_CORE
    {
        //Object will hold the connection 
       
        public TDAPIOLELib.TDConnection tDConnection = new TDAPIOLELib.TDConnection();

        private String ALMServerURL = "";

        public TestPlan TestPlan;
        public TestLab TestLab;
        public Defect Defect;
        public Releases Releases;
        public Requirements Requirements;
        public Analysis Analysis;
        public Dashboard Dashboard;
        public TestResources TestResources;

        /// <summary>
        /// Login to ALM
        /// <para/>true if successfull
        /// </summary>
        /// <param name="URL">ALM URL this should end with QCBin</param>
        /// <param name="UserName">ALM Username</param>
        /// <param name="Password">ALM Password</param>
        /// <param name="Domain">ALM Domain name</param>
        /// <param name="Project">ALM Project Name</param>
        /// <returns>true if successfull</returns>
        public TDAPIOLELib.TDConnection LoginALM(String URL, String UserName, String Password, String Domain, String Project)
        {
            try
            {
                ALMServerURL = URL;

                //Check if OTA Client is registered
                //if (!IsOTARegistered())
                //{
                //    throw (new Exception("OTA Client is Not Registered on the machine"));
                //}

                tDConnection.InitConnectionEx(ALMServerURL);

                if (tDConnection.Equals(null))
                {
                    throw (new Exception("Unable to initiate connection with ALM Server"));
                }
                else
                {
                    tDConnection.Login(UserName, Password);

                    if (tDConnection.LoggedIn == false)
                    {
                        throw (new Exception("Unable to login to ALM"));
                    }
                }

                tDConnection.Connect(Domain, Project);

                if (!(tDConnection.ProjectConnected))
                {
                    throw (new Exception("Unable to Connect to the project"));
                }

                TestPlan = new TestPlan(tDConnection);
                TestLab = new TestLab(tDConnection);
                Defect = new Defect(tDConnection);
                Releases = new Releases(tDConnection);
                Requirements = new Requirements(tDConnection);
                Analysis = new Analysis(tDConnection);
                Dashboard = new Dashboard(tDConnection);
                TestResources = new TestResources(tDConnection);
                return tDConnection;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred while connecting to ALM : " + ex.Message.ToString());
                throw ex;
            }
        }

        /// <summary>
        /// Logout ALM. This should not be call except in the scenario where you want to switch projects in between executions. This will be autometically called at the end of executions.
        /// <para/>true if successfull
        /// </summary>
        /// <returns>true if successfull</returns>
        private Boolean LogoutALM()
        {
            try
            {
                if (tDConnection.ProjectConnected == true)
                {
                    tDConnection.DisconnectProject();
                }

                if (tDConnection.LoggedIn == true)
                {
                    tDConnection.Logout();
                }

                tDConnection.ReleaseConnection();
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                return false;
            }
                
            
        }

        /// <summary>
        /// Checks if the ALM project is still connected
        /// <para/>true if project is connected
        /// </summary>
        /// <returns>return true if successfull</returns>
        public Boolean IsProjectConnected()
        {
            try
            {
                return tDConnection.ProjectConnected;
            } catch
            {
                return false;
            }
            
        }

        /// <summary>
        /// Checks if user is still connected
        /// <para/>true if user is logged in
        /// </summary>
        /// <returns>return true if successfull</returns>
        public Boolean IsLoggedIn()
        {
            try
            {
                return tDConnection.LoggedIn;
            }
            catch
            {
                return false;
            }
            
        }


        /// <summary>
        /// Executes queries on ALM database. Only select queries can be executed using this function
        /// <para/> returns TDAPIOLELib.Recordset object
        /// </summary>
        /// <param name="QueryToExecute">String Query to execute</param>
        /// <returns>TDAPIOLELib.Recordset Object</returns>
        public TDAPIOLELib.Recordset ExecuteQuery(String QueryToExecute)
        {
            try
            {
                if (!(IsProjectConnected() && IsLoggedIn()))
                    throw (new Exception("No Connected to ALM. Please call LoginALM Function first"));

                return Utilities.ExecuteQuery(QueryToExecute, tDConnection);
               
            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message.ToString()));
            }
        }

        /// <summary>
        /// Logout at the end of executions
        /// </summary>
        ~ALM_CORE()
        {
            LogoutALM();
        }

        /// <summary>
        /// Verifies if the OTA is registered on the machine
        /// <para/> returns true is OTA is registered on the machine
        /// </summary>
        /// <returns></returns>
        public Boolean IsOTARegistered()
        {
            using (var classesRootKey = Microsoft.Win32.RegistryKey.OpenBaseKey(
                   Microsoft.Win32.RegistryHive.ClassesRoot, Microsoft.Win32.RegistryView.Default))
            {
                const string clsid = "{C5CBD7B2-490C-45f5-8C40-B8C3D108E6D7}";

                var clsIdKey = classesRootKey.OpenSubKey(@"Wow6432Node\CLSID\" + clsid) ??
                                classesRootKey.OpenSubKey(@"CLSID\" + clsid);

                if (clsIdKey != null)
                {
                    clsIdKey.Dispose();
                    return true;
                }

                return false;
            }
        }

        public Boolean Login(String URL, String UserName, String Password)
        {
            ALMServerURL = URL;

            tDConnection.InitConnectionEx(ALMServerURL);

            if (tDConnection.Equals(null))
            {
                throw (new Exception("Unable to initiate connection with ALM Server"));
            }
            else
            {
                tDConnection.Login(UserName, Password);

                if (tDConnection.LoggedIn == false)
                {
                    throw (new Exception("Unable to login to ALM"));
                }
            }

            return true;
        }

        public Boolean ConnectProject(String Domain, String Project)
        {
            tDConnection.Connect(Domain, Project);

            if (tDConnection.ProjectConnected)
                return true;

            throw (new Exception("Unable to connect to Project"));
        }

        public List<String> GetVisibleDomains()
        {
            List<String> VisibleDomains = new List<string>();

            foreach (String VD in tDConnection.VisibleDomains)
                VisibleDomains.Add(VD);

            return VisibleDomains;
        }

        public List<String> GetVisibleProjects(String DomainName)
        {
            List<String> VisibleDomains = new List<string>();

            foreach (String VD in tDConnection.VisibleProjects[DomainName])
                VisibleDomains.Add(VD);

            return VisibleDomains;
        }
    }
}
