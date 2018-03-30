using System;
using System.Collections.Generic;

namespace ALM_Wrapper
{
    public enum Fieldtype { String = 0, DateTime = 1 };

    public class Cycle
    {
        private TDAPIOLELib.TDConnection tDConnection;

        public Cycle(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        /// <summary>
        /// Creates a cycle inside release
        /// <Para/>returns TDAPIOLELib.Cycle Object
        /// </summary>
        /// <param name="cycleName">Name of Cycle</param>
        /// <param name="startDate">Cycle Start Date</param>
        /// <param name="endDate">Cycle End Date</param>
        /// <param name="release">Release under which Cycle Should be created</param>
        /// <param name="post">Post this Cycle</param>
        /// <returns>TDAPIOLELib.Cycle Object</returns>
        public TDAPIOLELib.Cycle Create(String cycleName, DateTime startDate, DateTime endDate, TDAPIOLELib.Release release, Boolean post = true)
        {
            TDAPIOLELib.CycleFactory cycleFactory = release.CycleFactory;
            TDAPIOLELib.Cycle cycle = cycleFactory.AddItem(System.DBNull.Value);

            cycle.Name = cycleName;
            cycle.EndDate = endDate;
            cycle.StartDate = startDate;
            
            if (post)
                cycle.Post();

            return cycle;
        }

        /// <summary>
        /// Update Cycle Field value
        /// </summary>
        /// <param name="cycle"></param>
        /// <param name="fieldName"></param>
        /// <param name="newValue"></param>
        /// <param name="Post"></param>
        /// <returns></returns>
        public Boolean UpdateFieldValue(TDAPIOLELib.Cycle cycle, String fieldName, String newValue, Boolean Post = true)
        {
           
            cycle[fieldName.ToUpper()] = newValue;
           
            if (Post)
                cycle.Post();

            return true;
        }
        
        /// <summary>
        /// Delete Cycle
        /// <para/> True if successfull
        /// </summary>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <returns>True if successfull</returns>
        public Boolean Delete(TDAPIOLELib.Cycle cycle)
        {
            TDAPIOLELib.CycleFactory cycleFactory = tDConnection.CycleFactory;
            cycleFactory.RemoveItem(cycle.ID);
            return true;
        }

        /// <summary>
        /// Finds Release for a Cycle
        /// <para>returns TDAPIOLELib.Release Object</para>
        /// </summary>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <returns>TDAPIOLELib.Release Object</returns>
        public TDAPIOLELib.Release GetRelease(TDAPIOLELib.Cycle cycle)
        {
            return cycle.Parent;
        }

        /// <summary>
        /// Get All Cycle Field Values. This Function uses TDAPIOLELib.Recordset, So You must have access in ALM to run queries
        /// <para>returns TDAPIOLELib.Recordset Object</para>
        /// </summary>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <returns>TDAPIOLELib.Recordset Object</returns>
        public TDAPIOLELib.Recordset GetAllDetails(TDAPIOLELib.Cycle cycle)
        {
            return Utilities.ExecuteQuery("Select * from RELEASE_CYCLES where RCYC_ID = " + cycle.ID, tDConnection);
        }

        /// <summary>
        /// Download Cycle Attachments
        /// <para>returns true if successfull</para>
        /// </summary>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <param name="attachmentDownloadPath">attachments download path</param>
        /// <returns>True if successfull</returns>
        public Boolean DownloadAttachments(TDAPIOLELib.Cycle cycle, String attachmentDownloadPath)
        {
            return Utilities.DownloadAttachments(cycle.Attachments, attachmentDownloadPath);
        }

        /// <summary>
        /// Add attachments to Cycle
        /// <para>returns true if successfull</para>
        /// </summary>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <param name="attachmentsPath">List of attachment path</param>
        /// <returns>True if successfull</returns>
        public Boolean AddAttachment(TDAPIOLELib.Cycle cycle, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(cycle.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Delete all Cycle attachments
        /// <para>returns true if successfull</para>
        /// </summary>
        /// <param name="cycle">TDAPIOLELib.Cycle Object</param>
        /// <returns>true if successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.Cycle cycle)
        {
            return Utilities.DeleteAllAttachments(cycle.Attachments);
        }

        /// <summary>
        /// Delete attachment by name
        /// <para>returns true if successfull</para>
        /// </summary>
        /// <param name="cycle"></param>
        /// <param name="attachmentName"></param>
        /// <returns>true if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.Cycle cycle, String attachmentName)
        {
            return Utilities.DeleteAttachmentByName(cycle.Attachments, attachmentName);
        }
    }
}
