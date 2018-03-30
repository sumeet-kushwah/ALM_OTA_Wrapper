using System;
using System.Collections.Generic;

namespace ALM_Wrapper
{
    public class Defect
    {
        private TDAPIOLELib.TDConnection tDConnection;

        //Constructor
        public Defect(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        /// <summary>
        /// creates a defect in ALM
        /// <para/>returns TDAPIOLELib.Bug Object
        /// </summary>
        /// <param name="defectDetails">Dictionary Object with Field name and field value strings</param>
        /// <returns>TDAPIOLELib.Bug Object</returns>
        public TDAPIOLELib.Bug Create(Dictionary<String, String> defectDetails)
        {
            TDAPIOLELib.BugFactory OBGFactory = tDConnection.BugFactory;
            TDAPIOLELib.Bug OBug = OBGFactory.AddItem(System.DBNull.Value);

            foreach (KeyValuePair<string, string> kvp in defectDetails)
            {
                OBug[kvp.Key.ToUpper()] = kvp.Value;
            }
            OBug.Post();

            return OBug;
        }

        /// <summary>
        /// Update defect field value
        /// <para/> returns True if successfull
        /// </summary>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="fieldName">Database Field name</param>
        /// <param name="newValue">New Field value</param>
        /// <param name="Post">Post the defect after updating valie</param>
        /// <returns>True if successfull</returns>
        public Boolean UpdateFieldValue(TDAPIOLELib.Bug bug, String fieldName, String newValue, Boolean Post = true)
        {
            bug[fieldName.ToUpper()] = newValue;
            if (Post)
                bug.Post();
            return true;
        }


        //public TDAPIOLELib.Bug GetObjectWithID(int ID)
        //{
        //    TDAPIOLELib.BugFactory OBugFactory = tDConnection.BugFactory as TDAPIOLELib.BugFactory;
        //    TDAPIOLELib.TDFilter OTDFilter = OBugFactory.Filter as TDAPIOLELib.TDFilter;
        //    TDAPIOLELib.List OBugList;

        //    TDAPIOLELib.Bug OBug;

        //    try
        //    {

        //        OTDFilter["BG_BUG_ID"] = Convert.ToString(ID);
        //        OBugList = OBugFactory.NewList(OTDFilter.Text);

        //        if (OBugList != null && OBugList.Count == 1)
        //        {
        //            OBug = OBugList[1];
        //            return OBug;
        //        }
        //        else
        //        {
        //            throw (new Exception("Unable to find test with ID : " + ID));
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}

        /// <summary>
        /// Get TDAPIOLELib.Bug with ID
        /// </summary>
        /// <param name="ID">Defect ID</param>
        /// <returns>TDAPIOLELib.Bug Object</returns>
        public TDAPIOLELib.Bug GetObjectWithID(int id)
        {
            TDAPIOLELib.BugFactory OBugFactory = tDConnection.BugFactory;
            return OBugFactory[id];
        }

        /// <summary>
        /// Counts Number of Defects in a project
        /// </summary>
        /// <returns>Count of Defects</returns>
        public int CountAll()
        {
            TDAPIOLELib.BugFactory OBugFactory = tDConnection.BugFactory;
            return OBugFactory.NewList("").Count;
        }

        /*public int CountAll()
        {
            TDAPIOLELib.Recordset ORecordSet = Utilities.ExecuteQuery("Select Count(*) from Bug", tDConnection);
            ORecordSet.First();
            return Convert.ToInt32(ORecordSet[0]);
        }*/

        /// <summary>
        /// Deletes a defect
        /// </summary>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <returns>true if successfull</returns>
        public Boolean Delete(TDAPIOLELib.Bug bug)
        {
            try
            {
                TDAPIOLELib.BugFactory OBGFactory = tDConnection.BugFactory;
                OBGFactory.RemoveItem(bug.ID);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// Add Attachments to defect
        /// </summary>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="attachmentsPath">List of attachment paths</param>
        /// <returns>True if Successfull</returns>
        public Boolean AddAttachment(TDAPIOLELib.Bug bug, List<String> attachmentsPath)
        {
            TDAPIOLELib.AttachmentFactory OAttachmentFactory;
            TDAPIOLELib.Attachment OAttachment;

            OAttachmentFactory = bug.Attachments;

            foreach (String AP in attachmentsPath)
            {
                if (System.IO.File.Exists(AP))
                {
                    OAttachment = OAttachmentFactory.AddItem(System.DBNull.Value);
                    OAttachment.FileName = AP;
                    OAttachment.Type = Convert.ToInt16(TDAPIOLELib.tagTDAPI_ATTACH_TYPE.TDATT_FILE);
                    OAttachment.Post();
                }
            }

            return true;
        }

        /// <summary>
        /// Downloads defect attachments
        /// </summary>
        /// <param name="bug"></param>
        /// <param name="attachmentDownloadPath"></param>
        /// <returns>True if Successfull</returns>
        public Boolean DownloadAttachments(TDAPIOLELib.Bug bug, String attachmentDownloadPath)
        {
            try
            {
                TDAPIOLELib.AttachmentFactory OAttachmentFactory;
                TDAPIOLELib.ExtendedStorage OExtendedStorage;

                if (bug.HasAttachment)
                {
                    if ((System.IO.Directory.Exists(attachmentDownloadPath)) == false)
                    {
                        throw (new Exception("Attachment download path does not exist"));
                    }

                    //System.IO.Directory.CreateDirectory(AttachmentDownloadPath + "\\" + OBug.ID.ToString());

                    OAttachmentFactory = bug.Attachments;

                    foreach (TDAPIOLELib.Attachment OAttachment in OAttachmentFactory.NewList(""))
                    {
                        OExtendedStorage = OAttachment.AttachmentStorage;
                        OExtendedStorage.ClientPath = attachmentDownloadPath;// + "\\" + OBug.ID.ToString();
                        OAttachment.Load(true, OAttachment.Name);
                    }
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Delete all attachments from defect
        /// </summary>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <returns>True if Successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.Bug bug)
        {
            TDAPIOLELib.AttachmentFactory OAttachmentFactory;

            OAttachmentFactory = bug.Attachments;
            TDAPIOLELib.List AttachmentsList = OAttachmentFactory.NewList("");

            foreach (TDAPIOLELib.Attachment OAttach in AttachmentsList)
            {
                OAttachmentFactory.RemoveItem(OAttach.ID);
            }

            return true;
        }

        /// <summary>
        /// Delete defect attachment by name
        /// </summary>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="attachmentName">Name of attachment</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAttachmentByName(TDAPIOLELib.Bug bug, String attachmentName)
        {
            TDAPIOLELib.AttachmentFactory OAttachmentFactory;

            OAttachmentFactory = bug.Attachments;
            TDAPIOLELib.List AttachmentsList = OAttachmentFactory.NewList("");

            foreach (TDAPIOLELib.Attachment OAttach in AttachmentsList)
            {
                if (OAttach.Name.EndsWith(attachmentName))
                {
                    OAttachmentFactory.RemoveItem(OAttach.ID);
                    break;
                }

            }

            return true;
        }

        /// <summary>
        /// Filter defects using the filetr string
        /// </summary>
        /// <param name="filterString">filter string, you can copy this from ALM</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List Filter(String filterString)
        {
            TDAPIOLELib.BugFactory OBugFactory = tDConnection.BugFactory as TDAPIOLELib.BugFactory;
            TDAPIOLELib.TDFilter OTDFilter = OBugFactory.Filter as TDAPIOLELib.TDFilter;
            TDAPIOLELib.List OTestList;

            try
            {
                OTDFilter.Text = filterString;
                OTestList = OBugFactory.NewList(OTDFilter.Text);
                return OTestList;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Link Defect to TestSet Test
        /// </summary>
        /// <param name="tSTest">TDAPIOLELib.TSTest Object</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if Successfull</returns>
        public Boolean LinkDefectToTestSetTest(TDAPIOLELib.TSTest tSTest, TDAPIOLELib.Bug bug, String comment = "")
        {
            return LinkDefectToEntities(tSTest, bug, comment);
        }

        /// <summary>
        /// Links defect to test step in test run
        /// </summary>
        /// <param name="step">TDAPIOLELib.Step Object</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if Successfull</returns>
        public Boolean LinkDefectToStep(TDAPIOLELib.Step step, TDAPIOLELib.Bug bug, String comment = "")
        {
            return LinkDefectToEntities(step, bug, comment);
        }

        /// <summary>
        /// Links Defect to Requirement
        /// </summary>
        /// <param name="requirement">TDAPIOLELib.Req Object</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if Successfull</returns>
        public Boolean LinkDefectToRequirement(TDAPIOLELib.Req requirement, TDAPIOLELib.Bug bug, String comment = "")
        {
            return LinkDefectToEntities(requirement, bug, comment);
        }

        /// <summary>
        /// Links Defect to Test
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if Successfull</returns>
        public Boolean LinkDefectToTest(TDAPIOLELib.Test test, TDAPIOLELib.Bug bug, String comment = "")
        {
            return LinkDefectToEntities(test, bug, comment);
        }

        /// <summary>
        /// Links Defect to TestSet
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if Successfull</returns>
        public Boolean LinkDefectToTestSet(TDAPIOLELib.TestSet testSet, TDAPIOLELib.Bug bug, String comment = "")
        {
            return LinkDefectToEntities(testSet, bug, comment);
        }

        /// <summary>
        /// Links Defect to Run
        /// </summary>
        /// <param name="run">TDAPIOLELib.Run Object</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if Successfull</returns>
        public Boolean LinkDefectToRun(TDAPIOLELib.Run run, TDAPIOLELib.Bug bug, String comment = "")
        {
            return LinkDefectToEntities(run, bug, comment);
        }

        /// <summary>
        /// Links defects to any ALM object 
        /// </summary>
        /// <param name="obj">Object class from ALM</param>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <param name="comment">Comment</param>
        /// <returns>True if successfull</returns>
        private Boolean LinkDefectToEntities(Object obj, TDAPIOLELib.Bug bug, String comment = "")
        {
            TDAPIOLELib.ILinkable linkable = obj as TDAPIOLELib.ILinkable;
            TDAPIOLELib.LinkFactory OLinkFactory = linkable.BugLinkFactory;
            TDAPIOLELib.Link OLink = OLinkFactory.AddItem(bug);

            if (comment.Length > 0)
            {
                OLink.Comment = comment;
            }
            OLink.Post();

            return true;
        }

        /// <summary>
        /// Gets the list of defects linked to Run
        /// </summary>
        /// <param name="run">TDAPIOLELib.Run Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetLinkedDefectsToRun(TDAPIOLELib.Run run)
        {
            return GetLinkedDefectsToEntities(run);
        }

        /// <summary>
        /// Gets the list of defects linked to test
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetLinkedDefectsToTest(TDAPIOLELib.Test test)
        {
            return GetLinkedDefectsToEntities(test);
        }

        /// <summary>
        /// Gets the list of defects linked to step
        /// </summary>
        /// <param name="step">TDAPIOLELib.Step Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetLinkedDefectsToStep(TDAPIOLELib.Step step)
        {
            return GetLinkedDefectsToEntities(step);
        }

        /// <summary>
        /// Gets the list of defects linked to testset
        /// </summary>
        /// <param name="testSet">TDAPIOLELib.TestSet Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetLinkedDefectsToTestSet(TDAPIOLELib.TestSet testSet)
        {
            return GetLinkedDefectsToEntities(testSet);
        }

        /// <summary>
        /// Gets the list of defects linked to testset test
        /// </summary>
        /// <param name="tSTest">TDAPIOLELib.TSTest Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetLinkedDefectsToTestSetTest(TDAPIOLELib.TSTest tSTest)
        {
            return GetLinkedDefectsToEntities(tSTest);
        }

        /// <summary>
        /// Gets the list of defects linked to requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        public TDAPIOLELib.List GetLinkedDefectsToRequirement(TDAPIOLELib.Req req)
        {
            return GetLinkedDefectsToEntities(req);
        }

        /// <summary>
        /// Gets the list of defects linked to ALM Entities
        /// </summary>
        /// <param name="Obj">Object Object</param>
        /// <returns>TDAPIOLELib.List Object</returns>
        private TDAPIOLELib.List GetLinkedDefectsToEntities(Object Obj)
        {
            TDAPIOLELib.ILinkable linkable = Obj as TDAPIOLELib.ILinkable;
            TDAPIOLELib.LinkFactory linkFactory = linkable.BugLinkFactory;

            TDAPIOLELib.List list = new TDAPIOLELib.List();
            foreach (TDAPIOLELib.Link link in linkFactory.NewList(""))
            {
                list.Add(link.TargetEntity as TDAPIOLELib.Bug); 
            }

            return list;
        }

        /// <summary>
        /// Get all defect database field values. This Functions uses RecordSet object. 
        /// <para/>You must have access to executing queries in ALM
        /// </summary>
        /// <param name="bug">TDAPIOLELib.Bug Object</param>
        /// <returns>true if successfull</returns>
        public TDAPIOLELib.Recordset GetAllDetails(TDAPIOLELib.Bug bug)
        {
            TDAPIOLELib.Recordset ORecordSet = Utilities.ExecuteQuery("Select * from Bug where BG_BUG_ID = " + bug.ID, tDConnection);
            ORecordSet.First();
            return ORecordSet;
        }

    }
}
