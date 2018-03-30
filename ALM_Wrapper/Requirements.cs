using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALM_Wrapper
{
    public class Requirements
    {
        private TDAPIOLELib.TDConnection tDConnection;

        public Requirements(TDAPIOLELib.TDConnection tDConnection)
        {
            this.tDConnection = tDConnection;
        }

        /// <summary>
        /// Rename a requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="newName">New requirement name</param>
        /// <returns>True if Successfull</returns>
        public Boolean Rename(TDAPIOLELib.Req req, String newName)
        {
            req["RQ_REQ_NAME"] = newName;
            req.Post();
            return true;
        }

        /// <summary>
        /// Creates a new requirement
        /// </summary>
        /// <param name="requirementDetails">dictionary object with requirement database field names and values</param>
        /// <param name="folderPath"></param>
        /// <param name="reqType"></param>
        /// <returns></returns>
        public TDAPIOLELib.Req Create(Dictionary<String, String> requirementDetails, String folderPath, TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES reqType = TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES.REQ_TYPE_UNDEFINED)
        {
            TDAPIOLELib.Req req = GetReqByPath(folderPath);
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;

            TDAPIOLELib.Req newReq = reqFactory.AddItem(System.DBNull.Value);

            newReq.ParentId = req.ID;
            newReq["RQ_TYPE_ID"] = reqType;

            foreach (KeyValuePair<String, String> kvp in requirementDetails)
            {
                newReq[kvp.Key] = kvp.Value;
            }

            newReq.Post();

            return newReq;
        }

        /// <summary>
        /// Creates a requirement Folder
        /// </summary>
        /// <param name="folderName">Requirement folder name</param>
        /// <param name="folderPath">Requirement folder path</param>
        /// <returns>TDAPIOLELib.Req Object</returns>
        public TDAPIOLELib.Req CreateFolder(String folderName, String folderPath)
        {
            TDAPIOLELib.Req req = GetReqByPath(folderPath);
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;

            TDAPIOLELib.Req newReq = reqFactory.AddItem(System.DBNull.Value);

            newReq.ParentId = req.ID;
            newReq["RQ_TYPE_ID"] = TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES.REQ_TYPE_FOLDER;

            newReq.Name = folderName;

            newReq.Post();

            return newReq;
        }

        /// <summary>
        /// Updates a requirement field value
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="fieldName"></param>
        /// <param name="newValue"></param>
        /// <param name="Post"></param>
        /// <returns></returns>
        public Boolean UpdateFieldValue(TDAPIOLELib.Req req, String fieldName, String newValue, Boolean Post = true)
        {
            req[fieldName.ToUpper()] = newValue;
            if (Post)
                req.Post();

            return true;
        }

        /// <summary>
        /// Finds filtered child requirements
        /// </summary>
        /// <param name="req">Parent requirement</param>
        /// <param name="reqType">Filter type of requirement</param>
        /// <returns>List of requirements</returns>
        public TDAPIOLELib.List GetChildRequirements(TDAPIOLELib.Req req, TDAPIOLELib.TDAPI_PREDEFINED_REQ_TYPES reqType)
        {
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;
            TDAPIOLELib.TDFilter tDFilter = reqFactory.Filter;

            tDFilter["RQ_TYPE_ID"] = reqType.ToString();

            return reqFactory.GetFilteredChildrenList(req.ID, tDFilter);
        }

        /// <summary>
        /// Finds child requirements
        /// </summary>
        /// <param name="req">parent requirement</param>
        /// <returns>List of child requirements</returns>
        public TDAPIOLELib.List GetChildRequirements(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;
            return reqFactory.GetChildrenList(req.ID);
        }

        /// <summary>
        /// Delete a requirement
        /// </summary>
        /// <param name="req">Requirement to be deleted</param>
        /// <returns>True if successfull</returns>
        public Boolean Delete(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;
            reqFactory.RemoveItem(req.ID);
            return true;
        }

        /// <summary>
        /// Add Attachment to Requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="attachmentsPath">List of Attachments path</param>
        /// <returns>True if Successfull</returns>
        public Boolean AddAttachements(TDAPIOLELib.Req req, List<String> attachmentsPath)
        {
            return Utilities.AddAttachment(req.Attachments, attachmentsPath);
        }

        /// <summary>
        /// Delete all attachments for a requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>True if Successfull</returns>
        public Boolean DeleteAllAttachments(TDAPIOLELib.Req req)
        {
            return Utilities.DeleteAllAttachments(req.Attachments);
        }

        /// <summary>
        /// Deletes a requirement attachment by name
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="attachmentName">Attachment name to be deleted</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteAttachementByName(TDAPIOLELib.Req req, String attachmentName)
        {
            return Utilities.DeleteAttachmentByName(req.Attachments, attachmentName);
        }

        /// <summary>
        /// downloads requirement attachments
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="attachmentDownloadPath">attachment download path</param>
        /// <returns></returns>
        public Boolean DownloadAttachements(TDAPIOLELib.Req req, String attachmentDownloadPath)
        {
            return Utilities.DownloadAttachments(req.Attachments, attachmentDownloadPath);
        }


        /// <summary>
        /// Get TDAPIOLELib.Req object from path
        /// </summary>
        /// <param name="folderPath">Requirement folder path</param>
        /// <returns>TDAPIOLELib.Req Object</returns>
        public TDAPIOLELib.Req GetReqByPath(String folderPath)
        {
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;
            TDAPIOLELib.Req parentReq = reqFactory[0];
            TDAPIOLELib.List reqList;

            foreach (String Req in folderPath.Split('\\'))
            {
                if (Req.ToUpper() != "REQUIREMENTS")
                {
                    reqList = reqFactory.Find(parentReq.ID, "RQ_REQ_NAME", Req, Convert.ToInt16(TDAPIOLELib.TDAPI_REQMODE.TDREQMODE_FIND_EXACT));

                    parentReq = null;
                    String FirstItemOfList = reqList[1] as String;

                    parentReq = reqFactory[FirstItemOfList.Split(',')[0]];
                }
            }

            return parentReq;
        }

        /// <summary>
        /// Add test to requirement coverage
        /// </summary>
        /// <param name="test">TDAPIOLELib.Test Object</param>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>True if successfull</returns>
        public Boolean AddTestToRequirementsCoverage(TDAPIOLELib.Test test, TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ICoverableReq coverable = req as TDAPIOLELib.ICoverableReq;
            coverable.AddTestToCoverage(test.ID);
            return true;
        }

        /// <summary>
        /// Add a test plan folder to requiremnt coverage
        /// </summary>
        /// <param name="testFolderID"></param>
        /// <param name="req"></param>
        /// <returns></returns>
        public Boolean AddTestPlanFolderToRequirementsCoverage(int testFolderID, TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ICoverableReq coverable = req as TDAPIOLELib.ICoverableReq;
            coverable.AddSubjectToCoverage(testFolderID, "");
            return true;
        }

        /// <summary>
        /// Set target releases for a requirements.
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="listofReleaseIDs">TDAPIOLELib.List of Release ID's</param>
        /// <returns>True if successfull</returns>
        public Boolean SetTargetReleases(TDAPIOLELib.Req req, TDAPIOLELib.List listofReleaseIDs)
        {
            req["RQ_TARGET_REL"] = listofReleaseIDs;
            req.Post();
            return true;
        }

        /// <summary>
        /// Find target releases for requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>List of Target releases</returns>
        public TDAPIOLELib.List GetTargetReleases(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.List listofReleaseIDs = req["RQ_TARGET_REL"];
            return listofReleaseIDs;
        }

        /// <summary>
        /// Delete target releases
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>True if Successfull</returns>
        public Boolean DeleteTargetReleases(TDAPIOLELib.Req req)
        {
            req["RQ_TARGET_REL"] = "";
            req.Post();
            return true;
        }

        /// <summary>
        /// Set target cycles for requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Obejct</param>
        /// <param name="listofCycleIDs">List of Cycle ID's</param>
        /// <returns>True if successfull</returns>
        public Boolean SetTargetCycles(TDAPIOLELib.Req req, TDAPIOLELib.List listofCycleIDs)
        {
            req["RQ_TARGET_RCYC"] = listofCycleIDs;
            req.Post();
            return true;
        }

        /// <summary>
        /// Delete target cycles
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>True if successfull</returns>
        public Boolean DeleteTargetCycles(TDAPIOLELib.Req req)
        {
            req["RQ_TARGET_RCYC"] = "";
            req.Post();
            return true;
        }

        /// <summary>
        /// Get target cycles
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>True if successfull</returns>
        public TDAPIOLELib.List GetTargetCycles(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.List listofCycleIDs = req["RQ_TARGET_RCYC"];
            return listofCycleIDs;
        }

        /// <summary>
        /// Gets parent requirement
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>Parent requirement TDAPIOLELib.Req Object</returns>
        public TDAPIOLELib.Req GetParentRequirement(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ReqFactory reqFactory = tDConnection.ReqFactory;
            return reqFactory[req.ID];
        }

        /// <summary>
        /// Adds Requirement tracebility TraceFrom
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="traceFromReq">TDAPIOLELib.Req Object</param>
        /// <returns>TDAPIOLELib.Trace Object</returns>
        public TDAPIOLELib.Trace AddRequirementTraceability_TraceFrom(TDAPIOLELib.Req req, TDAPIOLELib.Req traceFromReq)
        {
            TDAPIOLELib.ReqTraceFactory reqTraceFactory = req.ReqTraceFactory[(int)TDAPIOLELib.tagTDAPI_TRACE_DIRECTION.TDOLE_TRACED_FROM];
            return reqTraceFactory.AddItem(traceFromReq.ID);
        }

        /// <summary>
        /// Adds Requirement tracebility TraceTo
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="traceToReq">TDAPIOLELib.Req Object</param>
        /// <returns>TDAPIOLELib.Trace Object</returns>
        public TDAPIOLELib.Trace AddRequirementTraceability_TraceTo(TDAPIOLELib.Req req, TDAPIOLELib.Req traceToReq)
        {
            TDAPIOLELib.ReqTraceFactory reqTraceFactory = req.ReqTraceFactory[(int)TDAPIOLELib.tagTDAPI_TRACE_DIRECTION.TDOLE_TRACED_TO];
            return reqTraceFactory.AddItem(traceToReq.ID);
        }

        /// <summary>
        /// Delete all requirements from Tracefrom
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="traceFromReq">TDAPIOLELib.Trace Object</param>
        /// <returns>True if Successfull</returns>
        public Boolean DeleteRequirementTraceability_TraceFrom(TDAPIOLELib.Req req, TDAPIOLELib.Trace traceFromReq)
        {
            TDAPIOLELib.ReqTraceFactory reqTraceFactory = req.ReqTraceFactory[(int)TDAPIOLELib.tagTDAPI_TRACE_DIRECTION.TDOLE_TRACED_FROM];
            reqTraceFactory.RemoveItem(traceFromReq.ID);

            return true;
        }

        /// <summary>
        /// Delete all requirements from TraceTo
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <param name="traceToReq">TDAPIOLELib.Trace Object</param>
        /// <returns>True if Successfull</returns>
        public Boolean DeleteRequirementTraceability_TraceTo(TDAPIOLELib.Req req, TDAPIOLELib.Trace traceToReq)
        {
            TDAPIOLELib.ReqTraceFactory reqTraceFactory = req.ReqTraceFactory[(int)TDAPIOLELib.tagTDAPI_TRACE_DIRECTION.TDOLE_TRACED_TO];
            reqTraceFactory.RemoveItem(traceToReq.ID);

            return true;
        }

        /// <summary>
        /// Get all requirements from Tracefrom
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>List of TDAPIOLELib.Trace Object</returns>
        public TDAPIOLELib.List GetRequirementTraceability_TraceFrom(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ReqTraceFactory reqTraceFactory = req.ReqTraceFactory[(int)TDAPIOLELib.tagTDAPI_TRACE_DIRECTION.TDOLE_TRACED_FROM];
            return reqTraceFactory.NewList("");
        }

        /// <summary>
        /// Get all requirements from Traceto
        /// </summary>
        /// <param name="req">TDAPIOLELib.Req Object</param>
        /// <returns>List of TDAPIOLELib.Trace Object</returns>
        public TDAPIOLELib.List GetRequirementTraceability_TraceTo(TDAPIOLELib.Req req)
        {
            TDAPIOLELib.ReqTraceFactory reqTraceFactory = req.ReqTraceFactory[(int)TDAPIOLELib.tagTDAPI_TRACE_DIRECTION.TDOLE_TRACED_TO];
            return reqTraceFactory.NewList("");
        }
    }
}
