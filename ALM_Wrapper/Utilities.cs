using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALM_Wrapper
{
    static class Utilities
    {
        /// <summary>
        /// Executes query on ALM Database
        /// </summary>
        /// <param name="QueryToExecute">Query to execute</param>
        /// <param name="tDConnection">TDAPIOLELib.TDConnection Object with active ALM Connection</param>
        /// <returns>TDAPIOLELib.Recordset Object</returns>
        public static TDAPIOLELib.Recordset ExecuteQuery(String QueryToExecute, TDAPIOLELib.TDConnection tDConnection)
        {
            try
            {
                if (!(QueryToExecute.Trim().ToUpper().StartsWith("SELECT")))
                    throw (new Exception("Only Select Query can be executed using this funtion"));

                TDAPIOLELib.Command OCommand = (TDAPIOLELib.Command)tDConnection.Command;
                OCommand.CommandText = QueryToExecute;
                return (TDAPIOLELib.Recordset)OCommand.Execute();
            }
            catch (Exception ex)
            {
                throw (new Exception(ex.Message.ToString()));
            }
        }

        /// <summary>
        /// Add attachment using attachment factory object
        /// </summary>
        /// <param name="attachmentFactory">AttachmentFactory Object</param>
        /// <param name="attachmentsPath">List of paths</param>
        /// <returns>True if Successfull</returns>
        public static Boolean AddAttachment(TDAPIOLELib.AttachmentFactory attachmentFactory, List<String> attachmentsPath)
        {
            TDAPIOLELib.Attachment OAttachment;
            foreach (String AP in attachmentsPath)
            {
                if (System.IO.File.Exists(AP))
                {
                    OAttachment = attachmentFactory.AddItem(System.DBNull.Value);
                    OAttachment.FileName = AP;
                    OAttachment.Type = Convert.ToInt16(TDAPIOLELib.tagTDAPI_ATTACH_TYPE.TDATT_FILE);
                    OAttachment.Post();
                }
            }

            return true;
        }

        /// <summary>
        /// Delete all attachments using attachmentfactory Object
        /// </summary>
        /// <param name="attachmentFactory">TDAPIOLELib.AttachmentFactory Object</param>
        /// <returns>True if successfull</returns>
        public static Boolean DeleteAllAttachments(TDAPIOLELib.AttachmentFactory attachmentFactory)
        {
            TDAPIOLELib.List AttachmentsList = attachmentFactory.NewList("");
            foreach (TDAPIOLELib.Attachment OAttach in AttachmentsList)
            {
                attachmentFactory.RemoveItem(OAttach.ID);
            }

            return true;
        }

        /// <summary>
        /// Delete Attachments by name
        /// </summary>
        /// <param name="attachmentFactory">TDAPIOLELib.AttachmentFactory Object</param>
        /// <param name="attachmentName">name of the attachment to be deleted</param>
        /// <returns>Return true if successfull</returns>
        public static Boolean DeleteAttachmentByName(TDAPIOLELib.AttachmentFactory attachmentFactory, String attachmentName)
        {
            TDAPIOLELib.List AttachmentsList = attachmentFactory.NewList("");
            foreach (TDAPIOLELib.Attachment OAttach in AttachmentsList)
            {
                if (OAttach.Name.EndsWith(attachmentName))
                {
                    attachmentFactory.RemoveItem(OAttach.ID);
                    break;
                }

            }
            return true;
        }

        /// <summary>
        /// Download all attachments using attachmentfactory objecy
        /// </summary>
        /// <param name="attachmentFactory">TDAPIOLELib.AttachmentFactory Object</param>
        /// <param name="attachmentDownloadPath">Path to download attachments</param>
        /// <returns></returns>
        public static Boolean DownloadAttachments(TDAPIOLELib.AttachmentFactory attachmentFactory, String attachmentDownloadPath)
        {
            try
            {
                TDAPIOLELib.ExtendedStorage OExtendedStorage;

                if (attachmentFactory.NewList("").Count > 0)
                {
                    if ((System.IO.Directory.Exists(attachmentDownloadPath)) == false)
                    {
                        throw (new Exception("Attachment download path does not exist"));
                    }

                    foreach (TDAPIOLELib.Attachment OAttachment in attachmentFactory.NewList(""))
                    {
                        OExtendedStorage = OAttachment.AttachmentStorage;
                        OExtendedStorage.ClientPath = attachmentDownloadPath;
                        OAttachment.Load(true, OAttachment.Name);
                    }
                    return true;
                }
                else
                {
                    throw (new Exception("Attachments not Found"));
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

    }
}
