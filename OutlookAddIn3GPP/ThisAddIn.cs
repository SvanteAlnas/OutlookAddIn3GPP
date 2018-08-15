using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Win32;
using System.Net;

namespace OutlookAddIn3GPP
{
    public partial class ThisAddIn
    {
        const string VER = "1.0";
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        const string REGKEY = "HKEY_CURRENT_USER\\Software\\Alnas\\OutlookAttachment\\";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            items = inbox.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }
        void items_ItemAdd(object Item)
        {
            try
            {
                Outlook.MailItem mail = (Outlook.MailItem)Item;
                if (Item != null)
                {
                    if (mail.MessageClass == "IPM.Note")
                    {
                        if ((mail.Attachments != null) && (mail.Attachments.Count > 0))
                        {
                            object o = null;

                            // 0 = Keep original, 1 = Delete original attachment
                            int Delete = 1;
                            o = Registry.GetValue(REGKEY, "DeleteOriginal", 1);
                            if (o != null)
                                Delete = (int)o;

                            // 0 = No info at all, 1 = Only at Errors, 2 = Always if downloaded content, 99 = Always
                            int Verbose = 1;
                            o = Registry.GetValue(REGKEY, "Verbose", 1);
                            if (o != null)
                                Verbose = (int)o;

                            bool Found = false;
                            bool Err = false;
                            string Msg = "3GPP Attachment replacements, Ver: " + VER + Environment.NewLine;
                            for (int i = mail.Attachments.Count; i > 0; i--)
                            {
                                Msg += Environment.NewLine + "Attachment# " + i.ToString() + Environment.NewLine;
                                try
                                {
                                    if ((mail.Attachments[i].FileName != null) &&
                                        (mail.Attachments[i].FileName.Length > 11) &&
                                        (mail.Attachments[i].FileName.Substring(0, 11) == "Attachment ") &&
                                        (mail.Attachments[i].FileName.IndexOf(".txt") > 0))
                                    {
                                        //                                    string s = mail.Attachments[i].GetTemporaryFilePath();
                                        string s = System.IO.Path.GetTempPath() + System.IO.Path.GetRandomFileName();
                                        mail.Attachments[i].SaveAsFile(s);
                                        string[] lines = System.IO.File.ReadAllLines(s);
                                        if ((lines != null) && (lines.Length > 3) && (lines.Length < 6) && (lines[0].IndexOf("Attachment:") > -1) && (lines[1].Length > 2) && (lines[3].IndexOf("://list.etsi.org/") > -1))
                                        {
                                            Found = true;
                                            string Name = lines[1].Substring(1, lines[1].Length - 2);
                                            WebClient W = new WebClient();
                                            //                                        string s2 = System.IO.Path.GetTempPath() + System.IO.Path.GetRandomFileName();
                                            Msg += "Name: '" + mail.Attachments[i].FileName + "'" + Environment.NewLine + "Downloading attachment: '" + lines[3] + "', ";
                                            W.DownloadFile(lines[3], Name);
                                            long L = new System.IO.FileInfo(Name).Length;
                                            Msg += "Done" + Environment.NewLine + "Attachment downloded, size: " + L.ToString();
                                            if ((W.ResponseHeaders != null) && (W.ResponseHeaders["Content-Type"] != null))
                                                Msg += ", Content-Type: '" + W.ResponseHeaders["Content-Type"] + "'";
                                            if (Delete == 1)
                                                Msg += Environment.NewLine + "Replacing the attachment with the downloaded file, ";
                                            else
                                                Msg += Environment.NewLine + "Adding attachment with the downloaded file, ";
                                            mail.Attachments.Add(Name, Outlook.OlAttachmentType.olByValue, 1, Name);
                                            if (Delete == 1)
                                                mail.Attachments[i].Delete();
                                            Msg += "Done" + Environment.NewLine;
                                            if (Delete == 1)
                                            {
                                                Msg += "Content of the original attachment: " + Environment.NewLine;
                                                for (int j = 0; j < lines.Length; j++)
                                                    Msg += "\t" + "Line: " + j.ToString() + ": '" + lines[j] + "'" + Environment.NewLine;
                                            }
                                            System.IO.File.Delete(s);
                                            System.IO.File.Delete(Name);
                                        }
                                    }
                                    else
                                    {
                                        Msg += "Attachment ignored (non 3GPP document link)." + Environment.NewLine;
                                    }
                                }
                                catch (Exception Ex)
                                {
                                    Msg += "Error: '" + Ex.Message + "'" + Environment.NewLine;
                                    Err = true;
                                }
                            }
                            Msg += Environment.NewLine + Environment.NewLine + "------------------------------------------------" + Environment.NewLine + "3GPP Attachment replacement add-in, Svante Alnås" + Environment.NewLine + @"http://www.alnas.com/3GPPAttachment";
                            if (((Verbose == 1) && Err) || ((Verbose == 2) && Found) || (Verbose == 99))
                            {
                                try
                                {
                                    string Tmp = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".txt";
                                    System.IO.File.WriteAllText(Tmp, Msg);
                                    mail.Attachments.Add(Tmp, Outlook.OlAttachmentType.olByValue, 1, "3GPPAttachmentAddIn.txt");
                                    System.IO.File.Delete(Tmp);
                                }
                                catch (Exception) { }
                            }
                        }
                    }
                }
            }
            catch (Exception Ex)
            { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
