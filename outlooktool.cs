using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace outlookaddin
{
    public class OutlookHelperTool
    {

        private const String subjectprefix = "[Crypted]";
        private const String TagEncryptedOpen = "{encryption}";
        private const String TagEncryptedClose = "{/encryption}";
        private const String TagInitVectorOpen = "{initvector}";
        private const String TagInitVectorClose = "{/initvector}";
        private static String initvectorPattern = String.Format(@"{0}(.*?){1}", TagInitVectorOpen, TagInitVectorClose);
        private static String encryptedStringPattern = String.Format(@"{0}(.*?){1}", TagEncryptedOpen, TagEncryptedClose);

        public static String getinitVector(String text)
        {
            Match match = Regex.Match(text, initvectorPattern);

            String result = match.Groups[1].Value;

            return result;
        }



        public static String getSubject()
        {
            return getMailItem().Subject;
        }

        public static void setCryptedSubject()
        {
            if (getSubject().Contains(subjectprefix))
            {
                return;
            }
            getMailItem().Subject = subjectprefix + getSubject();
        }
        public static void cutCryptedSubject()
        {
            if (getSubject().Contains(subjectprefix))
            {
                return;
            }
            String tempSubject = getMailItem().Subject;
            getMailItem().Subject = tempSubject.Substring(subjectprefix.Length);
        }

        public static void cutCryptedSubject(MailItem mailitem)
        {

            String tempSubject = mailitem.Subject;
            mailitem.Subject = tempSubject.Substring(subjectprefix.Length);
        }


        private static MailItem getMailItem()
        {
            dynamic wind = ThisAddIn.AppObj.ActiveWindow();



            if (wind is Microsoft.Office.Interop.Outlook.Inspector)
            {
                Microsoft.Office.Interop.Outlook.Inspector inspector = (Microsoft.Office.Interop.Outlook.Inspector)wind;

                if (inspector.CurrentItem is Microsoft.Office.Interop.Outlook.MailItem)
                {
                    Microsoft.Office.Interop.Outlook.MailItem mailitem = (Microsoft.Office.Interop.Outlook.MailItem)inspector.CurrentItem;
                    return mailitem;
                }
            }

            if (wind is Microsoft.Office.Interop.Outlook.Explorer)
            {
                Microsoft.Office.Interop.Outlook.Explorer explorer = (Microsoft.Office.Interop.Outlook.Explorer)wind;

                Microsoft.Office.Interop.Outlook.Selection selectedItems = explorer.Selection;
                if (selectedItems.Count != 1)
                {
                    return null;
                }

                if (selectedItems[1] is Microsoft.Office.Interop.Outlook.MailItem)
                {
                    Microsoft.Office.Interop.Outlook.MailItem mailitem = selectedItems[1];
                    return mailitem;
                }
            }

            return null;
        }

        public static String getMailText(bool encrypted)
        {
            Microsoft.Office.Interop.Outlook.MailItem mailitem = getMailItem();
            switch (mailitem.BodyFormat)
            {
                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatUnspecified:
                    return manageEncryptionRead(mailitem.Body, encrypted);

                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain:
                    return manageEncryptionRead(mailitem.Body, encrypted);

                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML:
                    return manageEncryptionRead(mailitem.HTMLBody, encrypted);

                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatRichText:
                    return manageEncryptionRead(mailitem.Body, encrypted);
                default:
                    break;
            }

            return null;

        }

        public static String getMailText(MailItem mailitem, bool encrypted)
        {
            switch (mailitem.BodyFormat)
            {
                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatUnspecified:
                    return manageEncryptionRead(mailitem.Body, encrypted);

                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatPlain:
                    return manageEncryptionRead(mailitem.Body, encrypted);

                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML:
                    return manageEncryptionRead(mailitem.HTMLBody, encrypted);

                case Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatRichText:
                    return manageEncryptionRead(mailitem.Body, encrypted);
                default:
                    break;
            }

            return null;

        }


        public static void InspectorWrapper(object Item)
        {
            if (!(Item is Inspector))
            {
                return;
            }
            Inspector inspector = (Inspector)Item;
            if (!(inspector.CurrentItem is MailItem))
            {
                return;
            }


            MailItem mailitem = (MailItem)inspector.CurrentItem;

            mailitem.Read += new ItemEvents_10_ReadEventHandler(mailItem_Read);



         


        }

        private static void mailItem_Read()
        {
            MailItem mailitem = getMailItem();
            if (mailitem.Subject.Contains(subjectprefix) && !mailitem.Subject.StartsWith("FW: ") && !mailitem.Subject.StartsWith("AW: ") && 
                (mailitem.Body.Contains(TagEncryptedClose) && mailitem.Body.Contains(TagEncryptedOpen) && mailitem.Body.Contains(TagInitVectorClose) && mailitem.Body.Contains(TagInitVectorOpen)) 
                )
                {
                DecryptForm decryptform = new DecryptForm(mailitem);
            }
        }

        public static void ItemLoadEventHandler(object item)
        {

            if (!(item is MailItem))
            {
                return;
            }


            MailItem mailitem = (MailItem)item;

         
            if(mailitem == null)
            {
                return;
            }

            if (!mailitem.Subject.Contains(subjectprefix))
            {
                return;
            }
            MessageBox.Show("Encrypted E-Mail");





        }

        private static String manageEncryptionRead(String text, bool encrypted)
        {
            String result = text;
            if (encrypted)
            {
                Match match = Regex.Match(text, encryptedStringPattern);

                result = match.Groups[1].Value;


            }

            return result;
        }

        public static String getPreparedMessage(String crypted, String initVector)
        {
            return prepareString(crypted, initVector);
        }

        public static void setMailText(String text)
        {
            MailItem mailitem = getMailItem();
            mailitem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
            mailitem.HTMLBody = text;
        }

        public static void setMailText(MailItem mailitem, String text)
        {
            mailitem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
            mailitem.HTMLBody = text;
        }

        private static String prepareString(String encrpytedString, String initvector)
        {
            String formatedString = String.Format("<html>" +
                "<body>" +
                "<h3><b>This E-Mail is encrypted</b></h3>\n" +
                "<hr>" +
                "<p>{1}{0}{2}</p>\n" +
                "<p>{3}{4}{5}</p>" +
                "</body>" +
                "</html>", encrpytedString, TagEncryptedOpen, TagEncryptedClose, TagInitVectorOpen, initvector, TagInitVectorClose);

            return formatedString;
        }

    }
}
