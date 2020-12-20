using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SendtoSOMEAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void HomeTabBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("Do you accept send selected mail to Cyber team?", "Send to Cyber", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
                if (dialog == DialogResult.Yes)
                {
                    Outlook._Application _Application = new Outlook.Application();
                    Outlook.MailItem selectedMail = new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection[1];
                    Outlook.MailItem mail = (Outlook.MailItem)_Application.CreateItem(Outlook.OlItemType.olMailItem);
                    RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Office\Outlook\Addins\SendtoSOMEAddIn");
                    string Receipient = key.GetValue("ReceipientAddress").ToString();
                    mail.To = Receipient;
                    mail.Subject = "Suspicious Mail";
                    mail.Body = "This mail seems like suspicious please investigate it. Regards";
                    mail.Importance = Outlook.OlImportance.olImportanceNormal;
                    mail.Attachments.Add(selectedMail, Outlook.OlAttachmentType.olByValue, 1, selectedMail.Subject);
                    ((Outlook._MailItem)mail).Send();
                    MessageBox.Show("Mail sent.", "succes", MessageBoxButtons.OK,MessageBoxIcon.Information);
                    key.Close();
                }
            }
            catch
            {
                MessageBox.Show("Please select a mail", "Send to Cyber",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
        }

        private void ReadTabBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("Do you accept send selected mail to Cyber team?", "Send to Cyber", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialog == DialogResult.Yes)
                {
                    Outlook._Application _Application = new Outlook.Application();
                    Outlook.MailItem selectedMail = new Microsoft.Office.Interop.Outlook.Application().ActiveExplorer().Selection[1];
                    Outlook.MailItem mail = (Outlook.MailItem)_Application.CreateItem(Outlook.OlItemType.olMailItem);
                    RegistryKey key = Registry.LocalMachine.OpenSubKey(@"Software\Microsoft\Office\Outlook\Addins\SendtoSOMEAddIn");
                    string Receipient = key.GetValue("ReceipientAddress").ToString();
                    mail.To = Receipient;
                    mail.Subject = "Suspicious Mail";
                    mail.Body = "This mail seems like suspicious please investigate it. Regards";
                    mail.Importance = Outlook.OlImportance.olImportanceNormal;
                    mail.Attachments.Add(selectedMail, Outlook.OlAttachmentType.olByValue, 1, selectedMail.Subject);
                    ((Outlook._MailItem)mail).Send();
                    MessageBox.Show("Mail sent.", "succes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    key.Close();
                }
            }
            catch
            {
                MessageBox.Show("Please select a mail", "Send to Cyber", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
