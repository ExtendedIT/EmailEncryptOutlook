using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace EmailEncrypt
{
    public partial class Ribbon1
    {


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }



        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var m = e.Control.Context as Inspector;
            var mailItem = m.CurrentItem as MailItem;

            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "[Encrypted] " + mailItem.Subject;
                }
            
                      }

        }
    }
}
