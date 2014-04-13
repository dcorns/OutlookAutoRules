using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;


namespace OutlookAutoRules
{
    public partial class ThisAddIn
    {
        //private Office.CommandBar menuBar;
        //private Office.CommandBarPopup newMenuBar;
        //private Office.CommandBarButton buttonOne;
        //private Office.CommandBarButton buttonTwo;
        
        
        private Outlook.MailItem item;
        private Office.CommandBarButton btnCheckForRule;
        private Outlook.Selection selection;
        private List<Outlook.Rule> RulesList = new List<Outlook.Rule>();
        private Outlook.Stores AllStores;
        private List<Office.CommandBarButton> btnRule = new List<Office.CommandBarButton>();
        
        
        
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //AddMenuBar();
            
            Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);
            Application.ContextMenuClose += new Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler(Application_ContextMenuClose);
            
            AllStores = Application.Session.Stores;
            foreach (Outlook.Store OS in AllStores)
            {
                foreach (Outlook.Rule OR in OS.GetRules())
                {
                    RulesList.Add(OR);
                }
            }
            
        }
        //This
        void Application_ContextMenuClose(Outlook.OlContextMenu ContextMenu)
        {
            selection = null;
            item = null;
            if (btnCheckForRule != null)
            {
                btnCheckForRule.Click -= new Office
                    ._CommandBarButtonEvents_ClickEventHandler(
                    btnRules_Click);
            }
            btnCheckForRule = null;

        }
        //
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //this
            Application.ItemContextMenuDisplay -= new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay);
            Application.ContextMenuClose -= new Outlook.ApplicationEvents_11_ContextMenuCloseEventHandler(Application_ContextMenuClose);
            //
        }
        
        //This
        void Application_ItemContextMenuDisplay(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {
            selection = Selection;
            if (GetMessageClass(selection[1])=="IPM.Note" &&selection.Count==1)
            {
                item = (Outlook.MailItem)selection[1];
                btnCheckForRule = (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing);
                btnCheckForRule.Caption = "CheckRules";
                btnCheckForRule.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btnCheckForRule_Click);
               
                foreach (Outlook.Rule R in Application.Session.DefaultStore.GetRules())
                {
               Office.CommandBarButton test=     (Office.CommandBarButton)CommandBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing);
               test.Caption = @"ADD2RULE: "+R.Name;
               test.Click += new Office._CommandBarButtonEvents_ClickEventHandler(btnRules_Click);
                }
             
                  
            }
        }

        
        
        void btnRules_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            
            int rulecount = 1;
            int ruleindex = 1;
            string RuleName = Ctrl.Caption.Substring(10);//Load selected Rule Name from button caption
            bool AlreadyThere = false;
            Outlook.MailItem citem = Application.ActiveExplorer().Selection[1];//Load selected mail item
          
            Outlook.Rules MyRules = Application.Session.DefaultStore.GetRules();//Retrieve Rules
           
            Outlook.Folder Folder = Application.ActiveExplorer().CurrentFolder//Retrieve Folder
        as Outlook.Folder;
            string Email = citem.SenderEmailAddress;//Extract Selected Mail Item Email Address
            
            foreach (Outlook.Rule RL in MyRules)
            {
                
                if (RuleName == RL.Name)
                {
                    ruleindex = rulecount;//Assign indext to selected Rule
                }
                rulecount++;
            }


            foreach (Outlook.RuleCondition RC in MyRules[ruleindex].Conditions)
            {

                if (RC.Enabled) //Add selected item parts to respective conditions if condition enabled
                {
                    switch (RC.ConditionType)//When I put this in as switch condition vs automatically added all the case statements below!
                    {
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionAccount:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionAnyCategory:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionBody:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionBodyOrSubject:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionCategory:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionCc:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionDateRange:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFlaggedForAction:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFormName:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFrom:

                            foreach (Outlook.Recipient ite in MyRules[ruleindex].Conditions.From.Recipients)
                            {
                                string rec = ite.Address;
                                string name = ite.Name;
                                if (rec == item.SenderEmailAddress || name==item.SenderName) AlreadyThere = true;

                            }


                            if (!AlreadyThere)
                            {
                                MyRules[ruleindex].Conditions.From.Recipients.Add(item.SenderEmailAddress);
                                MyRules[ruleindex].Conditions.From.Recipients.Add(item.SenderName);
                                MyRules[ruleindex].Conditions.From.Recipients.GetEnumerator();//This here thingy keeps the new address and rule condition from returning void and hince being added multiple times and not executing when rule is run. Otherwise even though the new condition recipient shows up in the wizard, but has no effect when wizard is run.
                                
                            }
                            else System.Windows.Forms.MessageBox.Show(item.SenderEmailAddress + @" already in sender list!",
                                   "Error Adding To Rule", System.Windows.Forms.MessageBoxButtons.OK);

                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFromAnyRssFeed:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionFromRssFeed:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionHasAttachment:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionImportance:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionLocalMachineOnly:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionMeetingInviteOrUpdate:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionMessageHeader:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionNotTo:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionOOF:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionOnlyToMe:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionOtherMachine:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionProperty:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionRecipientAddress:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSenderAddress:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSenderInAddressBook:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSensitivity:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSentTo:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSizeRange:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionSubject:
                            //Next to do#########################################
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionTo:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionToOrCc:
                            break;
                        case Microsoft.Office.Interop.Outlook.OlRuleConditionType.olConditionUnknown:
                            break;
                        default:
                            break;
                    }
                    
                    
                }
            }
            
            
            MyRules.Save(false);
           
            
            MyRules[ruleindex].Execute(true, Folder, Type.Missing, Type.Missing);
            
        }

        void btnCheckForRule_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Outlook.MailItem sitem = Application.ActiveExplorer().Selection[1];//Load selected mail item
            Outlook.Rules TheRules = Application.Session.DefaultStore.GetRules();//Retrieve Rules
            string TheEmail = sitem.SenderEmailAddress;
            string TheName = sitem.SenderName;
            string TheDisplayName = sitem.SenderName;
            int recipientindex=1;
            int deleteADRindex = 0;
            bool deleteADRrecipient = false;
            int deleteNAMEindex = 0;
            bool deleteNAMErecipient = false;
            bool NoMatchFound = true;
            int RecipientCount = 0;
            foreach (Outlook.Rule RL in TheRules)
            {
                recipientindex = 1;
                RecipientCount = RL.Conditions.From.Recipients.Count;
                
                foreach (Outlook.Recipient s in RL.Conditions.From.Recipients)//Check by sender address or name
                {
                    if (s.Name == TheName)
                    {
                        DialogResult Namefound;
                        Namefound = System.Windows.Forms.MessageBox.Show(TheName + @" exist in Rule <" + RL.Name + @">" + Environment.NewLine + "Remove from Rule?",
                                    "Rule Found", System.Windows.Forms.MessageBoxButtons.YesNo);
                        NoMatchFound = false;
                        if (Namefound == DialogResult.Yes)
                        {
                            deleteNAMErecipient = true;
                            deleteNAMEindex = recipientindex;
                        }
                    }

                    if (s.Address == TheEmail)//skip if already being removed based on name
                    {
                        DialogResult ADRfound;
                       ADRfound=  System.Windows.Forms.MessageBox.Show(TheEmail + @" exist in Rule <"+RL.Name+@">"+Environment.NewLine+"Remove from Rule?",
                                   "Rule Found", System.Windows.Forms.MessageBoxButtons.YesNo);
                       NoMatchFound = false;
                       if (ADRfound == DialogResult.Yes)
                       {
                           deleteADRrecipient = true;
                           deleteADRindex = recipientindex;
                       }


                    }
                    
                    recipientindex++;
                }
                if (deleteADRrecipient)
                {
                    deleteADRindex = deleteADRindex - (RecipientCount - RL.Conditions.From.Recipients.Count);//Adjust index to account for any deletions made to Recipient list after index was set
                    RL.Conditions.From.Recipients.Remove(deleteADRindex);
                    RL.Conditions.From.Recipients.GetEnumerator();
                    deleteADRrecipient = false;
                    
                }
                if(deleteNAMErecipient)
                {
                    deleteNAMEindex = deleteNAMEindex - (RecipientCount - RL.Conditions.From.Recipients.Count);//Adjust index to account for any deletions made to Recipient list after index was set
                    RL.Conditions.From.Recipients.Remove(deleteNAMEindex);
                    RL.Conditions.From.Recipients.GetEnumerator();
                    deleteNAMErecipient = false;

                }
            }
            if (NoMatchFound) System.Windows.Forms.MessageBox.Show("Not in any Rules",
                                    "Rule Check", System.Windows.Forms.MessageBoxButtons.OK);
            else TheRules.Save(false);
        }
        private string GetMessageClass(object item)
        {
            object[] args = new Object[] { };
            Type t = item.GetType();
            return t.InvokeMember("messageClass", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.GetField | System.Reflection.BindingFlags.GetProperty, null, item, args).ToString();
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
