/*
    OpenERP, Open Source Business Applications
    Copyright (c) 2011 OpenERP S.A. <http://openerp.com>

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU Affero General Public License as
    published by the Free Software Foundation, either version 3 of the
    License, or (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU Affero General Public License for more details.

    You should have received a copy of the GNU Affero General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/


using System.Net;
using System.Net.Sockets;
using Microsoft.Win32;
using NetOffice;
using NetOffice.OfficeApi.Enums;

namespace OpenERPOutlookPlugin
{
    using System;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using OpenERPClient;
    using Outlook = NetOffice.OutlookApi;
    using Office = NetOffice.OfficeApi;


    

    #region Read me for Add-in installation and setup information.
    // When run, the Add-in wizard prepared the registry for the Add-in.
    // At a later time, if the Add-in becomes unavailable for reasons such as:
    //   1) You moved this project to a computer other than which is was originally created on.
    //   2) You chose 'Yes' when presented with a message asking if you wish to remove the Add-in.
    //   3) Registry corruption.
    // you will need to re-register the Add-in by building the OpenERPOutlookPluginSetup project, 
    // right click the project in the Solution Explorer, then choose install.
    #endregion

    /// <summary>
    ///   The object for implementing an Add-in.
    /// </summary>
    /// <seealso class='IDTExtensibility2' />

    [GuidAttribute("C86B5760-1254-4F40-BD25-2094A2A678C4"), ProgId("OpenERPOutlookPlugin.Connect"), ComVisible(true)]
    public class Connect : Object, Extensibility.IDTExtensibility2
    {
        private static readonly string _addinOfficeRegistryKey = "Software\\Microsoft\\Office\\Outlook\\AddIns\\";
        private static readonly string _prodId = "OpenERPOutlookPlugin.Connect";
        private static readonly string _addinFriendlyName = "OpenERPOplugin";
        private static readonly string _addinDescription = "OpenERPOplugin used NetOffice";
        Outlook.Application _outlookApplication ;


        #region IDTExtensibility2
        /// <summary>
        ///		Implements the constructor for the Add-in object.
        ///		Place your initialization code within this method.
        /// </summary>

        public int cnt_mail = 0;

        /// <summary>
        ///      Implements the OnConnection method of the IDTExtensibility2 interface.
        ///      Receives notification that the Add-in is being loaded.
        /// </summary>
        /// <param term='application'>
        ///      Root object of the host application.
        /// </param>
        /// <param term='connectMode'>
        ///      Describes how the Add-in is being loaded.
        /// </param>
        /// <param term='addInInst'>
        ///      Object representing this Add-in.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        [STAThread]
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            try
            {
                _outlookApplication = new Outlook.Application(null, application);
                NetOffice.OutlookSecurity.Suppress.Enabled = true;
            }
            catch (Exception ex)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message);
                MessageBox.Show(message, "OPENERP-OnConnection", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        ///     Implements the OnDisconnection method of the IDTExtensibility2 interface.
        ///     Receives notification that the Add-in is being unloaded.
        /// </summary>
        /// <param term='disconnectMode'>
        ///      Describes how the Add-in is being unloaded.
        /// </param>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            try
            {
                if (null != _outlookApplication)
                    _outlookApplication.Dispose();
            }
            catch (Exception ex)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message);
                MessageBox.Show(message, "OPENERP-OnDisconetion", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        ///      Implements the OnAddInsUpdate method of the IDTExtensibility2 interface.
        ///      Receives notification that the collection of Add-ins has changed.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref System.Array custom)
        {
        }

        /// <summary>
        ///      Implements the OnStartupComplete method of the IDTExtensibility2 interface.
        ///      Receives notification that the host application has completed loading.
        /// </summary>
        /// <param term='custom'>
        ///      Array of parameters that are host application specific.
        /// </param>
        /// <seealso class='IDTExtensibility2' />
        /// 
        private Office.CommandBarButton btn_open_partner;
        private Office.CommandBarButton btn_open_document;
        private Office.CommandBarButton btn_open_configuration_form;
        //private Office.CommandBars oCommandBars;
        private Office.CommandBar menuBar;
        private Office.CommandBarPopup newMenuBar;

        public int countMail()
        {
            /*
             
             * Gives the number of selected mail.
             * returns: Number of selected mail.
             
             */
            cnt_mail = 0;
            Outlook.Application app = _outlookApplication;

            //app = new Microsoft.Office.Interop.Outlook.Application();

            try
            {
                cnt_mail = app.ActiveExplorer().Selection.Count;
            }
            catch (Exception)
            {
                return 0;
            }
            return cnt_mail;
        }
      
        public void OnStartupComplete(ref System.Array custom)
        {
           
            //-----------------------------------------------------------------------------------------------------------------------------------------------------
           /* /*
             
             * When outlook is opened it loads a Menu if Outlook plugin is installed.
             * OpenERP - > Push, Partner ,Documents, Configuration
             
             #1#
         */
            Outlook.Application app = _outlookApplication;
            try
            {
                object omissing = System.Reflection.Missing.Value;
                menuBar = app.ActiveExplorer().CommandBars.ActiveMenuBar;
                ConfigManager config = new ConfigManager();
                config.LoadConfigurationSetting();
                OpenERPOutlookPlugin openerp_outlook = Cache.OpenERPOutlookPlugin;
                OpenERPConnect openerp_connect = openerp_outlook.Connection;
                try
                {
                    if (openerp_connect.URL != null && openerp_connect.DBName != null && openerp_connect.UserId != null && openerp_connect.pswrd != "")
                    {                        
                        string decodpwd = Tools.DecryptB64Pwd(openerp_connect.pswrd);
                        openerp_connect.Login(openerp_connect.DBName, openerp_connect.UserId, decodpwd);                            
                    }
                }
                catch(Exception )
                {
                    //just shallow exception
                }
                newMenuBar = (Office.CommandBarPopup)menuBar.Controls.Add(MsoControlType.msoControlPopup, omissing, omissing, omissing, true);
                if (newMenuBar != null)
                {
                    newMenuBar.Caption = "OpenERP";
                    newMenuBar.Tag = "My";

                    btn_open_partner = (Office.CommandBarButton)newMenuBar.Controls.Add(MsoControlType.msoControlButton, omissing, omissing, 1, true);
                    btn_open_partner.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btn_open_partner.Caption = "Contact";
                    //Face ID will use to show the ICON in the left side of the menu.
                    btn_open_partner.FaceId = 3710;
                    newMenuBar.Visible = true;
                    btn_open_partner.ClickEvent += new Office.CommandBarButton_ClickEventHandler(this.btn_open_partner_Click); //Core._CommandBarButtonEvents_ClickEventHandler(this.btn_open_partner_Click);

                    btn_open_document = (Office.CommandBarButton)newMenuBar.Controls.Add(MsoControlType.msoControlButton, omissing, omissing, 2, true);
                    btn_open_document.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btn_open_document.Caption = "Documents";
                    //Face ID will use to show the ICON in the left side of the menu.
                    btn_open_document.FaceId = 258;
                    newMenuBar.Visible = true;
                    btn_open_document.ClickEvent += new Office.CommandBarButton_ClickEventHandler(this.btn_open_document_Click);

                    btn_open_configuration_form = (Office.CommandBarButton)newMenuBar.Controls.Add(MsoControlType.msoControlButton, omissing, omissing, 3, true);
                    btn_open_configuration_form.Style = MsoButtonStyle.msoButtonIconAndCaption;
                    btn_open_configuration_form.Caption = "Configuration";
                    //Face ID will use to show the ICON in the left side of the menu.
                    btn_open_configuration_form.FaceId = 5644;
                    newMenuBar.Visible = true;
                    btn_open_configuration_form.ClickEvent += new Office.CommandBarButton_ClickEventHandler(this.btn_open_configuration_form_Click);

                }

            }
            catch (Exception ex)
            {
                string message = string.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message);
                MessageBox.Show(message, "OPENERP-Initialize menu", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
                 

        }
        #endregion

        void btn_open_configuration_form_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            frm_openerp_configuration frm_config = new frm_openerp_configuration();
            frm_config.Show();

        }


        public static bool isLoggedIn()
        {
            /*
             
             * This will check that it is connecting with server or not.
             * If wrong server name or port is given then it will throw the message.
             * returns true If conneted with server, otherwise False.
             
             */
            if (Cache.OpenERPOutlookPlugin == null || Cache.OpenERPOutlookPlugin.isLoggedIn == false)
            {
                throw new Exception("OpenERP Server is not connected!\nPlease connect OpenERP Server from Configuration Menu.");
            }
            return true;
        }

        public static void handleException(Exception e)
        {
            string Title;
            if (Form.ActiveForm != null)
            {
                Title = Form.ActiveForm.Text;
            }
            else
            {
                Title = "OpenERP Addin";
            }
            MessageBox.Show(e.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        public static void displayMessage(string message)
        {
            string Title;
            if (Form.ActiveForm != null)
            {
                Title = Form.ActiveForm.Text;
            }
            else
            {
                Title = "OpenERP Addin";
            }
            MessageBox.Show(message, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public void CheckMailCount()
        {
            if (countMail() == 0)
            {
                throw new Exception("No email selected.\nPlease select one email.");
            }
            if (countMail() > 1)
            {
                throw new Exception("Multiple selction is not allowed.\nPlease select only one email.");
            }

        }
       
        void btn_open_partner_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Connect.isLoggedIn();
                this.CheckMailCount();
                if (countMail() == 1)
                {
                    foreach (Outlook.MailItem mailitem in Tools.MailItems())
                    {
                        

                        Object[] contact = Cache.OpenERPOutlookPlugin.RedirectPartnerPage(mailitem);
                        if ((int)contact[1] > 0)
                        {
                            Cache.OpenERPOutlookPlugin.RedirectWeb(contact[2]);
                        }
                        else
                        {
                            frm_contact contact_form = new frm_contact(mailitem.SenderName, mailitem.SenderEmailAddress);
                            contact_form.Show();
                        }

                    }
                }

            }
            catch (Exception e)
            {
                Connect.handleException(e);
            }

        }

        void btn_open_document_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            try
            {
                Connect.isLoggedIn();

                this.CheckMailCount();
                
                if (countMail() == 1)
                {
                    frm_choose_document_opt frm_doc = new frm_choose_document_opt();    
                }
            }
            catch (Exception e)
            {
                Connect.handleException(e);
            }

        }

        
        public void OnBeginShutdown(ref System.Array custom)
        {
        }

        //this available for debug purposes
        #region COM Register Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // add codebase value
                Assembly thisAssembly = Assembly.GetAssembly(typeof(Connect));
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable");

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();

                // add outlook addin key
                Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _prodId);
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _prodId, true);
                rk.SetValue("LoadBehavior", Convert.ToInt32(3));
                rk.SetValue("FriendlyName", _addinFriendlyName);
                rk.SetValue("Description", _addinDescription);
                rk.Close();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Register " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
                Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _prodId, false);
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Unregister " + _prodId, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}
