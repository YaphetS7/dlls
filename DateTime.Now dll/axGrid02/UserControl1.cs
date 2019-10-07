using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// Add references
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;
using Microsoft.Win32;


namespace axGrid02
{
    [ProgId("axGrid02.UserControl1")]
    [ClassInterface(ClassInterfaceType.AutoDual), ComSourceInterfaces(typeof(UserControlEvents))]
    //[ComClass(UserControl1.ClassId, UserControl1.InterfaceId, UserControl1.EventsId)]
    public partial class UserControl1: UserControl
    {
         //--< IDs GUID >--
    //public const String ClassId = "147DF0FB-83F3-4BCB-9D55-F36E29C221FE";
    //public const String InterfaceId = "E43550AD-D09C-48CD-9D6B-B07FDE61B45C";
    public const String EventsId = "EC2581AC-1994-4C16-8172-6C360D07C3C0";
    //--</ IDs GUID >--

        #region Control

        public UserControl1()
        {
            InitializeComponent();
            timer1.Start();
        }
        #endregion /Control

        #region Buttons
        //--------------------< region: Buttons >-----------------
       
        //--------------------< region: Buttons >-----------------
        #endregion  /Buttons

        #region Controls
        //----------< Controls >-------------
        private void ActiveXUserControl_Load(object sender, EventArgs e)
        {

        }
        //----------</ Controls >-------------
        #endregion /Controls


        #region Sys: Register As COM ActiveX
        //--------------------< region: Sys:Register COM ActiveX >-----------------
        // register COM ActiveX object
        [ComRegisterFunction()]
        public static void RegisterFunction(Type _type)
        {
            if (_type != null)
            {
                string sCLSID = "CLSID\\" + _type.GUID.ToString("B");
                try
                {
                    RegistryKey _key = Registry.ClassesRoot.OpenSubKey(sCLSID, true);
                    try
                    {
                        Guid _libID = Marshal.GetTypeLibGuidForAssembly(_type.Assembly);
                        int _major, _minor;
                        Marshal.GetTypeLibVersionForAssembly(_type.Assembly, out _major, out _minor);
                        using (RegistryKey _sub = _key.CreateSubKey("Control")) { }
                        using (RegistryKey _sub = _key.CreateSubKey("MiscStatus")) { _sub.SetValue("", "0", RegistryValueKind.String); }
                        using (RegistryKey _sub = _key.CreateSubKey("TypeLib")) { _sub.SetValue("", _libID.ToString("B"), RegistryValueKind.String); }
                        using (RegistryKey _sub = _key.CreateSubKey("Version")) { _sub.SetValue("", String.Format("{0}.{1}", _major, _minor), RegistryValueKind.String); }
                        using (RegistryKey _sub = _key.CreateSubKey("Control")) { }
                        using (RegistryKey _sub = _key.CreateSubKey("InprocServer32")) { _sub.SetValue("", Environment.SystemDirectory + "\\" + _sub.GetValue("", "mscoree.dll"), RegistryValueKind.String); }
                    }
                    finally
                    {
                    }
                }
                catch
                {
                }
            }
        }

        [ComUnregisterFunction()]
        public static void UnregisterClass(Type t) //(string key)
        {
            string keyName = @"CLSID\" + t.GUID.ToString("B");
            Registry.ClassesRoot.DeleteSubKeyTree(keyName);
            //StringBuilder skey = new StringBuilder(key);
            //skey.Replace(@"HKEY_CLASSES_ROOT\", "");
            //RegistryKey regKey = Registry.ClassesRoot.OpenSubKey(skey.ToString(), true);
            //regKey.DeleteSubKey("Control", false);
            //RegistryKey inprocServer32 = regKey.OpenSubKey("InprocServer32", true);
            //regKey.DeleteSubKey("CodeBase", false);
        }
        //--------------------< region: /Sys:Register COM ActiveX >-----------------
            #endregion /Sys:Register As COM ActiveX


            #region Interface Events
            //--------------------< region: Interface Events >-----------------
            //*Eventhandler interface    
        public delegate void ControlEventHandler(int NumVal);

        [Guid(EventsId)]
        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        public interface UserControlEvents
        {
            [DispId(0x60020001)]
            void OnButtonClick(int NumVal);
        }

        //----------< Events: Buttons >-------------
        public event ControlEventHandler OnButtonClick;
        public event ControlEventHandler OnBtnSetClick;
        public event ControlEventHandler btnGetClick;
        //----------</ Events: Buttons >-------------

        //--------------------</ region: Interface Events >-----------------
        #endregion /events

        #region Interface Properties
        //--------------------< region: Properties >-----------------
        //*Values are exchanged between Hosting-App and ActiveX-Control as GET-/ SET- Property
        public string Textbox1
        {
            get
            {
               // ptextVal = (int)(numericUpDown1.Value);
                return tbxText1.Text ;
            }
            set
            {
                tbxText1.Text = value;
            }
         
        }
        private void Tick()
        {
            Textbox1 = Convert.ToString(DateTime.Now);
        }
        //--------------------</ region: Properties >-----------------
        #endregion /Interface Properties

        #region Interface Methods, Functions
        //--------------------< region: Methods, Functions >-----------------
        public interface ICOMCallable
        {
            string axSet_Textbox1_Text();
        }

        public string axSet_Textbox1_Text()
        {
            //int i = (int)(numericUpDown1.Value);
            //MessageBox.Show("ActiveX method: GetTextBoxValue " + i.ToString());
            //return (i);
            tbxText1.Text = "set by extenal function(..)";
            return "set Textbox1";
        }
        //--------------------</ region: Methods, Functions >-----------------
        #endregion /Interface Methods, Functions

        private void timer1_Tick(object sender, EventArgs e)
        {
            Tick();
        }
    }

}



//notes:
//Referenz: http://www.codeguru.com/csharp/.net/net_general/comcom/article.php/c16257/Create-an-ActiveX-using-a-Csharp-Usercontrol.htm