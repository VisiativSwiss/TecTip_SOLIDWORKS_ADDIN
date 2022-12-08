using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

namespace SolidWorks_Addin
{
    [Guid("7477997F-8ABD-44AF-B59D-A40F6694F7D9"),ComVisible(true)]
    public class MainClass : SwAddin
    {
        SldWorks swApp;
        ICommandManager iCmdMgr;
        int _addinId;
        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = (SldWorks)ThisSW;
            _addinId = Cookie;
            swApp.SetAddinCallbackInfo2(0, this, _addinId);
            iCmdMgr = swApp.GetCommandManager(Cookie);
            int errors = 0;
            CommandGroup commandGroup = swApp.GetCommandManager(_addinId).CreateCommandGroup2(0, "MyCoolTitle", "MyCoolToolTip", "MyCoolHint", 6, true, ref errors);
            if (commandGroup != null)
            {
                int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
                commandGroup.IconList = null;
                commandGroup.ShowInDocumentType = (int)swDocTemplateTypes_e.swDocTemplateTypePART;
                commandGroup.AddCommandItem2("SampleAddin", -1, "MyCoolHint", "MyCoolToolTip", 0, "Main", "EnableButton", 0, menuToolbarOption);
                commandGroup.HasMenu = true;
                commandGroup.HasToolbar = true;
                commandGroup.Activate();
            }

            return true;
        }
        public int EnableButton()
        {
            return 1;
        }
        public bool DisconnectFromSW()
        {
            try
            {
                iCmdMgr.RemoveCommandGroup(1);
                Marshal.ReleaseComObject(swApp);
                swApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception)
            {

            }

            return true;
        }

        public void Main()
        {
            MessageBox.Show("Hallo Welt");
        }
    }
}
