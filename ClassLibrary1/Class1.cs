using System.Windows.Forms.Integration;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using ComBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;

namespace CustomReport2020
{
    [PluginAttribute("CustomReport.Report",                         //Plugin name
                    "com.arashnaderpour",                           //4 character Developer ID or GUID
                    ToolTip = "Export a customized report.",        //The tooltip for the item in the ribbon
                    DisplayName = "Report")]

    public class Report : AddInPlugin
    {
        //Document oDoc = Autodesk.Navisworks.Api.Application.ActiveDocument;

        // SubWindows: Generate Initial Data Window
        UserInputWindow userInputWindow;

        public override int Execute(params string[] parameters)
        {
            if(this.userInputWindow != null)
            {
                this.userInputWindow.Close();

                this.userInputWindow = new UserInputWindow();

                ElementHost.EnableModelessKeyboardInterop(this.userInputWindow);

                this.userInputWindow.Show();

                return 0;
            }
            else
            {
                this.userInputWindow = new UserInputWindow();

                ElementHost.EnableModelessKeyboardInterop(this.userInputWindow);

                this.userInputWindow.Show();

                return 0;
            }
        }
    }
}
