using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace DEMOAcronymsWordAddIn
{
    [ComVisible(true)]
    public class AcronymsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private bool _buttonClicked;

        public AcronymsRibbon()
        {
            _buttonClicked = false;
        }

        public bool EnableControl(IRibbonControl control)
        {
            if (control.Id == "xxxProposal5")
            {
                return false;
            }
            else
            {
                return true; // visible ... false = invisible
            }

        }

        public void YourUniqueId_Click(Office.IRibbonControl Control)
        {
            //Since the initial value is false and presumably the user just clicked for 
            //the first (or N-th) time you'll want to set the value to true
            if (!_buttonClicked)
            {
                _buttonClicked = true;
            }
            //Or if clicking for a second (or N-th + 1) time, set the value to false
            else
            {
                _buttonClicked = false;
            }

            //Now use the invalidate method from the ribbon variable (from the load method) 
            //to reset the specific control id (in this case "YourUniqueId") from the xml. 
            //Invalidating the control will call the method "GetYourLabelText"
            ribbon.InvalidateControl(Control.Id);
        }

        public void OpenTaskPane_Click(Office.IRibbonControl Control)
        {
            Globals.ThisAddIn.OpenTaskPane();
        }
        public string GetYourLabelText(Office.IRibbonControl Control)
        {
            return "Text " + Control.Id.ToString().Last() + " for " + Globals.ThisAddIn.selectedWord;
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordAddIn2.AcronymsRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
