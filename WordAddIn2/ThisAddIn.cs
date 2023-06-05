using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace DEMOAcronymsWordAddIn
{
    public partial class ThisAddIn
    {
        public string selectedWord;
        private AcronymsTaskPane myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            selectedWord = "EXMP";
            myUserControl1 = new AcronymsTaskPane();
            myCustomTaskPane = CustomTaskPanes.Add(myUserControl1, "DEMO Acronyms");
            myCustomTaskPane.Width = 400;
            this.Application.WindowSelectionChange += new Microsoft.Office.Interop.Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
        }

        public void OpenTaskPane()
        {
            myCustomTaskPane.Visible = true;
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        void Application_WindowSelectionChange(Selection sel)
        {
            //if (sel.Range.Words.Count == 1 && sel.Information[WdInformation.wdWithInTable] == false)
            //{
                // If only one word is selected and the selection is not within a table, execute your method here
                MyMethod(sel.Range.Words[1].Text.Trim());
            //}
        }

        void MyMethod(string _selectedWord)
        {
            try
            {
                selectedWord = _selectedWord;
                myUserControl1.setLabels(_selectedWord);
            }
            catch { }
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new AcronymsRibbon();
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
