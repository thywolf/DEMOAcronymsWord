using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DEMOAcronymsWordAddIn
{
    public partial class AcronymsTaskPane : UserControl
    {
        public AcronymsTaskPane()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        public void setLabels(string _selectedWord)
        {
            labelProposalHeader.Text = "Explanations for " + _selectedWord;
            labelProposal1.Text = "Text 1 for " + _selectedWord;
            labelProposal2.Text = "Text 2 for " + _selectedWord;
            labelProposal3.Text = "Text 3 for " + _selectedWord;
            labelProposal4.Text = "Text 4 for " + _selectedWord;
            labelProposal5.Text = "Text 5 for " + _selectedWord;
        }
    }
}
