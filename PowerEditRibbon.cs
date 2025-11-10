//using Microsoft.Office.Tools.Ribbon;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace PowerEditAddIn
//{
//    public partial class PowerEditRibbon
//    {
//        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
//        {

//        }

//        private void btnPowerEdit_Click(object sender, RibbonControlEventArgs e)
//        {
//            System.Windows.Forms.MessageBox.Show("Button Clicked!");
//            Globals.ThisAddIn.TogglePane();
//        }

//    }
//}



using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerEditAddIn
{
    public partial class PowerEditRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnPowerEdit_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // ✅ Optionally show old debug message (you can remove this line if not needed)
                // System.Windows.Forms.MessageBox.Show("Button Clicked!");

                // ✅ Open both panes together (left + right)
                Globals.ThisAddIn.OpenBothPanes();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
