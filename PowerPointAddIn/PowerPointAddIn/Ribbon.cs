using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace PowerPointAddIn
{
    public partial class Ribbon
    {
      
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            _pptApp = Globals.ThisAddIn.Application;
            // event selection
            _pptApp.WindowSelectionChange += SelectionChangeEvent;
        }

        private void SelectionChangeEvent(Selection Sel)
        {
            _slideClass.checkShapesStatus(_pptApp, Sel);
        }


        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.AddSlide(_pptApp);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.RemoveSlide(_pptApp);
        }

        private void fSizePlus_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.FontSizePlus(_pptApp);
        }

        private void fSizeMinus_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.FontSizeMinus(_pptApp);
        }

        private void copyTxt_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.CopyText(_pptApp);
        }

        private void pasteTxt_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.PasteText(_pptApp);
        }

        //
        private void copyPosition_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.CopyPosition(_pptApp);
        }

        private void pastePosition_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.PastePosition(_pptApp);
        }

        private void alignLeft_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.AlignLeft(_pptApp);
        }

        private void alignTop_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.AlignTop(_pptApp);
        }

        private void alignRight_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.AlignRight(_pptApp);
        }

        private void alignBottom_Click(object sender, RibbonControlEventArgs e)
        {
            _slideClass.AlignBottom(_pptApp);
        }

        //

    }
}
