using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
namespace PowerPointAddIn
{
    public class SlideService : ISlideService
    {
        private float? storedPositionX;
        private float? storedPositionY;
        private Shape  _selectedShape;
        //

        public void AddSlide(Application pptApp)
        {
            Slide currentSlide = pptApp.ActiveWindow.View.Slide;
            if (currentSlide != null)
            {
                pptApp.ActivePresentation.Slides.Add(currentSlide.SlideIndex + 1, PpSlideLayout.ppLayoutText);
            }
            else
            {
                pptApp.ActivePresentation.Slides.Add(pptApp.ActivePresentation.Slides.Count + 1, PpSlideLayout.ppLayoutText);
            }
        }

        public void RemoveSlide(Application pptApp)
        {
            Slide currentSlide = pptApp.ActiveWindow.View.Slide;

            if (currentSlide != null)
            {
                currentSlide.Delete();
            }
        }

        //

        public void FontSizePlus(Application pptApp)
        {
            TextRange selectedTextRange = GetSelectedText(pptApp);

            if (selectedTextRange != null)
            {
                selectedTextRange.Font.Size += 1;
            }
        }

        public void FontSizeMinus(Application pptApp)
        {
            TextRange selectedTextRange = GetSelectedText(pptApp);

            if (selectedTextRange != null)
            {
                selectedTextRange.Font.Size -= 1;
            }
        }

        private TextRange GetSelectedText(Application pptApp)
        {
            TextRange textRange = null;

            if (pptApp.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
            {
                textRange = pptApp.ActiveWindow.Selection.TextRange;
            }

            return textRange;
        }


        public void CopyText(Application pptApp)
        {
            TextRange selectedTextRange = GetSelectedText(pptApp);

            if (selectedTextRange != null)
            {
               Clipboard.SetText(selectedTextRange.Text);
            }
        }

        public void PasteText(Application pptApp)
        {
            TextRange selectedTextRange = GetSelectedText(pptApp);

            if (selectedTextRange != null)
            {
                selectedTextRange.Text = Clipboard.GetText();
            }
        }

        //

        public void CopyPosition(Application pptApp)
        {
            _selectedShape = GetSelectedShape();
            if (_selectedShape != null)
            {
                storedPositionX = _selectedShape.Left;
                storedPositionY = _selectedShape.Top;
            }
        }

        public void PastePosition(Application pptApp)
        {
            if (_selectedShape != null && storedPositionX.HasValue && storedPositionY.HasValue)
            {
                _selectedShape.Left = storedPositionX.Value;
                _selectedShape.Top = storedPositionY.Value;
            }
        }

        //

        public void AlignLeft(Application pptApp)
        {
            //AlignShapes(HorizontalAlignment.Left,pptApp);
        }

        public void AlignRight(Application pptApp)
        {
            throw new NotImplementedException();
        }

        public void AlignTop(Application pptApp)
        {
            throw new NotImplementedException();
        }

        public void AlignBottom(Application pptApp)
        {
            throw new NotImplementedException();
        }

        private void AlignShapes(MsoAlignCmd alignCmd, Application pptApp)
        {
            var selection = pptApp.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count >= 2)
            {
                selection.ShapeRange.Align(alignCmd, MsoTriState.msoFalse);
            }
        }

        private Shape GetSelectedShape()
        {
            Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1)
            {
                return selection.ShapeRange[1];
            }
            return null;
        }
    }
}