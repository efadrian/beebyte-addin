using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using PowerPointAddIn.Service;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
namespace PowerPointAddIn
{
    public class SlideService : ISlideService
    {
        private float? shapePositionX;
        private float? shapePositionY;
        private Shape _selectedShape;

        #region slide
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

        #endregion

        #region font size
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

        #endregion

        #region text 

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

        private TextRange GetSelectedText(Application pptApp)
        {
            TextRange textRange = null;

            if (pptApp.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionText)
            {
                textRange = pptApp.ActiveWindow.Selection.TextRange;
            }

            return textRange;
        }
        #endregion

        #region copy position

        public void CopyPosition(Application pptApp)
        {
            _selectedShape = GetSelectedShape(pptApp);

            if (_selectedShape != null)
            {
                shapePositionX = _selectedShape.Left;
                shapePositionY = _selectedShape.Top;
            }
        }

        public void PastePosition(Application pptApp)
        {
            if (_selectedShape != null && shapePositionX.HasValue && shapePositionY.HasValue)
            {
                _selectedShape.Left = shapePositionX.Value;
                _selectedShape.Top = shapePositionY.Value;
            }
        }
        private Shape GetSelectedShape(Application pptApp)
        {
            Selection allShapes = pptApp.ActiveWindow.Selection;
            Shape selectedshape = allShapes.ShapeRange[1];
            return selectedshape;
        }

        #endregion

        #region align

        public void AlignLeft(Application pptApp)
        {
            AlignShapes(pptApp, Alignment.Left);
        }

        public void AlignRight(Application pptApp)
        {
            AlignShapes(pptApp, Alignment.Right);
        }

        public void AlignTop(Application pptApp)
        {
            AlignShapes(pptApp, Alignment.Top);
        }

        public void AlignBottom(Application pptApp)
        {
            AlignShapes(pptApp, Alignment.Bottom);
        }

        private void AlignShapes(Application pptApp, Alignment align)
        {
            Selection allShapes = pptApp.ActiveWindow.Selection;
            float sWidth = pptApp.ActivePresentation.PageSetup.SlideWidth;
            float sHeight = pptApp.ActivePresentation.PageSetup.SlideHeight;

            foreach (Shape shape in allShapes.ShapeRange)
            {
                switch (align)
                {
                    case Alignment.Left:
                        shape.Left = 0;
                        break;
                    case Alignment.Right:
                        shape.Left = sWidth - shape.Width;
                        break;
                    case Alignment.Top:
                        shape.Top = 0;
                        break;
                    case Alignment.Bottom:
                        shape.Top = sHeight - shape.Height;
                        break;
                    default:
                        break;
                }
            }
        }

        public void checkShapesStatus(Application pptApp, Selection Sel)
        {
            var ribbon = Globals.Ribbons.GetRibbon<Ribbon>();
            // Text
            RibbonGroup groupText = ribbon.groupText;

            if (Sel.Type == PpSelectionType.ppSelectionText)
            {
                ToigleGroup(groupText, true);
            }
            else
            {
                ToigleGroup(groupText, false);
            }
            // shape
            RibbonGroup groupShape = ribbon.groupShape;

            if (Sel.Type == PpSelectionType.ppSelectionShapes)
            {
                ToigleGroup(groupShape, true);
            }
            else
            {
                ToigleGroup(groupShape, false);
            }
        }

        public void ToigleGroup(RibbonGroup group, bool flag)
        {
            if (group != null)
            {
                foreach (var control in group.Items)
                {
                    control.Enabled = flag;
                }
            }
        }
    }


    #endregion
}