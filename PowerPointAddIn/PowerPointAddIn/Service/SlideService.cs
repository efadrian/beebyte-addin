using Microsoft.Office.Interop.PowerPoint;
using System;
namespace PowerPointAddIn
{
    public class SlideService : ISlideService
    {
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
                System.Windows.Forms.Clipboard.SetText(selectedTextRange.Text);
            }
        }

        public void PasteText(Application pptApp)
        {
            TextRange selectedTextRange = GetSelectedText(pptApp);

            if (selectedTextRange != null)
            {
                selectedTextRange.Text = System.Windows.Forms.Clipboard.GetText();
            }
        }
    }
}