using Microsoft.Office.Interop.PowerPoint;
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
    }
}