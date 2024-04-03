using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    public class SlideClass
    {
        private readonly ISlideService _SlideService;

        public SlideClass(ISlideService SlideService)
        {
            _SlideService = SlideService;
        }

        public void AddSlide(Application pptApp)
        {
            _SlideService.AddSlide(pptApp);
        }

        public void RemoveSlide(Application pptApp)
        {
            _SlideService.RemoveSlide(pptApp);
        }

        //
        public void FontSizePlus(Application pptApp)
        {
            _SlideService.FontSizePlus(pptApp);
        }
        public void FontSizeMinus(Application pptApp)
        {
            _SlideService.FontSizeMinus(pptApp);
        }
        public void CopyText(Application pptApp)
        {
            _SlideService.CopyText(pptApp);
        }
        public void PasteText(Application pptApp)
        {
            _SlideService.PasteText(pptApp);
        }
    }
}
