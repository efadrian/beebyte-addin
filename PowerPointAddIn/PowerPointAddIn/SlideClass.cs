using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    public class MyClass
    {
        private readonly ISlideService _SlideService;

        public MyClass(ISlideService SlideService)
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
    }
}
