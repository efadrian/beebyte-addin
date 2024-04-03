using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    enum HorizontalAlignment
    {
        Left,
        Right
    }

    enum VerticalAlignment
    {
        Top,
        Bottom
    }

    public class SlideClass
    {
        private readonly ISlideService _service;

        public SlideClass(ISlideService SlideService)
        {
            _service = SlideService;
        }

        public void AddSlide(Application pptApp)
        {
            _service.AddSlide(pptApp);
        }

        public void RemoveSlide(Application pptApp)
        {
            _service.RemoveSlide(pptApp);
        }

        //
        public void FontSizePlus(Application pptApp)
        {
            _service.FontSizePlus(pptApp);
        }
        public void FontSizeMinus(Application pptApp)
        {
            _service.FontSizeMinus(pptApp);
        }
        public void CopyText(Application pptApp)
        {
            _service.CopyText(pptApp);
        }
        public void PasteText(Application pptApp)
        {
            _service.PasteText(pptApp);
        }

        public void CopyPosition(Application pptApp)
        {
            _service.CopyPosition(pptApp);
        }
        public void PastePosition(Application pptApp)
        {
            _service.PastePosition(pptApp);
        }
        public void AlignLeft(Application pptApp)
        {
            _service.AlignLeft(pptApp);
        }
        public void AlignRight(Application pptApp)
        {
            _service.AlignRight(pptApp);
        }
    }
}
