using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    public interface ISlideService
    {
        void AddSlide(Application pptApp);

        void RemoveSlide(Application pptApp);

    }
}
