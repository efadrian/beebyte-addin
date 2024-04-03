using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    public interface ISlideService
    {
        void AddSlide(Application pptApp);

        void RemoveSlide(Application pptApp);

        //
        void FontSizePlus(Application pptApp);

        void FontSizeMinus(Application pptApp);

        void CopyText(Application pptApp);

        void PasteText(Application pptApp);
    }
}
