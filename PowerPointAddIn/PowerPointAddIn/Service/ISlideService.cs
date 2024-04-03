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

        //
        void CopyPosition(Application pptApp);
        void PastePosition(Application pptApp);

        //

        void AlignLeft(Application pptApp);
        void AlignRight(Application pptApp);
        void AlignTop(Application pptApp);
        void AlignBottom(Application pptApp);
    }
}
