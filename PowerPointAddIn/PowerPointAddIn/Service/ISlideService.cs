using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn
{
    public interface ISlideService
    {
        // slide
        void AddSlide(Application pptApp);
        void RemoveSlide(Application pptApp);

        // font
        void FontSizePlus(Application pptApp);
        void FontSizeMinus(Application pptApp);
        // text
        void CopyText(Application pptApp);
        void PasteText(Application pptApp);

        // position
        void CopyPosition(Application pptApp);
        void PastePosition(Application pptApp);

        // align
        void AlignLeft(Application pptApp);
        void AlignRight(Application pptApp);
        void AlignTop(Application pptApp);
        void AlignBottom(Application pptApp);

        // status
        void checkShapesStatus(Application pptApp, Selection Sel);
    }
}
