using Unity;
namespace PowerPointAddIn
{
    public static class ContainerConfig
    {
        public static IUnityContainer RegisterServices()
        {
            var container = new UnityContainer();
            container.RegisterType<ISlideService, SlideService>();
            return container;
        }
    }
}
