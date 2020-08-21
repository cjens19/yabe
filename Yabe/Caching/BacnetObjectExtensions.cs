using System.IO.BACnet;

namespace Yabe.Caching
{
    public static class BacnetObjectExtensions
    {
        public static CachedObject ToCachedObject(this BacnetObject bacnetObject, string name = "")
        {
            var cachedObject = new CachedObject
            {
                Type = bacnetObject.Type,
                Exportable = bacnetObject.Exportable,
                Instance = bacnetObject.Instance,
                Name = name
            };
            
            return cachedObject;
        }
    }
}
