using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO.BACnet;
using System.Linq;

namespace Yabe.Caching
{
    public class DeviceCache
    {
        private readonly ConcurrentDictionary<uint, IList<BacnetObject>> _cache;

        public DeviceCache()
        {
            _cache = new ConcurrentDictionary<uint, IList<BacnetObject>>();
        }

        public IList<BacnetObject> GetObjects(uint deviceId)
        {
            return _cache[deviceId];
        }

        //public (bool exists, bool checkedState) Exists(uint deviceId, BacnetObject bacnetObject)
        //{
        //    if (_cache.ContainsKey(deviceId) && _cache[deviceId].Any(p => p.Instance == bacnetObject.Instance && p.Type == bacnetObject.Type))
        //    {
        //        return (true,
        //            _cache[deviceId].Single(p => p.Instance == bacnetObject.Instance && p.Type == bacnetObject.Type)
        //                .Exportable);
        //    }
            
        //}

        public bool ObjectCheckedState(uint deviceId, BacnetObject bacnetObject)
        {
            if (bacnetObject.Type == BacnetObjectTypes.OBJECT_DEVICE) return true; //always want to have Device objects selected for export

            if (!_cache.ContainsKey(deviceId)) return false;

            if (_cache[deviceId].Any(p => p.Instance == bacnetObject.Instance && p.Type == bacnetObject.Type))
            {
                return _cache[deviceId].SingleOrDefault(p => p.Instance == bacnetObject.Instance && p.Type == bacnetObject.Type)
                    .Exportable;
            }

            return false;
        }

        public int NumberOfExportableObjects(uint deviceId)
        {
            return _cache[deviceId].Count(p => p.Exportable);
        }

        public void AddOrUpdate(uint deviceId, BacnetObject bacnetObject)
        {
            if (_cache.ContainsKey(deviceId))
            {
                UpdateCache(deviceId, bacnetObject);
            }
            else
            {
                CreateCachedDevice(deviceId, bacnetObject);
            }
        }

        private void CreateCachedDevice(uint deviceId, BacnetObject bacnetObject)
        {
            _cache[deviceId] = new List<BacnetObject>
            {
                bacnetObject
            };
        }

        private void AddToCache(uint deviceId, BacnetObject bacnetObject)
        {
            _cache[deviceId].Add(bacnetObject);
        }

        private void UpdateCache(uint deviceId, BacnetObject bacnetObject)
        {
            var objects = _cache[deviceId];

            var exists = objects.Any(p => p.Instance == bacnetObject.Instance && p.Type == bacnetObject.Type);

            if (exists)
            {
                var toRemove = objects.Single(p => p.Instance == bacnetObject.Instance && p.Type == bacnetObject.Type);

                RemoveFromCache(deviceId, toRemove);
            }

            AddToCache(deviceId, bacnetObject);
        }

        private void RemoveFromCache(uint deviceId, BacnetObject bacnetObject)
        {
            _cache[deviceId].Remove(bacnetObject);
        }
    }
}
