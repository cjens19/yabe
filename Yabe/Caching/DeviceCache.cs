using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO.BACnet;
using System.Linq;
using System.Text.Json.Serialization;

namespace Yabe.Caching
{
    public class DeviceCache
    {
        private readonly ConcurrentDictionary<uint, IList<BacnetObject>> _cache;

        public DeviceCache()
        {
            _cache = new ConcurrentDictionary<uint, IList<BacnetObject>>();
        }

        public ConcurrentDictionary<uint, IList<BacnetObject>> GetCache()
        {
            return _cache;
        }

        public IList<BacnetObject> GetObjects(uint deviceId)
        {
            return _cache[deviceId];
        }

        public IList<uint> GetCachedDeviceIds()
        {
            return new List<uint>(_cache.Keys);
        }

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

        public int NumberOfExportableObjectsForDevice(uint deviceId)
        {
            return _cache[deviceId].Count(p => p.Exportable);
        }

        public int NumberOfExportableObjects()
        {
            return _cache.Sum(entry => entry.Value.Count(p => p.Exportable));
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

    public class JsonDevice
    {
        [JsonPropertyName("deviceId")]
        public string DeviceId { get; set; }

        [JsonPropertyName("objects")]
        public IList<JsonDeviceObject> Objects { get; set; }

        public JsonDevice(uint deviceId)
        {
            DeviceId = deviceId.ToString();
            Objects = new List<JsonDeviceObject>();
        }
    }

    public class JsonDeviceObject
    {
        [JsonPropertyName("type")]
        public string Type { get; set; }
        [JsonPropertyName("instance")]
        public uint Instance { get; set; }

        public JsonDeviceObject(BacnetObject bacnetObject)
        {
            Type = bacnetObject.Type.ToString();
            Instance = bacnetObject.Instance;
        }
    }
}
