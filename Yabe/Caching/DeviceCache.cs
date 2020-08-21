using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO.BACnet;
using System.Linq;
using System.Text.Json.Serialization;

namespace Yabe.Caching
{
    public class DeviceCache
    {
        private readonly ConcurrentDictionary<uint, IList<CachedObject>> _cache;

        public DeviceCache()
        {
            _cache = new ConcurrentDictionary<uint, IList<CachedObject>>();
        }

        public ConcurrentDictionary<uint, IList<CachedObject>> GetCache()
        {
            return _cache;
        }

        public IList<CachedObject> GetObjects(uint deviceId)
        {
            return _cache[deviceId];
        }

        public IList<uint> GetCachedDeviceIds()
        {
            return new List<uint>(_cache.Keys);
        }

        public bool ObjectCheckedState(uint deviceId, CachedObject cachedObject)
        {
            if (cachedObject.Type == BacnetObjectTypes.OBJECT_DEVICE) return true; //always want to have Device objects selected for export

            if (!_cache.ContainsKey(deviceId)) return false;

            if (_cache[deviceId].Any(p => p.Instance == cachedObject.Instance && p.Type == cachedObject.Type))
            {
                return _cache[deviceId].Single(p => p.Instance == cachedObject.Instance && p.Type == cachedObject.Type)
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

        public void UpdateCheckedState(uint deviceId, BacnetObjectTypes type, uint instance, bool state)
        {
            if (!_cache.ContainsKey(deviceId)) return;
            var cachedObject = _cache[deviceId].SingleOrDefault(p => p.Type == type && p.Instance == instance);

            if (cachedObject != null)
                cachedObject.Exportable = state;
        }

        public void AddOrUpdate(uint deviceId, CachedObject cachedObject)
        {
            if (_cache.ContainsKey(deviceId))
            {
                UpdateCache(deviceId, cachedObject);
            }
            else
            {
                CreateCachedDevice(deviceId, cachedObject);
            }
        }

        private void CreateCachedDevice(uint deviceId, CachedObject cachedObject)
        {
            _cache[deviceId] = new List<CachedObject>
            {
                cachedObject
            };
        }

        private void AddToCache(uint deviceId, CachedObject cachedObject)
        {
            _cache[deviceId].Add(cachedObject);
        }

        private void UpdateCache(uint deviceId, CachedObject cachedObject)
        {
            var objects = _cache[deviceId];

            var exists = objects.Any(p => p.Instance == cachedObject.Instance && p.Type == cachedObject.Type);
            
            if (exists)
            {
                var existingObject = objects.Single(p => p.Instance == cachedObject.Instance && p.Type == cachedObject.Type);

                UpdateIfChanged(existingObject, cachedObject);
            }
            else
            {
                AddToCache(deviceId, cachedObject);
            }
        }

        private void UpdateIfChanged(CachedObject existingObject, CachedObject cachedObject)
        {
            existingObject.Type = cachedObject.Type;
            existingObject.Exportable = cachedObject.Exportable;
            existingObject.Name = cachedObject.Name;
            existingObject.Instance = cachedObject.Instance;
        }

        //private void RemoveFromCache(uint deviceId, CachedObject cachedObject)
        //{
        //    _cache[deviceId].Remove(cachedObject);
        //}
    }

    public class CachedDevice
    {
        public uint DeviceId { get; set; }
        public string Name { get; set; }
        public IList<CachedObject> CachedObjects { get; set; }

        public CachedDevice()
        {
            CachedObjects = new List<CachedObject>();
        }
    }
    public class CachedObject
    {
        public string Name { get; set; }
        public uint Instance { get; set; }
        public BacnetObjectTypes Type { get; set; }
        public bool Exportable { get; set; }
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
        public BacnetObjectTypes Type { get; set; }
        [JsonPropertyName("instance")]
        public uint Instance { get; set; }
        [JsonPropertyName("name")]
        public string Name { get; set; }

        public JsonDeviceObject(CachedObject cachedObject)
        {
            Type = cachedObject.Type;
            Instance = cachedObject.Instance;
            Name = cachedObject.Name;
        }
    }
}
