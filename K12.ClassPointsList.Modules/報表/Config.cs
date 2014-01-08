using System;
using System.Collections.Generic;
using System.Text;
using FISCA.DSAUtil;
using System.ComponentModel;
using System.Xml;

namespace K12ClassPointsList.Modules
{
    public class Config
    {

        public const string LIST_ABSENCE = "GetAbsenceList";
        public const string LIST_ABSENCE_NUMBER = "36";
        [FISCA.Authentication.AutoRetryOnWebException()]
        public static DSResponse GetAbsenceList()
        {
            string serviceName = "GetAbsenceList";
            if (DataCacheManager.Get(serviceName) == null)
            {
                DSRequest request = new DSRequest();
                DSXmlHelper helper = new DSXmlHelper("GetAbsenceListRequest");
                helper.AddElement("Field");
                helper.AddElement("Field", "All");
                request.SetContent(helper);
                DSResponse dsrsp = FISCA.Authentication.DSAServices.CallService("SmartSchool.Others.GetAbsenceList", request);
                DataCacheManager.Add(serviceName, dsrsp);
            }
            return DataCacheManager.Get(serviceName);
        }

        public const string LIST_PERIODS = "GetPeriodList";
        public const string LIST_PERIODS_NUMBER = "35";
        [FISCA.Authentication.AutoRetryOnWebException()]
        public static DSResponse GetPeriodList()
        {
            string serviceName = LIST_PERIODS;
            if (DataCacheManager.Get(serviceName) == null)
            {
                DSRequest request = new DSRequest();
                DSXmlHelper helper = new DSXmlHelper("GetPeriodListRequest");
                helper.AddElement("Field");
                helper.AddElement("Field", "All");
                request.SetContent(helper);
                DSResponse dsrsp = FISCA.Authentication.DSAServices.CallService("SmartSchool.Others.GetPeriodList", request);
                DataCacheManager.Add(serviceName, dsrsp);
            }
            return DataCacheManager.Get(serviceName);
        }
    }

    public static class DataCacheManager
    {
        private static Dictionary<string, DSResponse> _dataManager = new Dictionary<string, DSResponse>();

        public static void Add(string name, DSResponse dsrsp)
        {
            lock (_dataManager)
            {
                if (!_dataManager.ContainsKey(name))
                    _dataManager.Add(name, dsrsp);
                else
                    _dataManager[name] = dsrsp;
            }
        }

        public static DSResponse Get(string name)
        {
            lock (_dataManager)
            {
                if (_dataManager.ContainsKey(name))
                    return _dataManager[name];
                return null;
            }
        }

        public static bool Contains(string serviceName)
        {
            return _dataManager.ContainsKey(serviceName);
        }

        public static void Remove(string serviceName)
        {
            if (_dataManager.ContainsKey(serviceName))
            {
                _dataManager.Remove(serviceName);
            }
        }

        public static void Set(string name, DSResponse dsrsp)
        {
            _dataManager[name] = dsrsp;
        }
    }
}
