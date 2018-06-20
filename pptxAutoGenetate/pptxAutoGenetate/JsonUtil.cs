using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pptxAutoGenetate
{
    class JsonUtil
    {
        public static ProjectItem getProjectItemFromJsonFile(string jsonFile)
        {
            
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;

            using (StreamReader sr = new StreamReader(jsonFile))
            using (JsonReader reader = new JsonTextReader(sr))
            {
                JObject jobject = (JObject)serializer.Deserialize(reader);
                return JsonConvert.DeserializeObject<ProjectItem>(jobject.ToString());
               
            }
        }
        public static Dictionary<string, Dictionary<string, decimal>> getLineStationDataFromJsonFile(string jsonFile)
        {
            Dictionary<string, Dictionary<string, decimal>> data = new Dictionary<string, Dictionary<string, decimal>>();
            JsonSerializer serializer = new JsonSerializer();
            serializer.NullValueHandling = NullValueHandling.Ignore;

            using (StreamReader sr = new StreamReader(jsonFile))
            using (JsonReader reader = new JsonTextReader(sr))
            {
                JObject jobject = (JObject)serializer.Deserialize(reader);
                foreach (JToken jt in jobject.Children())
                {
                    string areaDataKey = jt.Path;
                    JArray jarLines = JArray.Parse(jobject[areaDataKey].ToString());
                    for (int i = 0; i < jarLines.Count; i++)
                    {
                        Dictionary<string, decimal> tmpStationDic = new Dictionary<string, decimal>();
                        string lineKey = jarLines[i]["line"].ToString();
                        JArray jarStations = JArray.Parse(jarLines[i]["stations"].ToString());
                        for(int j=0;j<jarStations.Count;j++)
                        {
                            string stationKey = jarStations[j]["station"].ToString();
                            decimal value = Decimal.Parse(jarStations[j]["value"].ToString());
                            tmpStationDic.Add(stationKey, value);
                        }
                        data.Add(lineKey, tmpStationDic);
                    }
                }
            }
            return data;
        }

    }
}
