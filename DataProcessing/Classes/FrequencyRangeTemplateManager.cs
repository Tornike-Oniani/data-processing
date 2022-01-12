using DataProcessing.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    public class FrequencyRangeTemplateManager
    {
        private string jsonFilePath;

        public FrequencyRangeTemplateManager()
        {
            this.jsonFilePath = Path.Combine(Environment.CurrentDirectory, "frequencyRangeTemplates.json");
        }

        public List<FrequencyRangeTemplate> GetFrequencyRangeTemplates() 
        {
            return JsonConvert.DeserializeObject<List<FrequencyRangeTemplate>>(File.ReadAllText(jsonFilePath));
        }
        public void SaveFrequencyRangeTemplates(List<FrequencyRangeTemplate> frequencyRangeTemplates)
        {
            File.WriteAllText(jsonFilePath, JsonConvert.SerializeObject(frequencyRangeTemplates));
        }
    }
}
