using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;

namespace VstoHelperLibrary
{
    public  class DocumentHelper
    {
        public static bool HasNullHeaderFooter(BADocument Document)
        {
            // Initialization.  
            bool receivedResult = false;

            // Posting.  
            using (var client = new HttpClient())
            {
                // Setting Base address.  
                client.BaseAddress = new Uri("https://localhost:7275/");

                var json = JsonConvert.SerializeObject(Document);
                var data = new StringContent(json, Encoding.UTF8, "application/json");

                // HTTP POST  
                var response = client.PostAsync("api/startup/HasNullHeaderFooter", data).Result;

                // Verification  
                if (response.IsSuccessStatusCode)
                {
                    // Reading Response.  
                    var result = response.Content.ReadAsStringAsync().Result;
                    receivedResult = JsonConvert.DeserializeObject<bool>(result);
                }
            }

            return receivedResult;
        }
    }
}
