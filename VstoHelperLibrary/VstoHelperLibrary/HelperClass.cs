using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace VstoHelperLibrary
{
    public class HelperClass
    {
        public static string AddParagraph(Request request)
        {
            // Initialization.  
            string receivedXml = string.Empty;

            // Posting.  
            using (var client = new HttpClient())
            {
                // Setting Base address.  
                client.BaseAddress = new Uri("https://localhost:7275/");

                var json = JsonConvert.SerializeObject(request);
                var data = new StringContent(json, Encoding.UTF8, "application/json");

                // HTTP POST  
                var response = client.PostAsync("api/word/addparagraph", data).Result;

                // Verification  
                if (response.IsSuccessStatusCode)
                {
                    // Reading Response.  
                    var result = response.Content.ReadAsStringAsync().Result;
                    receivedXml = JsonConvert.DeserializeObject<string>(result);
                }
            }

            return receivedXml;
        }
    }
}
