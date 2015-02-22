using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WarriorCommon
{
    public class License
    {
        public string Edition;
        public string Company;
        public DateTime ValidUntil;
    }
	public class LicenseInfo
	{
		public string email;
		public string licenseKey;
	}

	public static class LicenseManager
	{
		private const string APPLICATION_KEY = "VezvNeMAfhGkpCKoEraaHeTmbzNSFM47";
		public static async Task<License> CheckLicense(string email, string licenseKey)
		{
			using (var client = new HttpClient())
			{
				client.BaseAddress = new Uri("https://ppwarrior.azure-mobile.net/");
				// Add an Accept header for JSON format.
				client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
				// add application key
				client.DefaultRequestHeaders.Add("X-ZUMO-APPLICATION", APPLICATION_KEY);

				// set up license quesry
				var licenseInfo = new LicenseInfo { email = email, licenseKey = licenseKey };
				var json = JsonConvert.SerializeObject(licenseInfo);
				HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");

				HttpResponseMessage response = await client.PostAsync("api/checkLicense", content);

				if (response.IsSuccessStatusCode)
				{
					var result = response.Content.ReadAsStringAsync();
					return JsonConvert.DeserializeObject<License>(result.Result);
				}

				return null;
			}
		}
	}
}
