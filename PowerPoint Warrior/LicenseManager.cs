using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
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
		public static Task<License> CheckLicenseAsync(string email, string licenseKey, CancellationToken cancel)
		{
			var client = new HttpClient();

			client.BaseAddress = new Uri("https://ppwarrior.azure-mobile.net/");
			// Add an Accept header for JSON format.
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
			// add application key
			client.DefaultRequestHeaders.Add("X-ZUMO-APPLICATION", APPLICATION_KEY);

			// set up license query
			var licenseInfo = new LicenseInfo { email = email, licenseKey = licenseKey };
			var json = JsonConvert.SerializeObject(licenseInfo);
			HttpContent content = new StringContent(json, Encoding.UTF8, "application/json");

			return client.PostAsync("api/checkLicense", content, cancel).ContinueWith(response =>
			{
				// when the task finished, we can dispose client
				client.Dispose();
				// return license if everything went ok and we found a license
				if (response.Status == TaskStatus.RanToCompletion && response.Result != null && response.Result.IsSuccessStatusCode)
				{
					var result = response.Result.Content.ReadAsStringAsync();
					return JsonConvert.DeserializeObject<License>(result.Result);
				}
				return null;
			});
		}
	}
}
