using System;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using UnityEngine;

namespace Example
{
    public class ExampleComponent : MonoBehaviour
    {
        [SerializeField] ExampleTargetComponent target;
        [SerializeField] string testSpreadsheetID;
        [SerializeField] ExampleData data = new ExampleData();
        [SerializeField] SuperclassData someSheetsData = new SuperclassData();
        [SerializeField] TextAsset key;

        SheetsIO.SheetsIO service;

        void Start() {
            var httpInit = GoogleCredential.FromJson(key.text).CreateScoped("https://spreadsheets.google.com/feeds", "https://docs.google.com/feeds");
            service = new SheetsIO.SheetsIO(new BaseClientService.Initializer {
                                                                                  HttpClientInitializer = httpInit,
                                                                                  ApplicationName       = "SheetsSerialization",
                                                                                  GZipEnabled = Application.isEditor ||
                                                                                                Application.platform != RuntimePlatform.Android
                                                                              });
        }

        public void WriteData() => InvokeSafe(Write(data));
        public void WriteSheets() => InvokeSafe(Write(someSheetsData));
        public void ReadData() => InvokeSafe(Read<ExampleData>(x => target.data = x));
        public void ReadSheets() => InvokeSafe(Read<SuperclassData>(x => target.someSheetsData = x));

        static async void InvokeSafe(Task task) {
            try {
                await task;
            }
            catch (Exception e) {
                Debug.LogError(e);
            }
        }

        async Task Write<T>(T obj) {
            var result = await service.WriteAsync(obj, testSpreadsheetID);
            print(result ? "Sheets were successfully updated." : "Write task failed");
        }

        async Task Read<T>(Action<T> callback, string sheet = null) {
            var result = await service.ReadAsync<T>(testSpreadsheetID, sheet ?? string.Empty);
            print(result is null ? "Failed." : "The requested object was successfully created.");
            callback.Invoke(result);
        }
    }
}
