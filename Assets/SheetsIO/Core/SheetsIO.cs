using System;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace SheetsIO
{
    public sealed class SheetsIO : IDisposable
    {
        internal const string FirstCell = "B2";
        internal const int MaxArrayElements = 100;
        internal const int A1LettersCount = 26;
        
        internal delegate bool ReadObjectDelegate(IOPointer p, out object o);

        readonly SheetsService service;
        readonly IValueSerializer serializer;

        /// <summary>Read the data from a spreadsheet and deserialize it to object of type T.</summary>
        /// <param name="spreadsheet">Spreadsheet ID.</param>
        /// <param name="sheet">Specify the name of the sheet, if the spreadsheet contains multiple sheets of a type.</param>
        /// <typeparam name="T">Type of object to read from the spreadsheet data.</typeparam>
        /// <returns>Object of type T.</returns>
        public async Task<T> ReadAsync<T>(string spreadsheet, string sheet = "") {
            var spreadsheets = await service.GetSpreadsheetAsync(spreadsheet);
            var context = new ReadContext(spreadsheets.GetSheetsList(), serializer);
            var meta = typeof(T).GetIOAttribute();
            if (!context.TryReadType(meta, meta.AppendNamePart(sheet), out var result))
                throw new Exception("Can't parse the requested object. Some required sheets are missing in the provided spreadsheet");

            var valueRanges = await service.GetValueRanges(spreadsheet, context.Ranges);
            if (!valueRanges.All(context.TryApplyRange))
                throw new Exception("Failed to assemble the object.");
            return (T)result;
        }

        /// <summary>Write a serialized object of type T to the spreadsheet.</summary>
        /// <param name="obj">Object to serialize.</param>
        /// <param name="spreadsheet">Spreadsheet ID.</param>
        /// <param name="sheet">Specify the name of the sheet, if the spreadsheet contains multiple sheets of a type.</param>
        /// <typeparam name="T">Type of serialized object.</typeparam>
        public Task<bool> WriteAsync<T>(T obj, string spreadsheet, string sheet = "") {
            var meta = typeof(T).GetIOAttribute();
            var ranges = new WriteContext(meta, meta.AppendNamePart(sheet), obj, serializer).ValueRanges;
            return service.WriteRangesAsync(spreadsheet, ranges);
        }

        /// <summary>Creates a new ready-to-use instance of the serializer.</summary>
        /// <param name="initializer"> An initializer for Google Sheets service. </param>
        /// <param name="valueSerializer"> Replaces the default serialization strategy. </param>
        public SheetsIO(BaseClientService.Initializer initializer, IValueSerializer valueSerializer = null) {
            // If you got an issue with missing System.IO.Compression lib, set GZipEnabled = FALSE in your initializer.
            service    = new SheetsService(initializer ?? throw new ArgumentException("SheetsService can't be null"));
            serializer = valueSerializer ?? new DefaultValueSerializer();
        }

        public void Dispose() => service.Dispose();
    }
}
