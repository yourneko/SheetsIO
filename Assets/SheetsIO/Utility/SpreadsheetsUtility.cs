using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Http;
using Google.Apis.Requests;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util;

namespace SheetsIO
{
    static class SpreadsheetsUtility
    {
        public static async Task<bool> WriteRangesAsync(this SheetsService service, string spreadsheet, IList<ValueRange> values) {
            var hashset = new HashSet<string>(values.Select(range => range.Range.GetSheetName()));
            bool hasRequiredSheets = await service.CreateSheetsAsync(spreadsheet, hashset);
            if (!hasRequiredSheets)
                return false;

            var result = await service.Spreadsheets.Values.BatchUpdate(UpdateRequest(values), spreadsheet).AddBackOffHandler().ExecuteAsync();
            return result.TotalUpdatedCells > 0;
        }

        public static async Task<Spreadsheet> GetSpreadsheetAsync(this SheetsService service, string spreadsheetID) =>
            await service.Spreadsheets.Get(spreadsheetID).AddBackOffHandler().ExecuteAsync();

        public static async Task<IList<ValueRange>> GetValueRanges(this SheetsService service, string spreadsheet, IEnumerable<string> ranges) {
            var request = service.Spreadsheets.Values.BatchGet(spreadsheet);
            request.Ranges         = ranges.ToArray();
            request.MajorDimension = SpreadsheetsResource.ValuesResource.BatchGetRequest.MajorDimensionEnum.COLUMNS;
            var result = await request.AddBackOffHandler().ExecuteAsync();
            return result.ValueRanges;
        }

        static async Task<bool> CreateSheetsAsync(this SheetsService service, string spreadsheet, IEnumerable<string> requiredSheets) {
            var spreadsheets = await service.GetSpreadsheetAsync(spreadsheet);
            string[] sheetsToCreate = requiredSheets.Except(spreadsheets.GetSheetsList()).ToArray();
            if (sheetsToCreate.Length == 0) return true;

            var result = await service.Spreadsheets.BatchUpdate(AddSheet(sheetsToCreate), spreadsheet).AddBackOffHandler().ExecuteAsync();
            return result.Replies.All(reply => reply.AddSheet.Properties != null);
        }

        static ClientServiceRequest<T> AddBackOffHandler<T>(this ClientServiceRequest<T> request, BackOffHandler handler = null) {
            request.AddExceptionHandler(handler ?? new BackOffHandler(new ExponentialBackOff(TimeSpan.FromMilliseconds(250), 5)));
            return request;
        }

        public static IEnumerable<string> GetSheetsList(this Spreadsheet spreadsheet) => spreadsheet.Sheets.Select(sheet => sheet.Properties.Title).ToArray();

        static BatchUpdateSpreadsheetRequest AddSheet(IEnumerable<string> list) =>
            new BatchUpdateSpreadsheetRequest {Requests = list.Select(AddSheet).ToList()};

        static Request AddSheet(string name) => new Request {AddSheet = new AddSheetRequest {Properties = new SheetProperties {Title = name}}};
        static BatchUpdateValuesRequest UpdateRequest(IList<ValueRange> data) => new BatchUpdateValuesRequest {Data = data, ValueInputOption = "USER_ENTERED"};
    }
}
