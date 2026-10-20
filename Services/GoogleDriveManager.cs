using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using SpreadsheetApp.Models;

namespace SpreadsheetApp.Services
{
    public class GoogleDriveManager
    {
        private const string AppDataFolderName = "root";
        private const string DefaultDataStoreFolderName = ".authdata";
        private const string ClientSecretsFileName = "client_secrets.json";
        
        private DriveService? _driveService;

        public async Task InitClientAsync()
        {
            if (_driveService != null) return;

            string secretsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ClientSecretsFileName);
            
            // For development, try looking up a few directories if not found
            if (!File.Exists(secretsPath))
            {
                var fallback = Path.Combine(Directory.GetCurrentDirectory(), ClientSecretsFileName);
                if (File.Exists(fallback)) secretsPath = fallback;
            }

            if (!File.Exists(secretsPath))
                throw new FileNotFoundException($"Файл з ключами не знайдено: {secretsPath}. Переконайтеся, що ви додали client_secrets.json.");

            var clientSecrets = (await GoogleClientSecrets.FromFileAsync(secretsPath)).Secrets;
            var dataStore = new FileDataStore(DefaultDataStoreFolderName);
            var scopes = new[] { DriveService.Scope.DriveFile };

            var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                clientSecrets,
                scopes,
                "user",
                CancellationToken.None,
                dataStore);

            _driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "SpreadsheetApp"
            });
        }

        public async Task SaveSpreadsheetAsync(Spreadsheet spreadsheet, string fileName = "spreadsheet.json")
        {
            try
            {
                await SaveSpreadsheetInternalAsync(spreadsheet, fileName);
            }
            catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException)
            {
                await ResetAuthAsync();
                await SaveSpreadsheetInternalAsync(spreadsheet, fileName);
            }
        }

        private async Task SaveSpreadsheetInternalAsync(Spreadsheet spreadsheet, string fileName)
        {
            if (_driveService == null) await InitClientAsync();

            var cellsData = spreadsheet.GetAllCells()
                .Where(c => !string.IsNullOrWhiteSpace(c.Expression))
                .Select(c => new { c.Name, c.Expression })
                .ToList();

            string jsonContent = JsonSerializer.Serialize(cellsData, new JsonSerializerOptions { WriteIndented = true });
            byte[] byteArray = Encoding.UTF8.GetBytes(jsonContent);

            using var stream = new MemoryStream(byteArray);

            // Check if file already exists
            var listRequest = _driveService!.Files.List();
            listRequest.Spaces = "drive";
            listRequest.Fields = "files(id, name)";
            var listResponse = await listRequest.ExecuteAsync();
            var existingFile = listResponse.Files?.FirstOrDefault(f => f.Name == fileName);

            if (existingFile != null)
            {
                // Update
                var updateRequest = _driveService.Files.Update(new Google.Apis.Drive.v3.Data.File(), existingFile.Id, stream, "application/json");
                await updateRequest.UploadAsync();
            }
            else
            {
                // Create
                var fileMetadata = new Google.Apis.Drive.v3.Data.File
                {
                    Name = fileName,
                    Parents = new List<string> { AppDataFolderName }
                };

                var createRequest = _driveService.Files.Create(fileMetadata, stream, "application/json");
                createRequest.Fields = "id";
                await createRequest.UploadAsync();
            }
        }

        public async Task LoadSpreadsheetAsync(Spreadsheet spreadsheet, string fileName = "spreadsheet.json")
        {
            try
            {
                await LoadSpreadsheetInternalAsync(spreadsheet, fileName);
            }
            catch (Google.Apis.Auth.OAuth2.Responses.TokenResponseException)
            {
                await ResetAuthAsync();
                await LoadSpreadsheetInternalAsync(spreadsheet, fileName);
            }
        }

        private async Task LoadSpreadsheetInternalAsync(Spreadsheet spreadsheet, string fileName)
        {
            if (_driveService == null) await InitClientAsync();

            var listRequest = _driveService!.Files.List();
            listRequest.Spaces = "drive";
            listRequest.Fields = "files(id, name)";
            var listResponse = await listRequest.ExecuteAsync();
            var existingFile = listResponse.Files?.FirstOrDefault(f => f.Name == fileName);

            if (existingFile == null)
                throw new Exception("Файл не знайдено на Google Drive.");

            var getRequest = _driveService.Files.Get(existingFile.Id);
            using var stream = new MemoryStream();
            await getRequest.DownloadAsync(stream);
            
            stream.Position = 0;
            using var reader = new StreamReader(stream, Encoding.UTF8);
            string jsonContent = await reader.ReadToEndAsync();

            var cellsData = JsonSerializer.Deserialize<List<CellData>>(jsonContent);
            if (cellsData != null)
            {
                // Очищаємо всі поточні клітинки перед завантаженням, щоб не було накладання старих даних
                foreach (var cell in spreadsheet.GetAllCells())
                {
                    cell.Expression = string.Empty;
                    cell.Value = null;
                    cell.Error = string.Empty;
                    cell.References.Clear();
                }

                foreach (var cellData in cellsData)
                {
                    var cell = spreadsheet.GetCell(cellData.Name);
                    if (cell != null)
                    {
                        cell.Expression = cellData.Expression;
                    }
                }
            }
        }

        private async Task ResetAuthAsync()
        {
            var dataStore = new FileDataStore(DefaultDataStoreFolderName);
            await dataStore.ClearAsync();
            _driveService = null;
            await InitClientAsync();
        }

        private class CellData
        {
            public string Name { get; set; } = string.Empty;
            public string Expression { get; set; } = string.Empty;
        }
    }
}

