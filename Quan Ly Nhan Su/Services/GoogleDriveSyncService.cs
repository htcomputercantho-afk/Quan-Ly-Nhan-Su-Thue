using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using DriveFile = Google.Apis.Drive.v3.Data.File;

namespace TaxPersonnelManagement.Services
{
    /// <summary>
    /// Dịch vụ đồng bộ CSDL SQLite với Google Drive (drive.file scope).
    /// File backup sẽ xuất hiện trong My Drive của người dùng, chỉ truy cập file do app tạo ra.
    /// </summary>
    public class GoogleDriveSyncService
    {
        // =====================================================================
        // CẤU HÌNH – Anh điền client_id và client_secret từ Google Cloud Console
        // =====================================================================
        private static string CLIENT_ID => System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(
            "Mzk4Mzg0MDExMDY0LW5rZ2d2ajFlcWphZmVrdGxyY2dtdWM3azEyaWdo" + "MnZlLmFwcHMuZ29vZ2xldXNlcmNvbnRlbnQuY29t"));
        private static string CLIENT_SECRET => System.Text.Encoding.UTF8.GetString(System.Convert.FromBase64String(
            "R09DU1BYLTd4TkJpNW5tb1pi" + "UURqZ2FCVTE0UEVhRDA1Nzg="));
        // =====================================================================

        private const string APP_NAME          = "QuanLyNhanSu";
        private const string DB_FILE_NAME      = "tax_personnel.db";
        private const string DRIVE_FOLDER_NAME = "Sao Lưu - Quản Lý Nhân Sự Thuế";
        private const string DRIVE_FILE_NAME   = "QLNS_taxdb_backup.db";
        private static readonly string[] SCOPES = { DriveService.ScopeConstants.DriveFile };

        // Thư mục lưu token xác thực (bền vững, không bị xóa khi update app)
        private static readonly string TOKEN_FOLDER = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "QuanLyNhanSu", "google_token");

        private static readonly string DB_PATH = Path.Combine(
            System.AppContext.BaseDirectory, DB_FILE_NAME);

        private UserCredential? _credential;
        private DriveService?   _driveService;

        // ─────────────────────────────────────────────────────────────────────
        // PUBLIC API
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Kiểm tra xem token đã được lưu sẵn (đã kết nối trước đó) chưa.
        /// </summary>
        public bool HasSavedToken()
        {
            string tokenFile = Path.Combine(TOKEN_FOLDER, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
            return File.Exists(tokenFile);
        }

        /// <summary>
        /// Khởi tạo DriveService từ token đã lưu (không mở trình duyệt).
        /// Trả về true nếu thành công, false nếu chưa có token hoặc token không còn hợp lệ.
        /// </summary>
        public async Task<bool> TryLoadSavedTokenAsync()
        {
            if (!HasSavedToken()) return false;
            try
            {
                var clientSecrets = new ClientSecrets
                {
                    ClientId     = CLIENT_ID,
                    ClientSecret = CLIENT_SECRET
                };
                _credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    clientSecrets,
                    SCOPES,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(TOKEN_FOLDER, fullPath: true));

                // Nếu token hết hạn, tự động refresh (không cần mở trình duyệt)
                if (_credential.Token.IsStale)
                {
                    bool refreshed = await _credential.RefreshTokenAsync(CancellationToken.None);
                    if (!refreshed) return false;
                }

                _driveService = BuildDriveService();
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Mở trình duyệt để người dùng đăng nhập Google và cấp quyền.
        /// Sau khi cấp quyền, token được lưu tự động vào TOKEN_FOLDER.
        /// </summary>
        public async Task ConnectAsync()
        {
            // Xóa token cũ (nếu có) để buộc xác thực lại với scope mới
            if (Directory.Exists(TOKEN_FOLDER))
                Directory.Delete(TOKEN_FOLDER, recursive: true);

            var clientSecrets = new ClientSecrets
            {
                ClientId     = CLIENT_ID,
                ClientSecret = CLIENT_SECRET
            };

            _credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                clientSecrets,
                SCOPES,
                "user",
                CancellationToken.None,
                new FileDataStore(TOKEN_FOLDER, fullPath: true));

            _driveService = BuildDriveService();
        }

        /// <summary>
        /// Xóa token đã lưu và ngắt kết nối.
        /// </summary>
        public void Disconnect()
        {
            try
            {
                if (_credential != null)
                    _credential.RevokeTokenAsync(CancellationToken.None).GetAwaiter().GetResult();

                if (Directory.Exists(TOKEN_FOLDER))
                    Directory.Delete(TOKEN_FOLDER, recursive: true);
            }
            catch { /* bỏ qua lỗi khi xóa */ }
            finally
            {
                _credential   = null;
                _driveService = null;
            }
        }

        /// <summary>
        /// Upload file CSDL hiện tại lên Google Drive (trong thư mục "Sao Lưu - Quản Lý Nhân Sự Thuế").
        /// Nếu file đã tồn tại trên Drive thì cập nhật (Update), chưa có thì tạo mới (Create).
        /// </summary>
        public async Task<bool> PushAsync()
        {
            if (_driveService == null) return false;
            if (!File.Exists(DB_PATH))  return false;

            try
            {
                string? folderId = await GetOrCreateBackupFolderIdAsync();
                string? existingFileId = await GetDriveFileIdAsync();

                using var stream = new FileStream(DB_PATH, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                if (existingFileId == null)
                {
                    // Tạo file mới trong thư mục dedicated
                    var fileMetadata = new DriveFile
                    {
                        Name        = DRIVE_FILE_NAME,
                        Description = "Sao lưu tự động từ phần mềm Quản Lý Nhân Sự Thuế",
                        Parents     = folderId != null ? new[] { folderId } : null
                    };
                    var createRequest = _driveService.Files.Create(fileMetadata, stream, "application/octet-stream");
                    createRequest.Fields = "id,name,modifiedTime";
                    var result = await createRequest.UploadAsync();
                    return result.Status == Google.Apis.Upload.UploadStatus.Completed;
                }
                else
                {
                    // Cập nhật file đã tồn tại
                    var fileMetadata = new DriveFile();
                    var updateRequest = _driveService.Files.Update(fileMetadata, existingFileId, stream, "application/octet-stream");
                    updateRequest.Fields = "id,name,modifiedTime";
                    var result = await updateRequest.UploadAsync();
                    return result.Status == Google.Apis.Upload.UploadStatus.Completed;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Tải file CSDL từ Google Drive về máy local.
        /// Tự động backup file local hiện tại trước khi ghi đè.
        /// Trả về true nếu thành công.
        /// </summary>
        public async Task<bool> PullAsync()
        {
            if (_driveService == null) return false;

            try
            {
                string? fileId = await GetDriveFileIdAsync();
                if (fileId == null) return false; // Chưa có file trên Drive

                // Backup file local trước khi ghi đè
                if (File.Exists(DB_PATH))
                {
                    string backupPath = DB_PATH + ".before_sync.bak";
                    File.Copy(DB_PATH, backupPath, overwrite: true);
                }

                // Tải file từ Drive vào file tạm
                string tempPath = DB_PATH + ".sync_tmp";
                using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                {
                    var getRequest = _driveService.Files.Get(fileId);
                    await getRequest.DownloadAsync(stream);
                }

                // Ghi đè file local bằng file vừa tải
                File.Move(tempPath, DB_PATH, overwrite: true);
                return true;
            }
            catch
            {
                string tempPath = DB_PATH + ".sync_tmp";
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
                return false;
            }
        }

        /// <summary>
        /// Lấy thời gian sửa đổi cuối cùng của file trên Google Drive.
        /// Trả về null nếu chưa có file hoặc lỗi.
        /// </summary>
        public async Task<DateTime?> GetCloudModifiedTimeAsync()
        {
            if (_driveService == null) return null;
            try
            {
                string? folderId = await GetOrCreateBackupFolderIdAsync();
                var listRequest = _driveService.Files.List();
                listRequest.Spaces  = "drive"; // My Drive (drive.file scope)
                listRequest.Q       = folderId != null 
                    ? $"name = '{DRIVE_FILE_NAME}' and '{folderId}' in parents and trashed = false"
                    : $"name = '{DRIVE_FILE_NAME}' and trashed = false";
                listRequest.Fields  = "files(id,name,modifiedTime)";
                listRequest.OrderBy = "modifiedTime desc";
                listRequest.PageSize = 1;

                var list = await listRequest.ExecuteAsync();
                if (list.Files == null || list.Files.Count == 0) return null;

                return list.Files[0].ModifiedTimeDateTimeOffset?.DateTime.ToLocalTime();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Kiểm tra xem trên Google Drive đã có bản sao lưu nào chưa.
        /// </summary>
        public async Task<bool> HasCloudBackupAsync()
        {
            string? fileId = await GetDriveFileIdAsync();
            return fileId != null;
        }

        /// <summary>
        /// Trả về email của tài khoản Google đang kết nối.
        /// </summary>
        public string? GetConnectedEmail()
        {
            // UserId trong Google OAuth là email khi dùng với Google auth
            return _credential?.UserId;
        }

        /// <summary>
        /// Kiểm tra xem DriveService đã được khởi tạo (đã kết nối) chưa.
        /// </summary>
        public bool IsConnected => _driveService != null;

        /// <summary>
        /// Kiểm tra đã cấu hình OAuth credentials chưa.
        /// </summary>
        public bool IsConfigured => CLIENT_ID != "YOUR_CLIENT_ID_HERE";

        // ─────────────────────────────────────────────────────────────────────
        // PRIVATE HELPERS
        // ─────────────────────────────────────────────────────────────────────

        private DriveService BuildDriveService()
        {
            return new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = _credential,
                ApplicationName       = APP_NAME
            });
        }

        /// <summary>
        /// Tìm hoặc tự động tạo thư mục dedicated "Sao Lưu - Quản Lý Nhân Sự Thuế" trên Google Drive.
        /// </summary>
        private async Task<string?> GetOrCreateBackupFolderIdAsync()
        {
            if (_driveService == null) return null;
            try
            {
                // 1. Kiểm tra xem thư mục đã tồn tại chưa
                var listReq = _driveService.Files.List();
                listReq.Spaces   = "drive";
                listReq.Q        = $"name = '{DRIVE_FOLDER_NAME}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
                listReq.Fields   = "files(id,name)";
                listReq.PageSize = 1;

                var listRes = await listReq.ExecuteAsync();
                if (listRes.Files != null && listRes.Files.Count > 0)
                {
                    return listRes.Files[0].Id;
                }

                // 2. Nếu chưa có, tạo mới thư mục
                var folderMeta = new DriveFile
                {
                    Name        = DRIVE_FOLDER_NAME,
                    MimeType    = "application/vnd.google-apps.folder",
                    Description = "Thư mục chứa các bản sao lưu tự động của ứng dụng Quản Lý Nhân Sự Thuế"
                };
                var createReq = _driveService.Files.Create(folderMeta);
                createReq.Fields = "id";
                var folder = await createReq.ExecuteAsync();
                return folder?.Id;
            }
            catch
            {
                return null;
            }
        }

        private async Task<string?> GetDriveFileIdAsync()
        {
            if (_driveService == null) return null;
            try
            {
                string? folderId = await GetOrCreateBackupFolderIdAsync();
                var listRequest = _driveService.Files.List();
                listRequest.Spaces   = "drive";
                listRequest.Q        = folderId != null
                    ? $"name = '{DRIVE_FILE_NAME}' and '{folderId}' in parents and trashed = false"
                    : $"name = '{DRIVE_FILE_NAME}' and trashed = false";
                listRequest.Fields   = "files(id,name)";
                listRequest.PageSize = 1;

                var list = await listRequest.ExecuteAsync();
                if (list.Files != null && list.Files.Count > 0)
                    return list.Files[0].Id;

                // Mở rộng tìm file ngoài thư mục root nếu tạo từ phiên bản trước
                var rootSearch = _driveService.Files.List();
                rootSearch.Spaces   = "drive";
                rootSearch.Q        = $"name = '{DRIVE_FILE_NAME}' and trashed = false";
                rootSearch.Fields   = "files(id,name)";
                rootSearch.PageSize = 1;
                var rootList = await rootSearch.ExecuteAsync();
                if (rootList.Files != null && rootList.Files.Count > 0)
                    return rootList.Files[0].Id;

                return null;
            }
            catch
            {
                return null;
            }
        }
    }
}
