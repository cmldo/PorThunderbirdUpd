using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PorThunderbirdUpd
{
    public class MainForm : Form
    {
        private Button updateButton;
        private ProgressBar progressBar;
        private TextBox logBox;

        private readonly string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        private readonly string logFile;
        private readonly string versionLogFile;
        private readonly string sevenZipExe;

        private readonly string portableDir;
        private readonly string desktopVersionFile;
        private readonly string portableVersionFile;

        public MainForm()
        {
            this.Text = "Thunderbird Updater";
            this.Width = 600;
            this.Height = 400;

            updateButton = new Button() { Text = "Update Thunderbird Portable", Left = 20, Top = 20, Width = 200, Height = 40 };
            updateButton.Click += async (_, __) => await UpdateThunderbirdAsync();

            progressBar = new ProgressBar() { Left = 20, Top = 70, Width = 540, Height = 25 };
            logBox = new TextBox() { Left = 20, Top = 110, Width = 540, Height = 220, Multiline = true, ScrollBars = ScrollBars.Vertical };

            this.Controls.Add(updateButton);
            this.Controls.Add(progressBar);
            this.Controls.Add(logBox);

            logFile = Path.Combine(baseDir, "update.log");
            versionLogFile = Path.Combine(baseDir, "version.log");
            sevenZipExe = Path.Combine(baseDir, "7zr.exe");

            portableDir = Path.Combine(baseDir, "thunderbird-portable");
            desktopVersionFile = Path.Combine(portableDir, "installed_desktop_version.txt");
            portableVersionFile = Path.Combine(portableDir, "installed_portable_version.txt");

            this.Load += async (_, __) => await CheckIfUpdateNeededAsync();
        }

        private async Task CheckIfUpdateNeededAsync()
        {
            try
            {
                var desktopUpdate = await CheckThunderbirdUpdateNeededAsync(desktopVersionFile);
                var portableUpdate = await CheckLatestPortableAsync(portableVersionFile);

                if (!desktopUpdate.updateNeeded && !portableUpdate.updateNeeded)
                {
                    LogAction("All components are up to date.");
                    updateButton.Enabled = false;
                }
                else
                {
                    if (desktopUpdate.updateNeeded)
                        LogAction($"thunderbird-desktop update needed: current {desktopUpdate.currentVersion ?? "none"} → latest {desktopUpdate.latestVersion}");
                    if (portableUpdate.updateNeeded)
                        LogAction($"thunderbird-portable update needed: current {portableUpdate.version ?? "none"} → latest {portableUpdate.version}");
                    updateButton.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                LogAction($"Error checking versions: {ex.Message}");
            }
        }

        private async Task<(bool updateNeeded, string latestVersion, string currentVersion)> CheckThunderbirdUpdateNeededAsync(string installedVersionFile)
        {
            string currentVersion = File.Exists(installedVersionFile)
                ? File.ReadAllText(installedVersionFile).Trim()
                : null;

            string latestVersion = await GetLatestThunderbirdVersionAsync();

            bool updateNeeded = currentVersion != latestVersion;
            return (updateNeeded, latestVersion, currentVersion);
        }

        private async Task<string> GetLatestThunderbirdVersionAsync()
        {
            using var handler = new HttpClientHandler { AllowAutoRedirect = false };
            using var client = new HttpClient(handler);
            client.DefaultRequestHeaders.UserAgent.ParseAdd("ThunderbirdUpdater");

            string url = "https://download.mozilla.org/?product=thunderbird-latest-SSL&os=win64&lang=en-US";
            using var response = await client.GetAsync(url);

            if (response.StatusCode != System.Net.HttpStatusCode.Found &&
                response.StatusCode != System.Net.HttpStatusCode.Redirect)
                throw new Exception("Unexpected response from Mozilla download endpoint.");

            string location = response.Headers.Location.ToString();
            var match = Regex.Match(location, @"releases/([\d\.]+)/");

            if (!match.Success)
                throw new Exception("Unable to parse Thunderbird version from redirect.");

            return match.Groups[1].Value;
        }

        private async Task<(string version, string url, bool updateNeeded, string currentVersion)> CheckLatestPortableAsync(string installedVersionFile)
        {
            string currentVersion = File.Exists(installedVersionFile)
                ? File.ReadAllText(installedVersionFile).Trim()
                : null;

            var (version, url) = await GetLatestRelease("portapps", "stormhen-portable", "*.7z");
            bool updateNeeded = currentVersion != version;

            return (version, url, updateNeeded, currentVersion);
        }

        private async Task<(string version, string url)> GetLatestRelease(string owner, string repo, string assetPattern)
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.UserAgent.ParseAdd("ThunderbirdUpdater");

            string apiUrl = $"https://api.github.com/repos/{owner}/{repo}/releases/latest";
            LogAction($"Fetching latest release for {repo}...");
            var response = await client.GetAsync(apiUrl);
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();

            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            string tagName = root.GetProperty("tag_name").GetString();
            foreach (var asset in root.GetProperty("assets").EnumerateArray())
            {
                string name = asset.GetProperty("name").GetString();
                if (Regex.IsMatch(name, assetPattern.Replace("*", ".*")))
                {
                    string url = asset.GetProperty("browser_download_url").GetString();
                    LogAction($"Latest {repo} asset found: {name}");
                    return (tagName, url);
                }
            }

            throw new Exception($"No asset matching {assetPattern} found for {repo}");
        }

        private async Task UpdateThunderbirdAsync()
        {
            updateButton.Enabled = false;
            try
            {
                LogAction("Starting update...");

                await Ensure7zExistsAsync();

                // 1️⃣ Check latest versions
                var desktopUpdate = await CheckThunderbirdUpdateNeededAsync(desktopVersionFile);
                var portableUpdate = await GetLatestRelease("portapps", "stormhen-portable", "*.7z");

                LogAction($"Thunderbird Desktop: {desktopUpdate.latestVersion}");
                LogAction($"Thunderbird Portable: {portableUpdate.version}");

                // 2️⃣ Download stormhen-portable first
                string portableFile = Path.Combine(baseDir, Path.GetFileName(portableUpdate.url));
                await DownloadFileAsync(portableUpdate.url, portableFile);

                // 3️⃣ Download Thunderbird setup
                string desktopFile = Path.Combine(baseDir, $"Thunderbird Setup {desktopUpdate.latestVersion}.exe");
                await DownloadFileAsync($"https://download.mozilla.org/?product=thunderbird-{desktopUpdate.latestVersion}-SSL&os=win64&lang=en-US", desktopFile);

                // 4️⃣ Extract stormhen-portable into thunderbird-portable
                if (!Directory.Exists(portableDir))
                    Directory.CreateDirectory(portableDir);

                await Run7zExtract(portableFile, portableDir);

                // 5️⃣ Delete old app folder
                string appDir = Path.Combine(portableDir, "app");
                if (Directory.Exists(appDir))
                    Directory.Delete(appDir, true);
                LogAction($"Deleted old app folder: {appDir}");
                Directory.CreateDirectory(appDir);
                LogAction($"Created app folder: {appDir}");

                // 6️⃣ Extract Thunderbird .exe to temp folder
                string tempExtractDir = Path.Combine(baseDir, "tmp_extracted");
                if (Directory.Exists(tempExtractDir))
                    Directory.Delete(tempExtractDir, true);
                Directory.CreateDirectory(tempExtractDir);

                await Run7zExtract(desktopFile, tempExtractDir);

                // 7️⃣ Copy only \core into thunderbird-portable\app
                string coreDir = Path.Combine(tempExtractDir, "core");
                if (Directory.Exists(coreDir))
                    CopyDirectory(coreDir, appDir);

                Directory.Delete(tempExtractDir, true);

                // 8️⃣ Save installed versions
                File.WriteAllText(desktopVersionFile, desktopUpdate.latestVersion);
                File.WriteAllText(portableVersionFile, portableUpdate.version);

                // 9️⃣ Cleanup downloads
                File.Delete(desktopFile);
                File.Delete(portableFile);

                LogAction("Update completed successfully.");
            }
            catch (Exception ex)
            {
                LogAction($"Error: {ex.Message}");
            }
            finally
            {
                updateButton.Enabled = true;
                progressBar.Invoke(() => progressBar.Value = 0);
                progressBar.Invoke(() => progressBar.Style = ProgressBarStyle.Blocks);
            }
        }

        private void CopyDirectory(string sourceDir, string targetDir)
        {
            foreach (var dir in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
                Directory.CreateDirectory(dir.Replace(sourceDir, targetDir));

            foreach (var file in Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories))
                File.Copy(file, file.Replace(sourceDir, targetDir), true);
        }

        private async Task Ensure7zExistsAsync()
        {
            if (File.Exists(sevenZipExe))
            {
                LogAction("7zr.exe is already present.");
                return;
            }

            LogAction("7zr.exe not found. Downloading from official site...");
            using var client = new HttpClient();
            using var response = await client.GetAsync("https://www.7-zip.org/a/7zr.exe", HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode();

            using var stream = await response.Content.ReadAsStreamAsync();
            using var fileStream = File.Create(sevenZipExe);
            await stream.CopyToAsync(fileStream);

            LogAction("7zr.exe downloaded successfully.");
        }

        private async Task DownloadFileAsync(string url, string destination)
        {
            LogAction($"Downloading {Path.GetFileName(destination)}...");

            using var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = true });
            using var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
            response.EnsureSuccessStatusCode();

            var totalBytes = response.Content.Headers.ContentLength ?? -1L;
            var canReportProgress = totalBytes != -1;

            using var stream = await response.Content.ReadAsStreamAsync();
            using var fileStream = File.Create(destination);

            var buffer = new byte[8192];
            long totalRead = 0;
            int read;
            while ((read = await stream.ReadAsync(buffer, 0, buffer.Length)) > 0)
            {
                await fileStream.WriteAsync(buffer, 0, read);
                totalRead += read;

                if (canReportProgress)
                {
                    int percent = (int)((totalRead * 100) / totalBytes);
                    progressBar.Invoke(() => progressBar.Value = percent);
                }
            }

            LogAction($"Downloaded {Path.GetFileName(destination)}.");
            progressBar.Invoke(() => progressBar.Value = 0);
        }

        private async Task Run7zExtract(string archiveFile, string outputDir)
        {
            if (!File.Exists(sevenZipExe))
                throw new Exception("7zr.exe not found in application folder.");

            LogAction($"Extracting {Path.GetFileName(archiveFile)} to {outputDir}...");
            progressBar.Invoke(() => progressBar.Style = ProgressBarStyle.Marquee);

            var psi = new ProcessStartInfo
            {
                FileName = sevenZipExe,
                Arguments = $"x \"{archiveFile}\" -o\"{outputDir}\" -y",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            var process = Process.Start(psi);
            string output = await process.StandardOutput.ReadToEndAsync();
            string err = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();

            progressBar.Invoke(() => progressBar.Style = ProgressBarStyle.Blocks);
            LogAction(output);
            if (!string.IsNullOrEmpty(err))
                LogAction(err);
        }

        private void LogAction(string message)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            logBox.AppendText(line + Environment.NewLine);
            File.AppendAllText(logFile, line + Environment.NewLine);
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}

