using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestKontrolProg
{
    internal static class TestKontrolUpdater
    {
        private const string ManifestUrl = "http://20.91.245.7/pdf-updates/test-kontrol/manifest.json";
        private const string ProductName = "Test Kontrol";
        private static readonly HttpClient Http = CreateHttpClient();
        private static Button _button;
        private static Form _owner;
        private static bool _busy;

        private static HttpClient CreateHttpClient()
        {
            var client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(5);
            client.DefaultRequestHeaders.UserAgent.ParseAdd("TestKontrolProg-Updater");
            return client;
        }

        public static void Initialize(Form owner, Button updateButton)
        {
            _owner = owner;
            _button = updateButton;
            _button.Text = "GÜNCELLE";
            _button.Click -= UpdateButton_Click;
            _button.Click += UpdateButton_Click;
            _owner.Shown -= Owner_Shown;
            _owner.Shown += Owner_Shown;
        }

        private static async void Owner_Shown(object sender, EventArgs e)
        {
            _owner.Shown -= Owner_Shown;
            await CheckAsync(false);
        }

        private static async void UpdateButton_Click(object sender, EventArgs e)
        {
            await CheckAsync(true);
        }

        private static async Task CheckAsync(bool interactive)
        {
            if (_busy) return;
            _busy = true;
            SetButton("KONTROL...", true);
            try
            {
                var manifest = await GetManifestAsync();
                var current = GetCurrentVersion();
                if (manifest == null || string.IsNullOrWhiteSpace(manifest.Version))
                {
                    SetButton("GÜNCELLE", false);
                    if (interactive) MessageBox.Show(_owner, "Güncelleme bilgisi alınamadı.", ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (CompareVersions(manifest.Version, current) <= 0)
                {
                    SetButton("GÜNCELLE", false);
                    if (interactive) MessageBox.Show(_owner, "Test Kontrol zaten güncel.", ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SetButton("GÜNCELLE *", false);
                if (!interactive) return;

                var answer = MessageBox.Show(_owner,
                    "Yeni Test Kontrol sürümü bulundu: " + manifest.Version + "\n\nŞimdi indirip kurmak ister misiniz?",
                    ProductName + " güncelleme", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (answer != DialogResult.Yes) return;

                await DownloadAndInstallAsync(manifest);
            }
            catch (Exception ex)
            {
                SetButton("GÜNCELLE", false);
                if (interactive) MessageBox.Show(_owner, "Güncelleme kontrolü/kurulumu başarısız oldu.\n\n" + ex.Message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                _busy = false;
            }
        }

        private static async Task<UpdateManifest> GetManifestAsync()
        {
            var url = ManifestUrl + "?_cache=" + DateTime.UtcNow.Ticks;
            using (var response = await Http.GetAsync(url))
            {
                response.EnsureSuccessStatusCode();
                using (var stream = await response.Content.ReadAsStreamAsync())
                {
                    var serializer = new DataContractJsonSerializer(typeof(UpdateManifest));
                    return serializer.ReadObject(stream) as UpdateManifest;
                }
            }
        }

        private static Version GetCurrentVersion()
        {
            Version version;
            return Version.TryParse(Application.ProductVersion, out version) ? version : new Version(0, 0, 0, 0);
        }

        private static int CompareVersions(string remote, Version current)
        {
            Version remoteVersion;
            if (!Version.TryParse(remote.TrimStart('v', 'V'), out remoteVersion)) return -1;
            return remoteVersion.CompareTo(current);
        }

        private static async Task DownloadAndInstallAsync(UpdateManifest manifest)
        {
            if (string.IsNullOrWhiteSpace(manifest.File) || manifest.Chunks == null || manifest.Chunks.Count == 0)
                throw new InvalidOperationException("Güncelleme manifestinde indirme parçaları bulunamadı.");

            SetButton("İNDİRİLİYOR...", true);
            var root = Path.Combine(Path.GetTempPath(), "TestKontrolUpdater");
            Directory.CreateDirectory(root);
            var zipPath = Path.Combine(root, "TestKontrolProg_update.zip");
            if (File.Exists(zipPath)) File.Delete(zipPath);

            await DownloadChunksAsync(manifest, zipPath);
            VerifyFile(zipPath, manifest.Size, manifest.Sha256);

            var extractDir = Path.Combine(root, "extract_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(extractDir);
            ZipFile.ExtractToDirectory(zipPath, extractDir);
            if (FindExecutable(extractDir) == null) throw new InvalidOperationException("Güncelleme ZIP'i içinde TestKontrolProg.exe bulunamadı.");

            var targetDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var targetExe = Path.Combine(targetDir, "TestKontrolProg.exe");
            StartInstaller(extractDir, targetDir, targetExe, Process.GetCurrentProcess().Id);
            _owner.Close();
        }

        private static async Task DownloadChunksAsync(UpdateManifest manifest, string target)
        {
            var buffers = new byte[manifest.Chunks.Count][];
            var gate = new object();
            long completed = 0;
            using (var semaphore = new SemaphoreSlim(4))
            {
                var tasks = new List<Task>();
                for (int i = 0; i < manifest.Chunks.Count; i++)
                {
                    int index = i;
                    tasks.Add(Task.Run(async () =>
                    {
                        await semaphore.WaitAsync();
                        try
                        {
                            var chunk = manifest.Chunks[index];
                            var url = BuildChunkUrl(chunk.File);
                            for (int attempt = 1; attempt <= 5; attempt++)
                            {
                                var data = await Http.GetByteArrayAsync(url + "?_cache=" + DateTime.UtcNow.Ticks + "_" + attempt);
                                if (data.LongLength == chunk.Size)
                                {
                                    buffers[index] = data;
                                    lock (gate)
                                    {
                                        completed += data.LongLength;
                                        SetButtonSafe("İNDİR " + completed / 1024 / 1024 + " MB", true);
                                    }
                                    return;
                                }
                            }
                            throw new InvalidOperationException("Güncelleme parçası eksik indirildi: " + chunk.File);
                        }
                        finally { semaphore.Release(); }
                    }));
                }
                await Task.WhenAll(tasks);
            }
            using (var output = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                foreach (var buffer in buffers) output.Write(buffer, 0, buffer.Length);
            }
        }

        private static string BuildChunkUrl(string file)
        {
            var baseUrl = ManifestUrl.Substring(0, ManifestUrl.LastIndexOf("/manifest.json", StringComparison.OrdinalIgnoreCase) + 1);
            return baseUrl + Uri.EscapeDataString(file);
        }

        private static void VerifyFile(string path, long expectedSize, string expectedSha256)
        {
            var actualSize = new FileInfo(path).Length;
            if (expectedSize > 0 && actualSize != expectedSize)
                throw new InvalidOperationException("Güncelleme boyutu doğrulanamadı: " + actualSize + "/" + expectedSize + " bayt.");
            using (var sha = SHA256.Create())
            using (var stream = File.OpenRead(path))
            {
                var actual = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
                var expected = (expectedSha256 ?? "").Replace("sha256:", "").ToLowerInvariant();
                if (!string.IsNullOrWhiteSpace(expected) && !string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Güncelleme SHA-256 doğrulaması başarısız.");
            }
        }

        private static string FindExecutable(string root)
        {
            var direct = Path.Combine(root, "TestKontrolProg.exe");
            if (File.Exists(direct)) return direct;
            foreach (var file in Directory.GetFiles(root, "TestKontrolProg.exe", SearchOption.AllDirectories)) return file;
            return null;
        }

        private static void StartInstaller(string sourceDir, string targetDir, string targetExe, int parentPid)
        {
            var script = Path.Combine(Path.GetTempPath(), "TestKontrolUpdater_" + Guid.NewGuid().ToString("N") + ".ps1");
            var escapedSource = sourceDir.Replace("'", "''");
            var escapedTarget = targetDir.Replace("'", "''");
            var escapedExe = targetExe.Replace("'", "''");
            var scriptText = @"
param([string]$Source, [string]$Target, [string]$Exe, [int]$ParentPid, [string]$Script)
$ErrorActionPreference = 'Stop'
try {
    for ($i=0; $i -lt 120; $i++) {
        if (-not (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue)) { break }
        Start-Sleep -Milliseconds 250
    }
    if (Get-Process -Id $ParentPid -ErrorAction SilentlyContinue) { throw 'Test Kontrol kapatılamadı.' }
    New-Item -ItemType Directory -Force -Path $Target | Out-Null
    Get-ChildItem -LiteralPath $Target -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Copy-Item -Path (Join-Path $Source '*') -Destination $Target -Recurse -Force
    if (-not (Test-Path -LiteralPath $Exe)) { throw 'Yeni TestKontrolProg.exe bulunamadı.' }
    Start-Process -FilePath $Exe -WorkingDirectory $Target
} catch {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(('Test Kontrol güncellemesi kurulamadı.`n`n' + $_.Exception.Message), 'Test Kontrol güncellemesi', 'OK', 'Error') | Out-Null
} finally {
    Remove-Item -LiteralPath $Source -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $Script -Force -ErrorAction SilentlyContinue
}
";
            File.WriteAllText(script, scriptText, new UTF8Encoding(false));
            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\" -Source \"" + escapedSource + "\" -Target \"" + escapedTarget + "\" -Exe \"" + escapedExe + "\" -ParentPid " + parentPid + " -Script \"" + script + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };
            Process.Start(psi);
        }

        private static void SetButton(string text, bool enabled)
        {
            if (_button == null || _button.IsDisposed) return;
            if (_button.InvokeRequired)
            {
                _button.BeginInvoke(new Action<string, bool>(SetButton), text, enabled);
                return;
            }
            _button.Text = text;
            _button.Enabled = enabled;
        }

        private static void SetButtonSafe(string text, bool enabled)
        {
            try { SetButton(text, enabled); } catch { }
        }

        [DataContract]
        private sealed class UpdateManifest
        {
            [DataMember(Name = "version")] public string Version { get; set; }
            [DataMember(Name = "build")] public string Build { get; set; }
            [DataMember(Name = "file")] public string File { get; set; }
            [DataMember(Name = "size")] public long Size { get; set; }
            [DataMember(Name = "sha256")] public string Sha256 { get; set; }
            [DataMember(Name = "chunks")] public List<UpdateChunk> Chunks { get; set; }
        }

        [DataContract]
        private sealed class UpdateChunk
        {
            [DataMember(Name = "file")] public string File { get; set; }
            [DataMember(Name = "size")] public long Size { get; set; }
        }
    }
}
