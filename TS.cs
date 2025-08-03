using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using Newtonsoft.Json;

public class ThemeChangerApplicationContext : ApplicationContext
{
    private NotifyIcon notifyIcon;
    private System.Windows.Forms.Timer timer;
    private AppConfig config;
    private bool isCurrentlyDay;
    private ToolStripMenuItem automaticSwitchingMenuItem;
    private string configPath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "appsettings.json");
    private string logPath = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "ThemeChanger.log");

    private void LogToFile(string message)
    {
        try
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            File.AppendAllText(logPath, $"[{timestamp}] {message}{Environment.NewLine}");
        }
        catch { /* Silent fail */ }
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);

    [DllImport("uxtheme.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int SetSystemVisualStyle(string pszFilename, string pszColor, string pszSize, int dwReserved);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out IntPtr lpdwResult);

    private const int SPI_SETDESKWALLPAPER = 20;
    private const int SPIF_UPDATEINIFILE = 0x01;
    private const int SPIF_SENDCHANGE = 0x02;
    private const int HWND_BROADCAST = 0xffff;
    private const int WM_SETTINGCHANGE = 0x001A;
    private const int SMTO_ABORTIFHUNG = 0x0002;

    public ThemeChangerApplicationContext()
    {
        try
        {
            LogToFile("[STARTUP] Application starting.");
            LoadConfiguration();
            InitializeTrayIcon();
            SetupTimer();
            UpdateTheme();
            LogToFile("[STARTUP] Application started successfully.");
        }
        catch (Exception ex)
        {
            string errorMsg = $"[FATAL ERROR] During startup: {ex.Message}";
            LogToFile(errorMsg);
            LogToFile($"[FATAL ERROR] Stack trace: {ex.StackTrace}");
            MessageBox.Show($"Theme Changer failed to start: {ex.Message}", "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void LoadConfiguration()
    {
        try
        {
            if (File.Exists(configPath))
            {
                string json = File.ReadAllText(configPath);
                config = JsonConvert.DeserializeObject<AppConfig>(json);
            }
            else
            {
                config = new AppConfig
                {
                    SunriseTime = new TimeSpan(6, 0, 0),
                    SunsetTime = new TimeSpan(18, 0, 0),
                    UseGeolocation = false,
                    CheckIntervalMinutes = 5,
                    DayThemePath = "",
                    NightThemePath = "",
                    DayWallpaperPath = "",
                    NightWallpaperPath = ""
                };
                SaveConfiguration();
            }
        }
        catch (Exception ex)
        {
            LogToFile($"[ERROR] Failed to load configuration: {ex.Message}");
            MessageBox.Show($"Error loading configuration: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            config = new AppConfig();
        }
    }

    private void SaveConfiguration()
    {
        try
        {
            string json = JsonConvert.SerializeObject(config, Formatting.Indented);
            File.WriteAllText(configPath, json);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error saving configuration: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void InitializeTrayIcon()
    {
        ContextMenuStrip contextMenu = new ContextMenuStrip();
        automaticSwitchingMenuItem = new ToolStripMenuItem("Automatic Switching", null, ToggleAutomaticSwitching_Click)
        {
            CheckOnClick = true,
            Checked = true
        };

        contextMenu.Items.Add("Toggle Theme", null, ToggleTheme_Click);
        contextMenu.Items.Add(automaticSwitchingMenuItem);
        contextMenu.Items.Add("Settings", null, Settings_Click);
        contextMenu.Items.Add(new ToolStripSeparator());
        contextMenu.Items.Add("Exit", null, Exit_Click);

        notifyIcon = new NotifyIcon()
        {
            Icon = SystemIcons.Application,
            ContextMenuStrip = contextMenu,
            Text = "Theme Changer",
            Visible = true
        };
    }

    private void SetupTimer()
    {
        timer = new System.Windows.Forms.Timer();
        timer.Tick += OnTimerElapsed;
        UpdateTimerInterval();
        timer.Start();
    }

    private void UpdateTimerInterval()
    {
        int intervalMs = config.CheckIntervalMinutes * 60 * 1000;
        timer.Interval = intervalMs > 0 ? intervalMs : 60000;
    }

    private void OnTimerElapsed(object sender, EventArgs e)
    {
        LogToFile("[TIMER] Timer elapsed. Updating theme.");
        UpdateTheme();
    }

    private void UpdateTheme()
    {
        TimeSpan currentTime = DateTime.Now.TimeOfDay;
        isCurrentlyDay = currentTime >= config.SunriseTime && currentTime < config.SunsetTime;

        if (isCurrentlyDay)
        {
            ApplyTheme(config.DayThemePath, config.DayWallpaperPath, true);
        }
        else
        {
            ApplyTheme(config.NightThemePath, config.NightWallpaperPath, false);
        }
    }

    private void ApplyTheme(string themePath, string wallpaperOverridePath, bool isDay)
    {
        LogToFile($"[THEME] Applying theme. IsDay: {isDay}, ThemePath: '{themePath}', WallpaperOverride: '{wallpaperOverridePath}'");

        // 1. Apply the visual style and wallpaper from the .theme file.
        if (!string.IsNullOrEmpty(themePath) && File.Exists(themePath))
        {
            ApplyThemeFile(themePath);
        }

        // 2. Apply a specific wallpaper override, if provided.
        if (!string.IsNullOrEmpty(wallpaperOverridePath) && File.Exists(wallpaperOverridePath))
        {
            LogToFile($"[WALLPAPER] Applying override: {wallpaperOverridePath}");
            SetWallpaper(wallpaperOverridePath, "2", "0"); // Default to Stretch
        }

        // 3. Set light/dark mode.
        SetWindowsTheme(isDay == false);

        // 4. Give the system a moment to process the registry changes before broadcasting.
        System.Threading.Thread.Sleep(500);

        // 5. Perform a robust refresh at the end.
        RefreshSystemParameters();
        notifyIcon.Text = $"Theme Changer - {(isDay ? "Day" : "Night")} Theme";
    }

    private void ApplyThemeFile(string themePath)
    {
        try
        {
            LogToFile($"[THEME_FILE] Parsing and applying: {themePath}");
            var themeSettings = ParseThemeFile(themePath);

            // Apply visual style from .msstyles
            if (themeSettings.TryGetValue("VisualStyle", out string stylePath) && File.Exists(stylePath))
            {
                LogToFile($"[THEME_FILE] Applying visual style: {stylePath}");
                SetSystemVisualStyle(stylePath, null, null, 0);
            }

            // Apply wallpaper from theme
            if (themeSettings.TryGetValue("Wallpaper", out string wallpaperPath) && File.Exists(wallpaperPath))
            {
                string style = themeSettings.GetValueOrDefault("WallpaperStyle", "2");
                string tile = themeSettings.GetValueOrDefault("TileWallpaper", "0");
                LogToFile($"[THEME_FILE] Applying wallpaper: {wallpaperPath}");
                SetWallpaper(wallpaperPath, style, tile);
            }
        }
        catch (Exception ex)
        {
            LogToFile($"[ERROR] Failed during ApplyThemeFile for '{themePath}': {ex.Message}");
        }
    }

    private Dictionary<string, string> ParseThemeFile(string themePath)
    {
        var settings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var lines = File.ReadAllLines(themePath);
            string currentSection = "";
            foreach (var line in lines)
            {
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    currentSection = line.Substring(1, line.Length - 2);
                    continue;
                }
                if (string.IsNullOrWhiteSpace(line) || !line.Contains("=")) continue;

                var parts = line.Split(new[] { '=' }, 2);
                string key = parts[0].Trim();
                string value = Environment.ExpandEnvironmentVariables(parts[1].Trim());

                if (currentSection.Equals("VisualStyles", StringComparison.OrdinalIgnoreCase) && key.Equals("Path", StringComparison.OrdinalIgnoreCase))
                {
                    settings["VisualStyle"] = value;
                }
                else if (currentSection.Equals("Control Panel\\Desktop", StringComparison.OrdinalIgnoreCase))
                {
                    if (key.Equals("Wallpaper", StringComparison.OrdinalIgnoreCase)) settings["Wallpaper"] = value;
                    else if (key.Equals("WallpaperStyle", StringComparison.OrdinalIgnoreCase)) settings["WallpaperStyle"] = value;
                    else if (key.Equals("TileWallpaper", StringComparison.OrdinalIgnoreCase)) settings["TileWallpaper"] = value;
                }
            }
        }
        catch (Exception ex)
        {
            LogToFile($"[ERROR] Failed to parse theme file '{themePath}': {ex.Message}");
        }
        return settings;
    }

    private void SetWindowsTheme(bool useDarkTheme)
    {
        try
        {
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", useDarkTheme ? 0 : 1, RegistryValueKind.DWord);
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", "SystemUsesLightTheme", useDarkTheme ? 0 : 1, RegistryValueKind.DWord);
        }
        catch (Exception ex)
        {
            LogToFile($"[ERROR] Failed to set Windows light/dark mode: {ex.Message}");
        }
    }

    private void SetWallpaper(string path, string style, string tile)
    {
        try
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop", true);
            key?.SetValue("WallpaperStyle", style);
            key?.SetValue("TileWallpaper", tile);
            key?.Close();
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE);
        }
        catch (Exception ex)
        {
            LogToFile($"[ERROR] Failed to set wallpaper '{path}': {ex.Message}");
        }
    }

    private void RefreshSystemParameters()
    {
        LogToFile("[REFRESH] Broadcasting system setting changes.");
        SendMessageTimeout((IntPtr)HWND_BROADCAST, WM_SETTINGCHANGE, IntPtr.Zero, "ImmersiveColorSet", SMTO_ABORTIFHUNG, 5000, out _);
        SendMessageTimeout((IntPtr)HWND_BROADCAST, WM_SETTINGCHANGE, IntPtr.Zero, "Themes", SMTO_ABORTIFHUNG, 5000, out _);
        SendMessageTimeout((IntPtr)HWND_BROADCAST, WM_SETTINGCHANGE, IntPtr.Zero, null, SMTO_ABORTIFHUNG, 5000, out _);
    }

    private void ToggleTheme_Click(object sender, EventArgs e)
    {
        LogToFile("[ACTION] Manual toggle initiated. The theme will revert to the schedule on the next timer tick.");
        isCurrentlyDay = !isCurrentlyDay;
        if (isCurrentlyDay)
        {
            ApplyTheme(config.DayThemePath, config.DayWallpaperPath, true);
        }
        else
        {
            ApplyTheme(config.NightThemePath, config.NightWallpaperPath, false);
        }
    }

    private void ToggleAutomaticSwitching_Click(object sender, EventArgs e)
    {
        if (automaticSwitchingMenuItem.Checked)
        {
            LogToFile("[ACTION] Automatic switching enabled.");
            UpdateTheme();
            timer.Start();
        }
        else
        {
            LogToFile("[ACTION] Automatic switching disabled.");
            timer.Stop();
        }
    }

    private void Settings_Click(object sender, EventArgs e)
    {
        bool wasTimerActive = timer.Enabled;
        timer.Stop();

        using (var settingsForm = new SettingsForm(config))
        {
            if (settingsForm.ShowDialog() == DialogResult.OK)
            {
                config = settingsForm.GetUpdatedConfig();
                SaveConfiguration();
                UpdateTimerInterval();
                LogToFile("[ACTION] Settings saved and updated.");

                if (automaticSwitchingMenuItem.Checked)
                {
                    UpdateTheme();
                }
            }
        }

        if (wasTimerActive && automaticSwitchingMenuItem.Checked)
        {
            timer.Start();
        }
    }

    private void Exit_Click(object sender, EventArgs e)
    {
        ExitThread();
    }

    protected override void ExitThreadCore()
    {
        if (notifyIcon != null)
        {
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
        }
        if (timer != null)
        {
            timer.Stop();
            timer.Dispose();
        }
        base.ExitThreadCore();
    }
}

public class AppConfig
{
    public TimeSpan SunriseTime { get; set; }
    public TimeSpan SunsetTime { get; set; }
    public bool UseGeolocation { get; set; }
    public int CheckIntervalMinutes { get; set; }
    public string DayWallpaperPath { get; set; }
    public string NightWallpaperPath { get; set; }
    public string DayThemePath { get; set; }
    public string NightThemePath { get; set; }
}

public class Location
{
    public double Latitude { get; set; }
    public double Longitude { get; set; }
}

public class ThemeInfo
{
    public string Name { get; set; }
    public string Path { get; set; }
}

public class SettingsForm : Form
{
    private AppConfig config;
    private TextBox sunriseTextBox;
    private TextBox sunsetTextBox;
    private CheckBox geolocationCheckBox;
    private NumericUpDown intervalNumericUpDown;
    private TextBox dayWallpaperTextBox;
    private TextBox nightWallpaperTextBox;
    private ComboBox dayThemeComboBox;
    private ComboBox nightThemeComboBox;

    public SettingsForm(AppConfig config)
    {
        this.config = JsonConvert.DeserializeObject<AppConfig>(JsonConvert.SerializeObject(config));
        InitializeComponent();
        LoadSettings();
    }

    private void InitializeComponent()
    {
        this.Text = "Theme Changer Settings";
        this.Width = 400;
        this.Height = 480;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.Padding = new Padding(10);

        int yPos = 20;
        Label sunriseLabel = new Label() { Text = "Sunrise Time (HH:mm:ss):", Location = new Point(20, yPos), Width = 150 };
        sunriseTextBox = new TextBox() { Location = new Point(180, yPos), Width = 150 };
        yPos += 30;
        Label sunsetLabel = new Label() { Text = "Sunset Time (HH:mm:ss):", Location = new Point(20, yPos), Width = 150 };
        sunsetTextBox = new TextBox() { Location = new Point(180, yPos), Width = 150 };
        yPos += 40;
        geolocationCheckBox = new CheckBox() { Text = "Use Geolocation (Not Implemented)", Location = new Point(20, yPos), Width = 300, Enabled = false };
        yPos += 40;
        Label intervalLabel = new Label() { Text = "Check Interval (min):", Location = new Point(20, yPos), Width = 150 };
        intervalNumericUpDown = new NumericUpDown() { Location = new Point(180, yPos), Width = 150, Minimum = 1, Maximum = 1440 };
        yPos += 40;
        Label dayThemeLabel = new Label() { Text = "Day Theme:", Location = new Point(20, yPos), Width = 120 };
        dayThemeComboBox = new ComboBox() { Location = new Point(150, yPos), Width = 215, DropDownStyle = ComboBoxStyle.DropDownList };
        yPos += 40;
        Label dayWallpaperLabel = new Label() { Text = "Day Wallpaper (Override):", Location = new Point(20, yPos), Width = 120 };
        dayWallpaperTextBox = new TextBox() { Location = new Point(150, yPos), Width = 180 };
        Button dayWallpaperButton = new Button() { Text = "...", Location = new Point(335, yPos), Width = 30 };
        dayWallpaperButton.Click += (s, e) => SelectFile(dayWallpaperTextBox, "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif", "Select Day Wallpaper");
        yPos += 40;
        Label nightThemeLabel = new Label() { Text = "Night Theme:", Location = new Point(20, yPos), Width = 120 };
        nightThemeComboBox = new ComboBox() { Location = new Point(150, yPos), Width = 215, DropDownStyle = ComboBoxStyle.DropDownList };
        yPos += 40;
        Label nightWallpaperLabel = new Label() { Text = "Night Wallpaper (Override):", Location = new Point(20, yPos), Width = 120 };
        nightWallpaperTextBox = new TextBox() { Location = new Point(150, yPos), Width = 180 };
        Button nightWallpaperButton = new Button() { Text = "...", Location = new Point(335, yPos), Width = 30 };
        nightWallpaperButton.Click += (s, e) => SelectFile(nightWallpaperTextBox, "Image Files|*.jpg;*.jpeg;*.png;*.bmp;*.gif", "Select Night Wallpaper");
        yPos += 50;
        Button saveButton = new Button() { Text = "Save", Location = new Point(100, yPos), Width = 85, DialogResult = DialogResult.OK };
        Button cancelButton = new Button() { Text = "Cancel", Location = new Point(200, yPos), Width = 85, DialogResult = DialogResult.Cancel };

        saveButton.Click += SaveButton_Click;
        this.AcceptButton = saveButton;
        this.CancelButton = cancelButton;

        this.Controls.AddRange(new Control[] {
            sunriseLabel, sunriseTextBox, sunsetLabel, sunsetTextBox, geolocationCheckBox, intervalLabel, intervalNumericUpDown,
            dayThemeLabel, dayThemeComboBox, dayWallpaperLabel, dayWallpaperTextBox, dayWallpaperButton,
            nightThemeLabel, nightThemeComboBox, nightWallpaperLabel, nightWallpaperTextBox, nightWallpaperButton,
            saveButton, cancelButton
        });
    }

    private void LoadSettings()
    {
        var themes = GetInstalledThemes();
        dayThemeComboBox.DataSource = new BindingSource(themes, null);
        dayThemeComboBox.DisplayMember = "Name";
        dayThemeComboBox.ValueMember = "Path";

        nightThemeComboBox.DataSource = new BindingSource(themes, null);
        nightThemeComboBox.DisplayMember = "Name";
        nightThemeComboBox.ValueMember = "Path";

        sunriseTextBox.Text = config.SunriseTime.ToString(@"hh\:mm\:ss");
        sunsetTextBox.Text = config.SunsetTime.ToString(@"hh\:mm\:ss");
        intervalNumericUpDown.Value = Math.Max(1, config.CheckIntervalMinutes);
        dayWallpaperTextBox.Text = config.DayWallpaperPath ?? "";
        nightWallpaperTextBox.Text = config.NightWallpaperPath ?? "";

        dayThemeComboBox.SelectedValue = themes.Any(t => t.Path == config.DayThemePath) ? config.DayThemePath : "";
        nightThemeComboBox.SelectedValue = themes.Any(t => t.Path == config.NightThemePath) ? config.NightThemePath : "";
    }

    private void SaveButton_Click(object sender, EventArgs e)
    {
        if (!ValidateInputs()) {
            this.DialogResult = DialogResult.None; // Keep form open
            return;
        }

        config.SunriseTime = TimeSpan.Parse(sunriseTextBox.Text);
        config.SunsetTime = TimeSpan.Parse(sunsetTextBox.Text);
        config.CheckIntervalMinutes = (int)intervalNumericUpDown.Value;
        config.DayWallpaperPath = dayWallpaperTextBox.Text;
        config.NightWallpaperPath = nightWallpaperTextBox.Text;
        config.DayThemePath = (string)dayThemeComboBox.SelectedValue;
        config.NightThemePath = (string)nightThemeComboBox.SelectedValue;
    }

    private bool ValidateInputs()
    {
        if (!TimeSpan.TryParse(sunriseTextBox.Text, out _))
        {
            MessageBox.Show("Invalid format for Sunrise Time. Please use HH:mm:ss.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        if (!TimeSpan.TryParse(sunsetTextBox.Text, out _))
        {
            MessageBox.Show("Invalid format for Sunset Time. Please use HH:mm:ss.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        return true;
    }

    private void SelectFile(TextBox textBox, string filter, string title)
    {
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = filter;
            openFileDialog.Title = title;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = openFileDialog.FileName;
            }
        }
    }

    private List<ThemeInfo> GetInstalledThemes()
    {
        var themes = new List<ThemeInfo>();
        var themePaths = new List<string>();
        string systemThemePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), @"Resources\Themes");
        if (Directory.Exists(systemThemePath)) themePaths.AddRange(Directory.GetFiles(systemThemePath, "*.theme"));
        string userThemePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Microsoft\Windows\Themes");
        if (Directory.Exists(userThemePath)) themePaths.AddRange(Directory.GetFiles(userThemePath, "*.theme"));

        foreach (var path in themePaths.Distinct())
        {
            try
            {
                string name = GetThemeDisplayName(path);
                if (!string.IsNullOrEmpty(name)) themes.Add(new ThemeInfo { Name = name, Path = path });
            }
            catch (Exception ex) { Debug.WriteLine($"Could not parse theme '{path}': {ex.Message}"); }
        }
        return themes.OrderBy(t => t.Name).ToList();
    }

    private string GetThemeDisplayName(string themePath)
    {
        try
        {
            var lines = File.ReadAllLines(themePath);
            string displayName = "";
            bool inThemeSection = false;
            foreach (var line in lines)
            {
                if (line.Trim().Equals("[Theme]", StringComparison.OrdinalIgnoreCase)) inThemeSection = true;
                else if (inThemeSection && line.Trim().StartsWith("DisplayName", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = line.Split('=');
                    if (parts.Length > 1) { displayName = parts[1].Trim(); break; }
                }
                else if (line.Trim().StartsWith("[")) inThemeSection = false;
            }
            if (displayName.StartsWith("@")) return GetStringResource(displayName);
            return string.IsNullOrEmpty(displayName) ? Path.GetFileNameWithoutExtension(themePath) : displayName;
        }
        catch { return Path.GetFileNameWithoutExtension(themePath); }
    }

    private string GetStringResource(string resource)
    {
        try
        {
            string[] parts = resource.Split(',');
            if (parts.Length != 2) return "Unknown Theme";
            string libraryPath = Environment.ExpandEnvironmentVariables(parts[0].TrimStart('@').Trim());
            if (!File.Exists(libraryPath)) return "Unknown Theme";
            if (!int.TryParse(parts[1].Trim(), out int resourceId)) return "Unknown Theme";
            IntPtr hModule = LoadLibraryEx(libraryPath, IntPtr.Zero, 0x00000002);
            if (hModule == IntPtr.Zero) return "Unknown Theme";
            System.Text.StringBuilder sb = new System.Text.StringBuilder(255);
            if (LoadString(hModule, (uint)Math.Abs(resourceId), sb, sb.Capacity) > 0) return sb.ToString();
            return "Unknown Theme";
        }
        catch { return "Unknown Theme"; }
    }

    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int LoadString(IntPtr hInstance, uint uID, System.Text.StringBuilder lpBuffer, int nBufferMax);

    public AppConfig GetUpdatedConfig() { return config; }
}

public class Program
{
    [STAThread]
    public static void Main()
    {
        AddToStartup();
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new ThemeChangerApplicationContext());
    }

    private static void AddToStartup()
    {
        try
        {
            string exePath = Application.ExecutablePath;
            string appName = Path.GetFileNameWithoutExtension(exePath);
            RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            registryKey?.SetValue(appName, "\"" + exePath + "\"");
        }
        catch (Exception ex) { Debug.WriteLine($"Failed to add to startup: {ex.Message}"); }
    }
}