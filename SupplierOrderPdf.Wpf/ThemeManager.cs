using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using SupplierOrderPdf.Core;
using MaterialDesignThemes.Wpf;

namespace SupplierOrderPdf.Wpf
{
    /// <summary>
    /// Управляет применением Material Design 3 тем в WPF приложении.
    /// 
    /// Поддерживает три режима:
    /// - Light: Светлая тема MD3
    /// - Dark: Тёмная тема MD3
    /// - Auto: Автоматическое переключение в зависимости от системной темы Windows
    /// 
    /// Работает с новой структурой ResourceDictionary:
    /// - Themes/Palettes/Palette.Light.xaml - светлая палитра
    /// - Themes/Palettes/Palette.Dark.xaml - тёмная палитра
    /// - Themes/Theme.xaml - главная точка входа (токены + палитра)
    /// </summary>
    public static class ThemeManager
    {
        /// <summary>
        /// Текущий режим темы (Auto/Light/Dark), выбранный пользователем.
        /// </summary>
        public static UiThemeMode CurrentTheme { get; private set; } = UiThemeMode.Auto;

        /// <summary>
        /// Событие, вызываемое при изменении темы.
        /// </summary>
        public static event EventHandler<UiThemeMode>? ThemeChanged;

        private static readonly PaletteHelper _paletteHelper = new PaletteHelper();
        private static bool _isInitialized = false;

        /// <summary>
        /// URI палитры светлой темы
        /// </summary>
        private const string LightPaletteUri = "Themes/Palettes/Palette.Light.xaml";
        
        /// <summary>
        /// URI палитры тёмной темы
        /// </summary>
        private const string DarkPaletteUri = "Themes/Palettes/Palette.Dark.xaml";

        /// <summary>
        /// Статический конструктор ThemeManager.
        /// Подписывается на системные события изменения пользовательских настроек Windows.
        /// </summary>
        static ThemeManager()
        {
            SubscribeToSystemThemeChanges();
        }

        /// <summary>
        /// Инициализирует ThemeManager с указанной начальной темой.
        /// Должен вызываться в App.xaml.cs в методе OnStartup.
        /// </summary>
        /// <param name="initialTheme">Начальная тема (по умолчанию Auto)</param>
        public static void Initialize(UiThemeMode initialTheme = UiThemeMode.Auto)
        {
            if (_isInitialized)
            {
                System.Diagnostics.Debug.WriteLine("[ThemeManager] Already initialized, applying theme only");
                ApplyTheme(initialTheme);
                return;
            }

            _isInitialized = true;
            System.Diagnostics.Debug.WriteLine($"[ThemeManager] Initialize with theme={initialTheme}");
            
            ApplyTheme(initialTheme);
        }

        /// <summary>
        /// Применяет указанную тему к приложению.
        /// </summary>
        /// <param name="theme">Режим темы для применения</param>
        public static void ApplyTheme(UiThemeMode theme)
        {
            // Обрабатываем устаревший Material как Light
            #pragma warning disable CS0618
            if (theme == UiThemeMode.Material)
            {
                theme = UiThemeMode.Light;
            }
            #pragma warning restore CS0618

            CurrentTheme = theme;

            var app = Application.Current;
            if (app == null)
            {
                System.Diagnostics.Debug.WriteLine("[ThemeManager] Application.Current is null");
                return;
            }

            // Определяем эффективную тему (Light или Dark)
            bool isDark = theme == UiThemeMode.Dark || 
                         (theme == UiThemeMode.Auto && IsSystemDarkTheme());

            System.Diagnostics.Debug.WriteLine($"[ThemeManager] ApplyTheme mode={theme}, isDark={isDark}");

            // 1. Переключаем палитру MD3
            SwitchPalette(isDark);

            // 2. Обновляем тему MaterialDesignThemes
            UpdateMaterialDesignTheme(isDark);

            // 3. Вызываем событие изменения темы
            ThemeChanged?.Invoke(null, theme);
        }

        /// <summary>
        /// Переключает палитру MD3 в MergedDictionaries.
        /// </summary>
        /// <param name="isDark">True для тёмной темы, False для светлой</param>
        private static void SwitchPalette(bool isDark)
        {
            var app = Application.Current;
            if (app == null) return;

            var dictionaries = app.Resources.MergedDictionaries;
            
            // Ищем текущую палитру (Palette.Light.xaml или Palette.Dark.xaml)
            ResourceDictionary? currentPalette = null;
            int paletteIndex = -1;

            for (int i = 0; i < dictionaries.Count; i++)
            {
                var dict = dictionaries[i];
                if (dict.Source != null)
                {
                    var source = dict.Source.OriginalString;
                    if (source.Contains("Palette.Light.xaml", StringComparison.OrdinalIgnoreCase) ||
                        source.Contains("Palette.Dark.xaml", StringComparison.OrdinalIgnoreCase))
                    {
                        currentPalette = dict;
                        paletteIndex = i;
                        break;
                    }
                }
                
                // Также проверяем вложенные словари (Theme.xaml содержит палитру)
                if (dict.MergedDictionaries.Count > 0)
                {
                    for (int j = 0; j < dict.MergedDictionaries.Count; j++)
                    {
                        var nested = dict.MergedDictionaries[j];
                        if (nested.Source != null)
                        {
                            var nestedSource = nested.Source.OriginalString;
                            if (nestedSource.Contains("Palette.Light.xaml", StringComparison.OrdinalIgnoreCase) ||
                                nestedSource.Contains("Palette.Dark.xaml", StringComparison.OrdinalIgnoreCase))
                            {
                                // Заменяем палитру внутри Theme.xaml
                                var newPaletteUri = isDark ? DarkPaletteUri : LightPaletteUri;
                                var newPalette = new ResourceDictionary
                                {
                                    Source = new Uri(newPaletteUri, UriKind.Relative)
                                };
                                
                                dict.MergedDictionaries[j] = newPalette;
                                System.Diagnostics.Debug.WriteLine($"[ThemeManager] Replaced nested palette at index {j} with {newPaletteUri}");
                                return;
                            }
                        }
                    }
                }
            }

            // Если палитра найдена на верхнем уровне, заменяем её
            if (currentPalette != null && paletteIndex >= 0)
            {
                var newPaletteUri = isDark ? DarkPaletteUri : LightPaletteUri;
                var newPalette = new ResourceDictionary
                {
                    Source = new Uri(newPaletteUri, UriKind.Relative)
                };

                dictionaries[paletteIndex] = newPalette;
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] Replaced palette at index {paletteIndex} with {newPaletteUri}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("[ThemeManager] No palette found to replace");
            }
        }

        /// <summary>
        /// Обновляет тему MaterialDesignThemes (BundledTheme).
        /// </summary>
        /// <param name="isDark">True для тёмной темы, False для светлой</param>
        private static void UpdateMaterialDesignTheme(bool isDark)
        {
            try
            {
                var theme = _paletteHelper.GetTheme();
                
                // Устанавливаем базовую тему
                theme.SetBaseTheme(isDark ? BaseTheme.Dark : BaseTheme.Light);

                // Устанавливаем цвета MD3 (фиолетовая палитра)
                if (isDark)
                {
                    // Dark theme colors
                    theme.SetPrimaryColor((Color)ColorConverter.ConvertFromString("#D0BCFF")); // Primary 80
                    theme.SetSecondaryColor((Color)ColorConverter.ConvertFromString("#CCC2DC")); // Secondary 80
                }
                else
                {
                    // Light theme colors
                    theme.SetPrimaryColor((Color)ColorConverter.ConvertFromString("#6750A4")); // Primary 40
                    theme.SetSecondaryColor((Color)ColorConverter.ConvertFromString("#625B71")); // Secondary 40
                }

                _paletteHelper.SetTheme(theme);
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] MaterialDesign theme updated, isDark={isDark}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] Error updating MaterialDesign theme: {ex.Message}");
            }
        }

        /// <summary>
        /// Определяет, использует ли система Windows тёмную тему.
        /// Читает значение из реестра Windows.
        /// </summary>
        /// <returns>True если системная тема тёмная, False если светлая</returns>
        public static bool IsSystemDarkTheme()
        {
            try
            {
                const string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
                const string appsValueName = "AppsUseLightTheme";

                using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(keyPath);
                if (key != null)
                {
                    if (key.GetValue(appsValueName) is int appsLight)
                    {
                        // 0 = dark, 1 = light
                        var isDark = appsLight == 0;
                        System.Diagnostics.Debug.WriteLine($"[ThemeManager] IsSystemDarkTheme: AppsUseLightTheme={appsLight}, isDark={isDark}");
                        return isDark;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] Error reading system theme: {ex.Message}");
            }

            // По умолчанию светлая тема
            return false;
        }

        /// <summary>
        /// Подписывается на изменения системной темы Windows.
        /// </summary>
        private static void SubscribeToSystemThemeChanges()
        {
            try
            {
                Microsoft.Win32.SystemEvents.UserPreferenceChanged += OnUserPreferenceChanged;
                System.Diagnostics.Debug.WriteLine("[ThemeManager] Subscribed to system theme changes");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] Failed to subscribe to system events: {ex.Message}");
            }
        }

        /// <summary>
        /// Обработчик изменения системных настроек Windows.
        /// В режиме Auto автоматически переключает тему приложения.
        /// </summary>
        private static void OnUserPreferenceChanged(object? sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
        {
            // Реагируем только в режиме Auto
            if (CurrentTheme != UiThemeMode.Auto)
                return;

            // Проверяем категории, связанные с темой
            if (e.Category == Microsoft.Win32.UserPreferenceCategory.General ||
                e.Category == Microsoft.Win32.UserPreferenceCategory.Color ||
                e.Category == Microsoft.Win32.UserPreferenceCategory.VisualStyle)
            {
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] System theme changed, category={e.Category}");
                
                // Применяем тему в UI потоке
                Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                {
                    ApplyTheme(UiThemeMode.Auto);
                }));
            }
        }

        /// <summary>
        /// Освобождает ресурсы ThemeManager.
        /// Отписывается от системных событий.
        /// Должен вызываться в App.OnExit.
        /// </summary>
        public static void Shutdown()
        {
            try
            {
                Microsoft.Win32.SystemEvents.UserPreferenceChanged -= OnUserPreferenceChanged;
                System.Diagnostics.Debug.WriteLine("[ThemeManager] Unsubscribed from system events");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThemeManager] Error during shutdown: {ex.Message}");
            }
        }
    }
}