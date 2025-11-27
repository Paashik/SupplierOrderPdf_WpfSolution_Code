using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Unicode;

namespace SupplierOrderPdf.Core;

/// <summary>
/// Режим выбора UI‑темы в WPF‑клиенте.
/// Auto — использовать системную тему Windows; Light/Dark — принудительный выбор.
/// Material устарел и будет обрабатываться как Light для обратной совместимости.
/// </summary>
public enum UiThemeMode
{
    /// <summary>
    /// Автоматический режим - следует системной теме Windows
    /// </summary>
    Auto = 0,
    
    /// <summary>
    /// Светлая тема Material Design 3
    /// </summary>
    Light = 1,
    
    /// <summary>
    /// Тёмная тема Material Design 3
    /// </summary>
    Dark = 2,
    
    /// <summary>
    /// [Устарело] Обрабатывается как Light для обратной совместимости
    /// </summary>
    [Obsolete("Use Light instead. Material mode is deprecated.")]
    Material = 3
}

/// <summary>
/// Класс настроек приложения SupplierOrderPdf.
/// Содержит все конфигурационные параметры приложения, включая:
/// - Пути к базам данных и выходным директориям
/// - Настройки пользовательского интерфейса и тем
/// - Конфигурацию SMTP/IMAP для отправки email
/// - Геометрию окон и предпочтения пользователя
/// 
/// Настройки сохраняются в SQLite базе данных через AppDatabase.
/// Поддерживает значения по умолчанию для всех параметров.
/// </summary>
public sealed class AppSettings
{
    /// <summary>
    /// Статический класс с константами значений по умолчанию для всех настроек приложения.
    /// Эти значения используются при первом запуске приложения или при отсутствии сохраненных настроек.
    /// </summary>
    public static class Defaults
    {
        /// <summary>
        /// Путь к базе данных Access по умолчанию.
        /// Указывает на сетевой ресурс с базой данных компонентов.
        /// </summary>
        public const string DbPath = @"\\NETGEAR\\sklad\\Компонент 2014\\Comp_geterodin.mdb";
        
        /// <summary>
        /// Пароль для подключения к базе данных Access.
        /// </summary>
        public const string DbPassword = "Component-2014";
        
        /// <summary>
        /// Директория для сохранения сгенерированных PDF файлов заявок.
        /// </summary>
        public const string OutputDir = @"\\Netgear\\projects\\0. ЗАКУПКА\\Заявки на закупку";
        
        /// <summary>
        /// Название компании/организации, которое отображается в генерируемых документах.
        /// </summary>
        public const string CompanyName = "Отдел 6.1 НТЦ-6, Отв. Ванчарин П.М.";
        
        /// <summary>
        /// Массив адресов доставки по умолчанию.
        /// Используется для выбора места доставки при создании заявки.
        /// </summary>
        public static readonly string[] DeliveryBook =
        {
            "141190, Московская область, гор. Фрязино, Заводской проезд, д. 3 ...",
            "105118, г. Москва, проспект Будённого, д. 32А"
        };
        
        /// <summary>
        /// Тема оформления интерфейса по умолчанию.
        /// </summary>
        public const UiThemeMode Theme = UiThemeMode.Auto;
    }

    // ==================== Основные параметры ====================
    
    /// <summary>
    /// Путь к базе данных Access с информацией о компонентах.
    /// Может содержать переменные среды (например, %APPDATA%), которые будут раскрыты при использовании.
    /// </summary>
    public string DbPath { get; set; } = Defaults.DbPath;
    
    /// <summary>
    /// Пароль для подключения к базе данных Access.
    /// Хранится в открытом виде в настройках (требует улучшения безопасности).
    /// </summary>
    public string DbPassword { get; set; } = Defaults.DbPassword;
    
    /// <summary>
    /// Директория для сохранения созданных PDF файлов заявок.
    /// Автоматически создается при первом использовании.
    /// </summary>
    public string OutputDir { get; set; } = Defaults.OutputDir;
    
    /// <summary>
    /// Название компании или организации, от имени которой создаются заявки.
    /// Отображается в заголовке генерируемых документов.
    /// </summary>
    public string CompanyName { get; set; } = Defaults.CompanyName;
    
    /// <summary>
    /// Email администратора для отправки уведомлений о работе системы.
    /// Используется как fallback для NotifyOnCreateTo и NotifyOnSendTo.
    /// </summary>
    public string AdminNotifyEmail { get; set; } = string.Empty;

    /// <summary>
    /// Список доступных адресов доставки.
    /// Хранится в JSON формате в базе данных.
    /// Пользователь может выбирать адрес из этого списка при создании заявки.
    /// </summary>
    public string[] DeliveryAddresses { get; set; } = (string[])Defaults.DeliveryBook.Clone();
    
    /// <summary>
    /// Идентификатор последнего выбранного контактного лица клиента.
    /// Используется для быстрого доступа к часто используемым контактам.
    /// </summary>
    public int? LastCustomerPersonId { get; set; }
    
    /// <summary>
    /// Словарь, хранящий последние выбранные контактные лица поставщиков.
    /// Ключ: ID поставщика, Значение: ID контактного лица.
    /// Ускоряет выбор контактов при работе с постоянными поставщиками.
    /// </summary>
    public Dictionary<int, int> LastSupplierContact { get; set; } = new();

    /// ==================== Поведение приложения ====================
    
    /// <summary>
    /// Флаг автоматического прикрепления счета к заявке.
    /// Если true, счет будет включен в PDF документ по умолчанию.
    /// </summary>
    public bool AttachInvoiceByDefault { get; set; } = true;
    
    /// <summary>
    /// Флаг автоматического открытия PDF файла после экспорта.
    /// Если true, созданный PDF будет открыт в программе просмотра по умолчанию.
    /// </summary>
    public bool OpenPdfAfterExport { get; set; } = false;
    
    /// <summary>
    /// Флаг автоматической отправки заявки по email после создания.
    /// Если true, заявка будет отправлена автоматически без дополнительного подтверждения.
    /// </summary>
    public bool SendByEmail { get; set; } = false;
    
    /// <summary>
    /// Флаг автоматической установки статуса "Заказано" при отправке.
    /// Если true, позиции заявки будут помечены как заказанные после отправки.
    /// </summary>
    public bool AutoSetOrderedOnSend { get; set; } = false;

    // ==================== Геометрия главного окна ====================
    
    /// <summary>
    /// X-координата левого верхнего угла главного окна.
    /// Значение -1 означает, что позиция не сохранена и будет использоваться позиция по умолчанию.
    /// </summary>
    public int WindowX { get; set; } = -1;
    
    /// <summary>
    /// Y-координата левого верхнего угла главного окна.
    /// Значение -1 означает, что позиция не сохранена и будет использоваться позиция по умолчанию.
    /// </summary>
    public int WindowY { get; set; } = -1;
    
    /// <summary>
    /// Ширина главного окна в пикселях.
    /// Значение по умолчанию оптимизировано для разрешения 1920x1080.
    /// </summary>
    public int WindowW { get; set; } = 1180;
    
    /// <summary>
    /// Высота главного окна в пикселях.
    /// Значение по умолчанию оптимизировано для разрешения 1920x1080.
    /// </summary>
    public int WindowH { get; set; } = 790;
    
    /// <summary>
    /// Флаг максимизированного состояния главного окна.
    /// Если true, окно будет запущено в развернутом виде.
    /// </summary>
    public bool WindowMaximized { get; set; } = false;
    
    /// <summary>
    /// Расстояние в пикселях от левого края до разделителя главного окна.
    /// Используется для настройки ширины панели навигации.
    /// </summary>
    public int SplitterDistance { get; set; } = 450;
    
    /// <summary>
    /// Словарь, хранящий ширины колонок грида заявок.
    /// Ключ: имя колонки, Значение: ширина в пикселях.
    /// Позволяет сохранять пользовательские настройки отображения таблицы.
    /// </summary>
    public Dictionary<string, int> GridColWidths { get; set; } = new();

    /// <summary>
    /// Режим выбора UI‑темы в WPF (Auto/Light/Dark).
    /// Сохраняется в настройках и используется ThemeManager для применения темы.
    /// </summary>
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public UiThemeMode ThemeMode { get; set; } = Defaults.Theme;

    /// <summary>
    /// Ширина окна настроек (SettingsWindow) для конкретного пользователя/машины.
    /// null — использовать значение по умолчанию из XAML.
    /// </summary>
    public double? SettingsWindowWidth { get; set; }

    /// <summary>
    /// Высота окна настроек (SettingsWindow) для конкретного пользователя/машины.
    /// null — использовать значение по умолчанию из XAML.
    /// </summary>
    public double? SettingsWindowHeight { get; set; }

    // ==================== Настройки доступа и ролей ====================
    
    /// <summary>
    /// Идентификатор текущего авторизованного пользователя в системе Access.
    /// Используется для определения прав доступа и персонализации интерфейса.
    /// </summary>
    public int? CurrentUserId { get; set; }
    
    /// <summary>
    /// Идентификатор роли администратора в системе Access.
    /// Пользователи с этой ролью получают расширенные права доступа.
    /// </summary>
    public int? AdminRoleId { get; set; }

    // ==================== Настройки SMTP для отправки Email ====================
    
    /// <summary>
    /// Адрес SMTP сервера для отправки электронной почты.
    /// Пример: smtp.gmail.com, smtp.mail.ru
    /// </summary>
    public string SmtpHost { get; set; } = string.Empty;
    
    /// <summary>
    /// Порт SMTP сервера.
    /// Типичные значения: 587 (STARTTLS), 465 (SMTPS), 25 (без шифрования).
    /// </summary>
    public int SmtpPort { get; set; } = 587;
    
    /// <summary>
    /// Флаг использования SSL/TLS шифрования для SMTP соединения.
    /// Рекомендуется использовать true для безопасности.
    /// </summary>
    public bool SmtpUseSsl { get; set; } = true;
    
    /// <summary>
    /// Логин для аутентификации на SMTP сервере.
    /// Обычно совпадает с адресом электронной почты.
    /// </summary>
    public string SmtpLogin { get; set; } = string.Empty;
    
    /// <summary>
    /// Пароль для аутентификации на SMTP сервере.
    /// Рекомендуется хранить в защищенном виде (сейчас хранится открыто).
    /// </summary>
    public string SmtpPassword { get; set; } = string.Empty;
    
    /// <summary>
    /// Адрес электронной почты отправителя.
    /// Используется в поле "From" исходящих сообщений.
    /// </summary>
    public string FromEmail { get; set; } = string.Empty;
    
    /// <summary>
    /// Отображаемое имя отправителя.
    /// Показывается получателям вместо email адреса в почтовых клиентах.
    /// </summary>
    public string FromDisplayName { get; set; } = string.Empty;

    // ==================== Настройки IMAP для получения Email ====================
    
    /// <summary>
    /// Адрес IMAP сервера для получения электронной почты.
    /// Пример: imap.gmail.com, imap.mail.ru
    /// </summary>
    public string ImapHost { get; set; } = string.Empty;
    
    /// <summary>
    /// Порт IMAP сервера.
    /// Типичные значения: 993 (с шифрованием), 143 (без шифрования).
    /// </summary>
    public int ImapPort { get; set; } = 993;
    
    /// <summary>
    /// Флаг использования SSL/TLS шифрования для IMAP соединения.
    /// Рекомендуется использовать true для безопасности.
    /// </summary>
    public bool ImapUseSsl { get; set; } = true;
    
    /// <summary>
    /// Логин для аутентификации на IMAP сервере.
    /// Обычно совпадает с адресом электронной почты.
    /// </summary>
    public string ImapLogin { get; set; } = string.Empty;
    
    /// <summary>
    /// Пароль для аутентификации на IMAP сервере.
    /// Рекомендуется хранить в защищенном виде (сейчас хранится открыто).
    /// </summary>
    public string ImapPassword { get; set; } = string.Empty;
    
    /// <summary>
    /// Название папки IMAP для сохранения отправленных сообщений.
    /// Типичные значения: "Sent", "INBOX.Sent", "Отправленные".
    /// </summary>
    public string ImapSentFolder { get; set; } = "Sent";

    // ==================== Настройки адресатов и уведомлений ====================
    
    /// <summary>
    /// Имя основного получателя заявок.
    /// Используется в шаблонах email сообщений.
    /// </summary>
    public string RequestToName { get; set; } = string.Empty;
    
    /// <summary>
    /// Email адрес основного получателя заявок.
    /// Заявки отправляются на этот адрес по умолчанию.
    /// </summary>
    public string RequestToEmail { get; set; } = string.Empty;
    
    /// <summary>
    /// Подпись, добавляемая к исходящим email сообщениям.
    /// Поддерживает HTML разметку.
    /// </summary>
    public string EmailSignature { get; set; } = string.Empty;

    /// <summary>
    /// Список email адресов для уведомлений о создании новых заявок.
    /// Адреса разделяются запятыми или точкой с запятой.
    /// Если пуст, используется AdminNotifyEmail как fallback.
    /// </summary>
    public string NotifyOnCreateTo { get; set; } = string.Empty;
    
    /// <summary>
    /// Список email адресов для уведомлений об отправке заявок.
    /// Адреса разделяются запятыми или точкой с запятой.
    /// Если пуст, используется AdminNotifyEmail как fallback.
    /// </summary>
    public string NotifyOnSendTo { get; set; } = string.Empty;

    /// <summary>
    /// Создает новый экземпляр AppSettings и загружает настройки из базы данных.
    /// Это основной метод для создания объекта настроек с данными из БД.
    /// </summary>
    /// <param name="db">Экземпляр базы данных для загрузки настроек</param>
    /// <returns>Новый экземпляр AppSettings с загруженными настройками</returns>
    public static AppSettings Load(AppDatabase db)
    {
        var s = new AppSettings();
        s.LoadFromDb(db);
        return s;
    }

    /// <summary>
    /// Сохраняет текущие настройки в базу данных.
    /// Вызывается при изменении настроек пользователем или при закрытии приложения.
    /// </summary>
    /// <param name="db">Экземпляр базы данных для сохранения настроек</param>
    public void Save(AppDatabase db)
    {
        SaveToDb(db);
    }

    /// <summary>
    /// Загружает настройки из базы данных AppDatabase в текущий экземпляр.
    /// Для каждой настройки выполняется:
    /// - попытка загрузки из БД
    /// - использование значения по умолчанию в случае отсутствия или ошибки парсинга
    /// - безопасная обработка исключений с fallback на стандартные значения
    /// 
    /// Специальная обработка предусмотрена для:
    /// - JSON сериализованных данных (массивы, словари)
    /// - Логических значений, хранящихся как "0"/"1"
    /// - Целочисленных значений с fallback на стандартные значения
    /// - Enum значений с безопасным парсингом
    /// </summary>
    /// <param name="db">База данных для загрузки настроек</param>
    private void LoadFromDb(AppDatabase db)
    {
        DbPath = db.GetSetting("DbPath", Defaults.DbPath);
        DbPassword = db.GetSetting("DbPassword", Defaults.DbPassword);
        OutputDir = db.GetSetting("OutputDir", Defaults.OutputDir);
        CompanyName = db.GetSetting("CompanyName", Defaults.CompanyName);
        AdminNotifyEmail = db.GetSetting("AdminNotifyEmail", string.Empty);

        // Загружаем тему - сначала пробуем новый ключ, потом старый для обратной совместимости
        var themeModeStr = db.GetSetting("ThemeMode", string.Empty);
        if (string.IsNullOrEmpty(themeModeStr))
        {
            // Обратная совместимость: пробуем старый ключ UiThemeMode
            themeModeStr = db.GetSetting("UiThemeMode", UiThemeMode.Auto.ToString());
        }
        
        if (Enum.TryParse<UiThemeMode>(themeModeStr, out var themeMode))
        {
            // Обрабатываем устаревший Material как Light
            #pragma warning disable CS0618
            ThemeMode = themeMode == UiThemeMode.Material ? UiThemeMode.Light : themeMode;
            #pragma warning restore CS0618
        }
        else
        {
            ThemeMode = UiThemeMode.Auto;
        }

        AttachInvoiceByDefault = db.GetSetting("AttachInvoiceByDefault", "1") == "1";
        OpenPdfAfterExport = db.GetSetting("OpenPdfAfterExport", "0") == "1";
        SendByEmail = db.GetSetting("SendByEmail", "0") == "1";
        AutoSetOrderedOnSend = db.GetSetting("AutoSetOrderedOnSend", "0") == "1";

        // Геометрия
        WindowX = TryInt(db.GetSetting("WindowX", "-1"), -1);
        WindowY = TryInt(db.GetSetting("WindowY", "-1"), -1);
        WindowW = TryInt(db.GetSetting("WindowW", "1180"), 1180);
        WindowH = TryInt(db.GetSetting("WindowH", "790"), 790);
        WindowMaximized = db.GetSetting("WindowMaximized", "0") == "1";
        SplitterDistance = TryInt(db.GetSetting("SplitterDistance", "450"), 450);

        var settingsWidthStr = db.GetSetting("SettingsWindowWidth", string.Empty);
        if (double.TryParse(settingsWidthStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var w))
            SettingsWindowWidth = w;

        var settingsHeightStr = db.GetSetting("SettingsWindowHeight", string.Empty);
        if (double.TryParse(settingsHeightStr, NumberStyles.Float, CultureInfo.InvariantCulture, out var h))
            SettingsWindowHeight = h;

        var addrsJson = db.GetSetting("DeliveryAddresses", string.Empty);
        if (!string.IsNullOrWhiteSpace(addrsJson))
        {
            try
            {
                DeliveryAddresses = JsonSerializer.Deserialize<string[]>(addrsJson)
                                    ?? (string[])Defaults.DeliveryBook.Clone();
            }
            catch
            {
                DeliveryAddresses = (string[])Defaults.DeliveryBook.Clone();
            }
        }

        // GridColWidths
        var gridJson = db.GetSetting("GridColWidths", string.Empty);
        if (!string.IsNullOrWhiteSpace(gridJson))
        {
            try
            {
                GridColWidths = JsonSerializer.Deserialize<Dictionary<string, int>>(gridJson) ?? new();
            }
            catch
            {
                GridColWidths = new();
            }
        }

        var lastSuppJson = db.GetSetting("LastSupplierContact", string.Empty);
        if (!string.IsNullOrWhiteSpace(lastSuppJson))
        {
            try
            {
                LastSupplierContact = JsonSerializer.Deserialize<Dictionary<int, int>>(lastSuppJson) ?? new();
            }
            catch
            {
                LastSupplierContact = new();
            }
        }

        var lastPers = db.GetSetting("LastCustomerPersonId", string.Empty);
        if (int.TryParse(lastPers, out var pid))
            LastCustomerPersonId = pid;

        var curUserStr = db.GetSetting("CurrentUserId", string.Empty);
        if (int.TryParse(curUserStr, out var uid)) CurrentUserId = uid;

        var adminRoleStr = db.GetSetting("AdminRoleId", string.Empty);
        if (int.TryParse(adminRoleStr, out var arid)) AdminRoleId = arid;

        SmtpHost = db.GetSetting("SmtpHost", string.Empty);
        SmtpPort = TryInt(db.GetSetting("SmtpPort", "587"), 587);
        SmtpUseSsl = db.GetSetting("SmtpUseSsl", "1") == "1";
        SmtpLogin = db.GetSetting("SmtpLogin", string.Empty);
        SmtpPassword = db.GetSetting("SmtpPassword", string.Empty);
        FromEmail = db.GetSetting("FromEmail", string.Empty);
        FromDisplayName = db.GetSetting("FromDisplayName", string.Empty);

        ImapHost = db.GetSetting("ImapHost", string.Empty);
        ImapPort = TryInt(db.GetSetting("ImapPort", "993"), 993);
        ImapUseSsl = db.GetSetting("ImapUseSsl", "1") == "1";
        ImapLogin = db.GetSetting("ImapLogin", string.Empty);
        ImapPassword = db.GetSetting("ImapPassword", string.Empty);
        ImapSentFolder = db.GetSetting("ImapSentFolder", "Sent");

        RequestToName = db.GetSetting("RequestToName", string.Empty);
        RequestToEmail = db.GetSetting("RequestToEmail", string.Empty);
        EmailSignature = db.GetSetting("EmailSignature", string.Empty);
        NotifyOnCreateTo = db.GetSetting("NotifyOnCreateTo", string.Empty);
        NotifyOnSendTo = db.GetSetting("NotifyOnSendTo", string.Empty);

        if (string.IsNullOrWhiteSpace(NotifyOnCreateTo) && !string.IsNullOrWhiteSpace(AdminNotifyEmail))
            NotifyOnCreateTo = AdminNotifyEmail;
        if (string.IsNullOrWhiteSpace(NotifyOnSendTo) && !string.IsNullOrWhiteSpace(AdminNotifyEmail))
            NotifyOnSendTo = AdminNotifyEmail;
    }

    /// <summary>
    /// Сохраняет все настройки из текущего экземпляра в базу данных AppDatabase.
    /// Каждая настройка записывается в таблицу Settings с использованием ключ-значение пар.
    /// Сложные типы (массивы, словари) сериализуются в JSON формат.
    /// Логические значения сохраняются как "0" и "1" для совместимости.
    /// Настройки JSON сериализуются с поддержкой кириллицы и без форматирования для экономии места.
    /// </summary>
    /// <param name="db">База данных для сохранения настроек</param>
    private void SaveToDb(AppDatabase db)
    {
        db.SetSetting("DbPath", DbPath);
        db.SetSetting("DbPassword", DbPassword);
        db.SetSetting("OutputDir", OutputDir);
        db.SetSetting("CompanyName", CompanyName);
        db.SetSetting("AdminNotifyEmail", AdminNotifyEmail);

        db.SetSetting("ThemeMode", ThemeMode.ToString());
        db.SetSetting("AttachInvoiceByDefault", AttachInvoiceByDefault ? "1" : "0");
        db.SetSetting("OpenPdfAfterExport", OpenPdfAfterExport ? "1" : "0");
        db.SetSetting("SendByEmail", SendByEmail ? "1" : "0");
        db.SetSetting("AutoSetOrderedOnSend", AutoSetOrderedOnSend ? "1" : "0");

        db.SetSetting("WindowX", WindowX.ToString());
        db.SetSetting("WindowY", WindowY.ToString());
        db.SetSetting("WindowW", WindowW.ToString());
        db.SetSetting("WindowH", WindowH.ToString());
        db.SetSetting("WindowMaximized", WindowMaximized ? "1" : "0");
        db.SetSetting("SplitterDistance", SplitterDistance.ToString());

        if (SettingsWindowWidth.HasValue)
            db.SetSetting("SettingsWindowWidth", SettingsWindowWidth.Value.ToString(CultureInfo.InvariantCulture));
        else
            db.SetSetting("SettingsWindowWidth", string.Empty);

        if (SettingsWindowHeight.HasValue)
            db.SetSetting("SettingsWindowHeight", SettingsWindowHeight.Value.ToString(CultureInfo.InvariantCulture));
        else
            db.SetSetting("SettingsWindowHeight", string.Empty);

        var opts = new JsonSerializerOptions
        {
            WriteIndented = false,
            Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic)
        };

        db.SetSetting("DeliveryAddresses", JsonSerializer.Serialize(DeliveryAddresses ?? Array.Empty<string>(), opts));
        db.SetSetting("GridColWidths", JsonSerializer.Serialize(GridColWidths ?? new Dictionary<string, int>(), opts));
        db.SetSetting("LastSupplierContact", JsonSerializer.Serialize(LastSupplierContact ?? new Dictionary<int, int>(), opts));

        db.SetSetting("LastCustomerPersonId", LastCustomerPersonId?.ToString() ?? string.Empty);
        db.SetSetting("CurrentUserId", CurrentUserId?.ToString() ?? string.Empty);
        db.SetSetting("AdminRoleId", AdminRoleId?.ToString() ?? string.Empty);

        db.SetSetting("SmtpHost", SmtpHost);
        db.SetSetting("SmtpPort", SmtpPort.ToString());
        db.SetSetting("SmtpUseSsl", SmtpUseSsl ? "1" : "0");
        db.SetSetting("SmtpLogin", SmtpLogin);
        db.SetSetting("SmtpPassword", SmtpPassword);
        db.SetSetting("FromEmail", FromEmail);
        db.SetSetting("FromDisplayName", FromDisplayName);

        db.SetSetting("ImapHost", ImapHost);
        db.SetSetting("ImapPort", ImapPort.ToString());
        db.SetSetting("ImapUseSsl", ImapUseSsl ? "1" : "0");
        db.SetSetting("ImapLogin", ImapLogin);
        db.SetSetting("ImapPassword", ImapPassword);
        db.SetSetting("ImapSentFolder", ImapSentFolder);

        db.SetSetting("RequestToName", RequestToName);
        db.SetSetting("RequestToEmail", RequestToEmail);
        db.SetSetting("EmailSignature", EmailSignature);
        db.SetSetting("NotifyOnCreateTo", NotifyOnCreateTo);
        db.SetSetting("NotifyOnSendTo", NotifyOnSendTo);
    }

    /// <summary>
    /// Безопасное преобразование строки в целое число с fallback значением.
    /// Используется для загрузки целочисленных настроек из базы данных.
    /// Если парсинг не удается, возвращается значение по умолчанию.
    /// </summary>
    /// <param name="s">Строка для парсинга</param>
    /// <param name="def">Значение по умолчанию при ошибке парсинга</param>
    /// <returns>Целое число или значение по умолчанию</returns>
    private static int TryInt(string? s, int def)
        => int.TryParse(s, out var v) ? v : def;

    /// <summary>
    /// Получает полный путь к базе данных с раскрытием переменных среды.
    /// Путь обрабатывается для поддержки относительных путей и переменных среды.
    /// </summary>
    /// <returns>Полный путь к базе данных</returns>
    public string GetResolvedDbPath() => ResolvePath(DbPath);
    
    /// <summary>
    /// Получает полный путь к директории вывода с раскрытием переменных среды.
    /// Путь обрабатывается для поддержки относительных путей и переменных среды.
    /// </summary>
    /// <returns>Полный путь к директории вывода</returns>
    public string GetResolvedOutputDir() => ResolvePath(OutputDir);

    /// <summary>
    /// Обрабатывает путь для поддержки переменных среды и относительных путей.
    /// 
    /// Алгоритм обработки:
    /// 1. Если путь пустой или null, возвращает пустую строку
    /// 2. Раскрывает переменные среды (например, %APPDATA%, %TEMP%)
    /// 3. Если путь относительный, делает его абсолютным относительно директории приложения
    /// 4. Возвращает полный обработанный путь
    /// </summary>
    /// <param name="p">Исходный путь (может содержать переменные среды)</param>
    /// <returns>Обработанный полный путь</returns>
    private static string ResolvePath(string? p)
    {
        if (string.IsNullOrWhiteSpace(p)) return string.Empty;
        p = Environment.ExpandEnvironmentVariables(p);
        if (!Path.IsPathRooted(p))
            p = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, p));
        return p;
    }
}
