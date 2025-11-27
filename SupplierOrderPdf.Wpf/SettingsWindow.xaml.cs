using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Win32;
using MimeKit;
using SupplierOrderPdf.Core;
using SupplierOrderPdf.Wpf;
using System;
using System.IO;
using System.Windows;

namespace SupplierOrderPdf;

/// <summary>
/// Окно настроек приложения SupplierOrderPdf.
/// 
/// Основные функции:
/// - Настройка путей к базам данных Access и выходным директориям
/// - Конфигурация SMTP сервера для отправки email
/// - Конфигурация IMAP сервера для сохранения отправленных писем
/// - Настройка адресатов и уведомлений
/// - Тестирование SMTP и IMAP подключений
/// - Отправка тестовых писем
/// 
/// Интерфейс разделен на секции:
/// - Общие настройки (БД, папки, поведение)
/// - SMTP настройки (сервер, порт, аутентификация)
/// - IMAP настройки (сервер для сохранения отправленных писем)
/// - Адресаты и уведомления (куда отправлять заявки и уведомления)
/// 
/// Окно открывается модально из главного окна и позволяет сохранить
/// изменения настроек, которые применяются ко всему приложению.
/// </summary>
public partial class SettingsWindow : Window
{
    private readonly AppDatabase _db;
    private readonly AppSettings _settings;
    private readonly AccessUser _currentUser;

    /// <summary>
    /// Индекс выбранной темы в ComboBox.
    /// 0 = Auto, 1 = Light, 2 = Dark
    /// </summary>
    public int ThemeModeIndex
    {
        get
        {
            #pragma warning disable CS0618
            // Обрабатываем устаревший Material как Light
            if (_settings.ThemeMode == UiThemeMode.Material)
                return 1; // Light
            #pragma warning restore CS0618
            
            return _settings.ThemeMode switch
            {
                UiThemeMode.Auto  => 0,
                UiThemeMode.Light => 1,
                UiThemeMode.Dark  => 2,
                _ => 0
            };
        }
        set
        {
            _settings.ThemeMode = value switch
            {
                0 => UiThemeMode.Auto,
                1 => UiThemeMode.Light,
                2 => UiThemeMode.Dark,
                _ => UiThemeMode.Auto
            };

            // Применяем тему немедленно для предпросмотра
            ThemeManager.ApplyTheme(_settings.ThemeMode);
        }
    }

    public SettingsWindow(AppDatabase db, AppSettings settings, AccessUser currentUser)
    {
        InitializeComponent();
        _db = db;
        _settings = settings;
        _currentUser = currentUser;

        DataContext = this;

        if (_settings.SettingsWindowWidth.HasValue)
            Width = _settings.SettingsWindowWidth.Value;
        if (_settings.SettingsWindowHeight.HasValue)
            Height = _settings.SettingsWindowHeight.Value;

        // --- общие ---
        DbPathTextBox.Text = _settings.DbPath ?? string.Empty;
        DbPasswordBox.Password = _settings.DbPassword ?? string.Empty;
        OutputDirTextBox.Text = _settings.OutputDir ?? string.Empty;

        AttachInvoiceDefaultCheckBox.IsChecked = _settings.AttachInvoiceByDefault;
        OpenPdfAfterExportCheckBox.IsChecked = _settings.OpenPdfAfterExport;

        // info по отправителю
        FromInfoText.Text = _currentUser.ToEmailFromInfo();

        // --- SMTP ---
        SmtpHostTextBox.Text = _settings.SmtpHost ?? string.Empty;
        SmtpPortTextBox.Text = _settings.SmtpPort > 0 ? _settings.SmtpPort.ToString() : "587";
        SmtpSslCheckBox.IsChecked = _settings.SmtpUseSsl;
        SmtpLoginTextBox.Text = _settings.SmtpLogin ?? string.Empty;
        SmtpPasswordBox.Password = _settings.SmtpPassword ?? string.Empty;

        // --- IMAP ---
        ImapHostTextBox.Text = _settings.ImapHost ?? string.Empty;
        ImapPortTextBox.Text = _settings.ImapPort > 0 ? _settings.ImapPort.ToString() : "993";
        ImapSslCheckBox.IsChecked = _settings.ImapUseSsl;
        ImapSentFolderTextBox.Text = string.IsNullOrWhiteSpace(_settings.ImapSentFolder)
            ? "Sent"
            : _settings.ImapSentFolder;

        // --- адресаты / уведомления ---
        RequestToNameTextBox.Text = _settings.RequestToName ?? string.Empty;
        RequestToEmailTextBox.Text = _settings.RequestToEmail ?? string.Empty;
        AdminNotifyEmailTextBox.Text = _settings.AdminNotifyEmail ?? string.Empty;
        NotifyOnCreateToTextBox.Text = _settings.NotifyOnCreateTo ?? string.Empty;
        NotifyOnSendToTextBox.Text = _settings.NotifyOnSendTo ?? string.Empty;
    }

    // ---------- ОБЩИЕ ----------

    private void BrowseDbButton_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "База Access (*.mdb;*.accdb)|*.mdb;*.accdb|Все файлы (*.*)|*.*"
        };
        if (dlg.ShowDialog(this) == true)
            DbPathTextBox.Text = dlg.FileName;
    }

    private void BrowseOutputButton_Click(object sender, RoutedEventArgs e)
    {
        // Хак через OpenFileDialog: выбираем "файл", а реально берем путь к папке
        var dlg = new OpenFileDialog
        {
            CheckFileExists = false,
            ValidateNames = false,
            FileName = "Выбор папки"
        };

        if (!string.IsNullOrWhiteSpace(OutputDirTextBox.Text) &&
            Directory.Exists(OutputDirTextBox.Text))
        {
            dlg.InitialDirectory = OutputDirTextBox.Text;
        }
        else
        {
            dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        if (dlg.ShowDialog(this) == true)
        {
            var folder = Path.GetDirectoryName(dlg.FileName);
            if (!string.IsNullOrEmpty(folder))
                OutputDirTextBox.Text = folder;
        }
    }

    // ---------- КНОПКИ ОК / ОТМЕНА ----------

    private void OkButton_Click(object sender, RoutedEventArgs e)
    {
        if (!int.TryParse(SmtpPortTextBox.Text.Trim(), out var smtpPort))
            smtpPort = 587;
        if (!int.TryParse(ImapPortTextBox.Text.Trim(), out var imapPort))
            imapPort = 993;

        _settings.DbPath = DbPathTextBox.Text.Trim();
        _settings.DbPassword = DbPasswordBox.Password;
        _settings.OutputDir = OutputDirTextBox.Text.Trim();
        _settings.AttachInvoiceByDefault = AttachInvoiceDefaultCheckBox.IsChecked == true;
        _settings.OpenPdfAfterExport = OpenPdfAfterExportCheckBox.IsChecked == true;

        _settings.SmtpHost = SmtpHostTextBox.Text.Trim();
        _settings.SmtpPort = smtpPort;
        _settings.SmtpUseSsl = SmtpSslCheckBox.IsChecked == true;
        _settings.SmtpLogin = SmtpLoginTextBox.Text.Trim();
        _settings.SmtpPassword = SmtpPasswordBox.Password;

        _settings.ImapHost = ImapHostTextBox.Text.Trim();
        _settings.ImapPort = imapPort;
        _settings.ImapUseSsl = ImapSslCheckBox.IsChecked == true;
        _settings.ImapSentFolder = ImapSentFolderTextBox.Text.Trim();

        _settings.RequestToName = RequestToNameTextBox.Text.Trim();
        _settings.RequestToEmail = RequestToEmailTextBox.Text.Trim();
        _settings.AdminNotifyEmail = AdminNotifyEmailTextBox.Text.Trim();
        _settings.NotifyOnCreateTo = NotifyOnCreateToTextBox.Text.Trim();
        _settings.NotifyOnSendTo = NotifyOnSendToTextBox.Text.Trim();

        _settings.Save(_db);

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    // ---------- ПРОВЕРКА SMTP ----------

    private async void TestSmtpButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TestSmtpButton.IsEnabled = false;

            if (!int.TryParse(SmtpPortTextBox.Text.Trim(), out var port))
                port = 587;

            using var smtp = new SmtpClient();
            var secure = SmtpSslCheckBox.IsChecked == true
                ? SecureSocketOptions.StartTlsWhenAvailable
                : SecureSocketOptions.None;

            await smtp.ConnectAsync(SmtpHostTextBox.Text.Trim(), port, secure);

            if (!string.IsNullOrWhiteSpace(SmtpLoginTextBox.Text))
                await smtp.AuthenticateAsync(SmtpLoginTextBox.Text.Trim(), SmtpPasswordBox.Password);

            await smtp.DisconnectAsync(true);

            MessageBox.Show(this, "Подключение к SMTP прошло успешно.",
                "Проверка SMTP", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при проверке SMTP:\n" + ex.Message,
                "Проверка SMTP", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TestSmtpButton.IsEnabled = true;
        }
    }

    private async void SendTestMailButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            SendTestMailButton.IsEnabled = false;

            var fromEmail = _currentUser.Email;
            if (string.IsNullOrWhiteSpace(fromEmail))
            {
                MessageBox.Show(this,
                    "В карточке сотрудника не указан e-mail. Тестовое письмо отправить нельзя.",
                    "Тестовое письмо", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var msg = new MimeMessage();
            var fromName = _currentUser.PersonName;
            msg.From.Add(new MailboxAddress(fromName, fromEmail));

            var to = RequestToEmailTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(to))
                to = fromEmail;

            msg.To.Add(MailboxAddress.Parse(to));
            msg.Subject = "Тестовое письмо для проверки SMTP";
            msg.Body = new TextPart("plain")
            {
                Text = "Это тестовое письмо, отправленное из окна настроек."
            };

            if (!int.TryParse(SmtpPortTextBox.Text.Trim(), out var port))
                port = 587;

            using var smtp = new SmtpClient();
            var secure = SmtpSslCheckBox.IsChecked == true
                ? SecureSocketOptions.StartTlsWhenAvailable
                : SecureSocketOptions.None;

            await smtp.ConnectAsync(SmtpHostTextBox.Text.Trim(), port, secure);

            if (!string.IsNullOrWhiteSpace(SmtpLoginTextBox.Text))
                await smtp.AuthenticateAsync(SmtpLoginTextBox.Text.Trim(), SmtpPasswordBox.Password);

            await smtp.SendAsync(msg);
            await smtp.DisconnectAsync(true);

            MessageBox.Show(this, "Тестовое письмо отправлено.",
                "Тестовое письмо", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при отправке тестового письма:\n" + ex.Message,
                "Тестовое письмо", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SendTestMailButton.IsEnabled = true;
        }
    }

    // ---------- ПРОВЕРКА IMAP ----------

    private async void TestImapButton_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TestImapButton.IsEnabled = false;

            if (string.IsNullOrWhiteSpace(ImapHostTextBox.Text))
            {
                MessageBox.Show(this, "Не указан IMAP сервер.",
                    "Проверка IMAP", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!int.TryParse(ImapPortTextBox.Text.Trim(), out var port))
                port = 993;

            using var imap = new ImapClient();
            await imap.ConnectAsync(ImapHostTextBox.Text.Trim(), port, ImapSslCheckBox.IsChecked == true);

            // логин/пароль берём такие же, как для SMTP
            if (!string.IsNullOrWhiteSpace(SmtpLoginTextBox.Text))
                await imap.AuthenticateAsync(SmtpLoginTextBox.Text.Trim(), SmtpPasswordBox.Password);

            var folderName = string.IsNullOrWhiteSpace(ImapSentFolderTextBox.Text)
                ? "Sent"
                : ImapSentFolderTextBox.Text.Trim();

            try
            {
                var folder = await imap.GetFolderAsync(folderName);
                await folder.OpenAsync(MailKit.FolderAccess.ReadOnly);
                await folder.CloseAsync();
            }
            catch
            {
                MessageBox.Show(this,
                    $"IMAP подключен, но папка '{folderName}' не найдена или недоступна.",
                    "Проверка IMAP", MessageBoxButton.OK, MessageBoxImage.Warning);
                await imap.DisconnectAsync(true);
                return;
            }

            await imap.DisconnectAsync(true);

            MessageBox.Show(this, "Подключение к IMAP прошло успешно.",
                "Проверка IMAP", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при проверке IMAP:\n" + ex.Message,
                "Проверка IMAP", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TestImapButton.IsEnabled = true;
        }
    }

    private void SettingsWindow_OnClosed(object? sender, EventArgs e)
    {
        _settings.SettingsWindowWidth = Width;
        _settings.SettingsWindowHeight = Height;
        _settings.Save(_db);
    }
}
