using System;
using System.IO;
using System.Windows;
using System.Windows.Threading;
using SupplierOrderPdf.Core;
using SupplierOrderPdf.Wpf;

namespace SupplierOrderPdf;

/// <summary>
/// Главный класс приложения SupplierOrderPdf.
/// Наследует от Application и управляет жизненным циклом WPF приложения.
/// 
/// Основные функции:
/// - Инициализация приложения и обработка запуска
/// - Создание и управление базой данных SQLite для настроек
/// - Обработка процесса авторизации через LoginWindow
/// - Создание главного окна MainWindow после успешной авторизации
/// - Обработка глобальных исключений и логирование ошибок
/// - Управление режимом завершения работы приложения
/// 
/// Приложение использует двухэтапную схему запуска:
/// 1. Сначала показывается LoginWindow для авторизации пользователя
/// 2. После успешной авторизации создается и показывается MainWindow
/// </summary>
public partial class App : Application
{
    /// <summary>
    /// Экземпляр базы данных SQLite для хранения настроек приложения и журнала заявок.
    /// Создается при запуске приложения и используется всеми компонентами.
    /// Может быть null до завершения инициализации.
    /// </summary>
    private AppDatabase? _db;
    
    /// <summary>
    /// Экземпляр настроек приложения, загруженных из базы данных.
    /// Содержит все пользовательские и системные настройки.
    /// Может быть null до завершения загрузки настроек.
    /// </summary>
    private AppSettings? _settings;

    /// <summary>
    /// Конструктор главного класса приложения.
    /// Вызывается при создании экземпляра приложения до его запуска.
    /// Настраивает глобальные обработчики исключений для всего приложения.
    /// </summary>
    public App()
    {
        // Регистрируем обработчик исключений WPF-диспетчера (UI поток)
        // Перехватывает исключения, возникающие в UI элементах и при обработке команд
        this.DispatcherUnhandledException += App_DispatcherUnhandledException;

        // Регистрируем обработчик критических исключений .NET (не UI поток)
        // Перехватывает исключения в фоновых потоках, включая ThreadException
        AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
    }

    /// <summary>
    /// Главный метод инициализации приложения при запуске.
    /// Вызывается автоматически при старте приложения после конструктора.
    /// Выполняет полную инициализацию приложения, включая:
    /// - Создание и инициализацию базы данных SQLite
    /// - Загрузку пользовательских настроек
    /// - Отображение окна авторизации
    /// - Создание главного окна после успешной авторизации
    /// 
    /// Алгоритм запуска:
    /// 1. Настройка режима завершения работы приложения
    /// 2. Создание базы данных и загрузка настроек
    /// 3. Показ LoginWindow для авторизации пользователя
    /// 4. При успешной авторизации - создание MainWindow
    /// 5. При отмене авторизации - завершение работы приложения
    /// 6. Обработка всех возможных исключений
    /// </summary>
    /// <param name="e">Аргументы события запуска, содержащие параметры командной строки</param>
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        // Устанавливаем начальный режим завершения работы приложения
        // OnExplicitShutdown означает, что приложение завершится только при вызове Shutdown()
        // Это критически важно, чтобы приложение не закрылось автоматически после LoginWindow
        ShutdownMode = ShutdownMode.OnExplicitShutdown;

        try
        {
            Log($"OnStartup. ShutdownMode: {ShutdownMode}");

            // Вычисляем путь к файлу базы данных SQLite
            string sqlitePath = ComputeSqlitePath();
            
            // Создаем экземпляр базы данных
            _db = new AppDatabase(sqlitePath);
            
            // Инициализируем структуру БД (создаем таблицы если их нет)
            _db.Initialize();
            
            // Загружаем настройки приложения из базы данных
            _settings = AppSettings.Load(_db);

            // Инициализируем ThemeManager с сохранённой темой из настроек
            // ThemeManager применит тему и подпишется на изменения системной темы Windows
            ThemeManager.Initialize(_settings.ThemeMode);

            // Создаем окно авторизации, передавая ему настройки для работы
            var loginWindow = new LoginWindow(_settings);
            Log($"LoginWindow created. MainWindow is: {MainWindow}");

            // Показываем окно авторизации как модальный диалог
            // Функция ShowDialog() блокирует выполнение до закрытия окна
            bool? result = loginWindow.ShowDialog();
            Log($"LoginWindow closed. Result: {result}. ShutdownMode: {ShutdownMode}");

            // Проверяем результат авторизации
            if (result == true && loginWindow.SelectedUser != null)
            {
                Log("Creating MainWindow...");
                
                // Создаем главное окно приложения, передавая ему БД, настройки и авторизованного пользователя
                var main = new MainWindow(_db, _settings, loginWindow.SelectedUser);
                
                // Устанавливаем главное окно как основное окно приложения
                MainWindow = main;

                // Показываем главное окно (делаем его видимым)
                main.Show();

                // Меняем режим завершения работы на закрытие по главному окну
                // Теперь приложение будет автоматически завершаться при закрытии главного окна
                ShutdownMode = ShutdownMode.OnMainWindowClose;
                Log("MainWindow shown. ShutdownMode set to OnMainWindowClose.");
            }
            else
            {
                // Пользователь отменил авторизацию или вход не удался
                Log("User cancelled login. Shutting down.");
                
                // Явно завершаем работу приложения
                Shutdown();
            }
        }
        catch (Exception ex)
        {
            // Обрабатываем любые исключения во время инициализации
            LogException("OnStartup", ex);
            
            // Завершаем приложение с кодом ошибки (-1)
            Shutdown(-1);
        }
    }

    /// <summary>
    /// Записывает информационное сообщение в файл логирования приложения.
    /// Создает файл wpf_error.log в директории приложения и дописывает сообщение с меткой времени.
    /// Используется для отладочной информации и отслеживания хода выполнения приложения.
    /// Метод безопасен - любые исключения при записи игнорируются.
    /// </summary>
    /// <param name="message">Текст сообщения для записи в лог</param>
    private void Log(string message)
    {
        try
        {
            // Формируем путь к файлу лога в директории приложения
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wpf_error.log");
            
            // Записываем сообщение с меткой времени и префиксом INFO
            File.AppendAllText(path, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [INFO] {message}\r\n");
        }
        catch { }
        // Исключения при записи в лог игнорируются, чтобы не прерывать работу приложения
    }

    /// <summary>
    /// Вычисляет путь к файлу базы данных SQLite для хранения настроек приложения.
    /// 
    /// Алгоритм определения пути:
    /// 1. Начинаем с директории исполняемого файла приложения
    /// 2. Пытаемся найти базу данных Access по пути из настроек по умолчанию
    /// 3. Если база данных Access найдена, используем её директорию для SQLite
    /// 4. Если база данных Access не найдена, остаемся в директории приложения
    /// 5. Возвращаем полный путь к файлу "SupplierOrderPdf.sqlite"
    /// 
    /// Такая логика обеспечивает, что файлы приложения и базы данных
    /// находятся в одной директории с базой данных Access.
    /// </summary>
    /// <returns>Полный путь к файлу базы данных SQLite</returns>
    private string ComputeSqlitePath()
    {
        string dir = AppDomain.CurrentDomain.BaseDirectory;

        try
        {
            // Получаем путь к базе данных Access из настроек по умолчанию
            string accessPath = AppSettings.Defaults.DbPath;
            
            // Раскрываем переменные среды в пути (например, %APPDATA%)
            accessPath = Environment.ExpandEnvironmentVariables(accessPath);
            
            // Если путь относительный, делаем его абсолютным относительно директории приложения
            if (!Path.IsPathRooted(accessPath))
                accessPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, accessPath));

            // Проверяем существование файла базы данных Access
            if (File.Exists(accessPath))
            {
                // Получаем директорию базы данных Access
                var d = Path.GetDirectoryName(accessPath);
                if (!string.IsNullOrWhiteSpace(d))
                    dir = d; // Используем директорию Access БД для SQLite файла
            }
        }
        catch
        {
            // В случае любых ошибок остаемся в директории приложения (fallback)
        }

        // Возвращаем полный путь к файлу SQLite в выбранной директории
        return Path.Combine(dir, "SupplierOrderPdf.sqlite");
    }

    // ==================== Глобальные обработчики исключений ====================
    
    /// <summary>
    /// Обработчик неожиданных исключений WPF диспетчера (UI поток).
    /// Вызывается при возникновении исключений в UI элементах, обработчиках событий,
    /// обработчиках команд и других операциях пользовательского интерфейса.
    /// 
    /// Исключение помечается как обработанное (e.Handled = true) для предотвращения
    /// аварийного завершения приложения. MessageBox не показывается для избежания
    /// проблем при завершении работы или возникновении рекурсивных ошибок.
    /// </summary>
    /// <param name="sender">Источник события (объект Dispatcher)</param>
    /// <param name="e">Аргументы события с информацией об исключении</param>
    private void App_DispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        // Записываем исключение в лог
        LogException("DispatcherUnhandledException", e.Exception);
        
        // Помечаем исключение как обработанное для предотвращения аварийного завершения
        e.Handled = true;
    }

    /// <summary>
    /// Обработчик критических исключений .NET (не UI поток).
    /// Вызывается при возникновении исключений в фоновых потоках, включая ThreadException.
    /// В отличие от DispatcherUnhandledException, здесь показывается MessageBox пользователю,
    /// так как такие исключения обычно указывают на серьезные проблемы.
    /// </summary>
    /// <param name="sender">Источник события (объект AppDomain)</param>
    /// <param name="e">Аргументы события с информацией об исключении</param>
    private void CurrentDomain_UnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        // Проверяем, что исключение действительно является объектом Exception
        if (e.ExceptionObject is Exception ex)
        {
            // Записываем исключение в лог
            LogException("CurrentDomain_UnhandledException", ex);

            // Показываем MessageBox с информацией об ошибке пользователю
            // Это критическая ошибка, требующая внимания пользователя
            MessageBox.Show("Критическая ошибка в приложении:\n\n" + ex,
                "Ошибка .NET", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// Записывает информацию об исключении в файл логирования.
    /// Создает детальную запись об ошибке с меткой времени, источником и полным стеком вызовов.
    /// Используется всеми глобальными обработчиками исключений для документирования проблем.
    ///
    /// Формат записи в лог:
    /// - Метка времени в формате yyyy-MM-dd HH:mm:ss
    /// - Источник исключения в квадратных скобках
    /// - Полное описание исключения и стек вызовов
    /// - Двойной перенос строки для разделения записей
    ///
    /// Метод безопасен - любые ошибки при записи игнорируются.
    /// </summary>
    /// <param name="source">Источник исключения (имя метода или обработчика)</param>
    /// <param name="ex">Объект исключения для записи в лог</param>
    private void LogException(string source, Exception ex)
    {
        try
        {
            // Формируем путь к файлу лога в директории приложения
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wpf_error.log");
            
            // Записываем детальную информацию об исключении
            // Включаем полный стек вызовов для диагностики проблем
            File.AppendAllText(path,
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{source}]\r\n{ex}\r\n\r\n");
        }
        catch
        {
            // Если файл лога не удается создать или записать в него,
            // приложение продолжает работу без логирования (graceful degradation)
        }
    }

    /// <summary>
    /// Вызывается при завершении работы приложения.
    /// Используется для аккуратного освобождения ресурсов, связанных с темизацией,
    /// в частности для отписки ThemeManager от системных событий Windows.
    /// </summary>
    /// <param name="e">Аргументы события завершения приложения</param>
    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            ThemeManager.Shutdown();
        }
        catch
        {
            // Игнорируем ошибки при завершении ThemeManager, чтобы не мешать штатному завершению приложения
        }

        base.OnExit(e);
    }
}