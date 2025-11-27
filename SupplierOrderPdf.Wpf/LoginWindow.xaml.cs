using System.Data.Odbc;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using SupplierOrderPdf.Core;

namespace SupplierOrderPdf.Wpf;

/// <summary>
/// Окно авторизации пользователей SupplierOrderPdf.
/// 
/// Основные функции:
/// - Загрузка списка пользователей из базы данных Access
/// - Предоставление пользователю выбора учетной записи
/// - Проверка пароля для выбранной учетной записи
/// - Возврат выбранного пользователя в основное приложение
/// 
/// Особенности:
/// - Подключается к базе данных Access через ODBC драйвер
/// - Извлекает информацию о пользователях и связанных с ними персон
/// - Отображает ошибки подключения и загрузки данных
/// - Использует временную проверку паролей (TODO: улучшить безопасность)
/// - Автоматически выбирает первого пользователя в списке
/// 
/// Дизайн ориентирован на простоту использования и быструю авторизацию
/// в корпоративной среде с ограниченным количеством пользователей.
/// </summary>
public partial class LoginWindow : Window
{
    /// <summary>
    /// Настройки приложения, используемые для подключения к базе данных.
    /// Передаются из главного приложения при создании окна.
    /// </summary>
    private readonly AppSettings _settings;
    
    /// <summary>
    /// Список пользователей, загруженных из базы данных Access.
    /// Используется для отображения в комбобоксе выбора пользователя.
    /// </summary>
    private readonly List<AccessUser> _users = new();

    /// <summary>
    /// Выбранный пользователь после успешной авторизации.
    /// Возвращается в основное приложение через это свойство.
    /// Может быть null, если авторизация не удалась или была отменена.
    /// </summary>
    public AccessUser? SelectedUser { get; private set; }

    private static string GetLastUserFilePath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SupplierOrderPdf");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "lastuser.txt");
    }

    private void SaveLastUserLogin(string login)
    {
        try
        {
            var path = GetLastUserFilePath();
            File.WriteAllText(path, login ?? string.Empty);
        }
        catch
        {
            // глушим ошибки — отказ записи не критичен для работы приложения
        }
    }

    private string? LoadLastUserLogin()
    {
        try
        {
            var path = GetLastUserFilePath();
            if (!File.Exists(path))
                return null;
            var login = File.ReadAllText(path).Trim();
            return string.IsNullOrWhiteSpace(login) ? null : login;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Конструктор окна авторизации.
    /// Инициализирует компоненты интерфейса и загружает список пользователей.
    /// </summary>
    /// <param name="settings">Настройки приложения для подключения к БД</param>
    public LoginWindow(AppSettings settings)
    {
        InitializeComponent();
        _settings = settings;

        Loaded += (_, _) =>
        {
            try
            {
                PasswordBox.Focus();
            }
            catch
            {
            }
        };

        LoadUsersFromAccess();
    }

    /// <summary>
    /// Загружает список пользователей из базы данных Access.
    /// Выполняет полный цикл подключения к БД, выполнения запроса и обработки результатов.
    /// 
    /// Этапы работы:
    /// 1. Проверка настроек приложения
    /// 2. Проверка существования файла базы данных Access
    /// 3. Подключение к базе данных через ODBC драйвер
    /// 4. Выполнение SQL запроса для получения пользователей и их данных
    /// 5. Создание объектов AccessUser из результатов запроса
    /// 6. Привязка списка к комбобоксу в интерфейсе
    /// 
    /// В случае ошибок отображает сообщения в элементе ErrorText.
    /// Все исключения перехватываются и логируются через Debug.WriteLine.
    /// </summary>
    private void LoadUsersFromAccess()
    {
        try
        {
            System.Diagnostics.Debug.WriteLine("[LoginWindow] LoadUsersFromAccess: старт вызова");
            
            // Проверяем наличие настроек приложения
            if (_settings == null)
            {
                System.Diagnostics.Debug.WriteLine("[LoginWindow] _settings == null -> прекращаем работу");
                ErrorText.Text = "Внутренняя ошибка настроек.";
                return;
            }

            // Получаем полный путь к базе данных Access с раскрытием переменных среды
            string dbPath = _settings.GetResolvedDbPath();
            System.Diagnostics.Debug.WriteLine($"[LoginWindow] dbPath='{dbPath}'");
            
            // Проверяем, что путь к БД не пустой
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                System.Diagnostics.Debug.WriteLine("[LoginWindow] dbPath пустой");
                ErrorText.Text = "Не задан путь к базе Access. Откройте настройки.";
                return;
            }

            // Проверяем существование файла базы данных
            bool fileExists = System.IO.File.Exists(dbPath);
            System.Diagnostics.Debug.WriteLine($"[LoginWindow] fileExists={fileExists}");
            if (!fileExists)
            {
                ErrorText.Text = "Файл базы Access не найден: " + dbPath;
                return;
            }

            System.Diagnostics.Debug.WriteLine($"[LoginWindow] UserCombo == null? {UserCombo == null}");
            System.Diagnostics.Debug.WriteLine($"[LoginWindow] ErrorText == null? {ErrorText == null}");

            // Формируем строку подключения ODBC для Microsoft Access
            // Включаем драйвер, путь к файлу и пароль из настроек
            string cs = $"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={dbPath};Pwd={_settings.DbPassword};";
            
            // Создаем и открываем подключение к базе данных
            using var conn = new OdbcConnection(cs);
            conn.Open();
            System.Diagnostics.Debug.WriteLine("[LoginWindow] Подключение к Access открыто");

            // Создаем команду SQL для получения пользователей и их персон
            // Используем LEFT JOIN для получения дополнительной информации из таблиц Person и Roles.
            // Для Microsoft Access требуется явно задавать порядок JOIN с помощью скобок.
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT u.ID, u.Name, u.RoleID, u.PersonID, u.Password,
       p.LastName, p.FirstName, p.SecondName,
       p.Email, p.Phone,
       r.Name AS RoleName
FROM (Users AS u
LEFT JOIN Person AS p ON p.ID = u.PersonID)
LEFT JOIN Roles AS r ON r.Id = u.RoleID
WHERE u.Hidden = 0 OR u.Hidden IS NULL
ORDER BY u.Name";

            // Выполняем запрос и читаем результаты
            using var rd = cmd.ExecuteReader();
            System.Diagnostics.Debug.WriteLine("[LoginWindow] Чтение пользователей начато");
            
            while (rd.Read())
            {
                // Создаем объект пользователя из данных базы
                var user = new AccessUser
                {
                    Id = Convert.ToInt32(rd["ID"]),
                    Login = rd["Name"]?.ToString() ?? string.Empty,
                    RoleId = rd["RoleID"] is DBNull ? (int?)null : Convert.ToInt32(rd["RoleID"]),
                    RoleName = rd["RoleName"]?.ToString() ?? string.Empty,
                    PersonId = rd["PersonID"] is DBNull ? (int?)null : Convert.ToInt32(rd["PersonID"]),
                    PersonFirstName = rd["FirstName"]?.ToString() ?? string.Empty,
                    PersonLastName = rd["LastName"]?.ToString() ?? string.Empty,
                    PersonSecondName = rd["SecondName"]?.ToString() ?? string.Empty,
                    Email = rd["Email"]?.ToString() ?? string.Empty,
                    Phone = rd["Phone"]?.ToString() ?? string.Empty,
                    Password = rd["Password"]?.ToString() ?? string.Empty
                };
                _users.Add(user);
            }
            System.Diagnostics.Debug.WriteLine($"[LoginWindow] Загружено пользователей: {_users.Count}");

            // Привязываем список пользователей к комбобоксу
            if (UserCombo == null)
            {
                System.Diagnostics.Debug.WriteLine("[LoginWindow] UserCombo == null при назначении ItemsSource");
            }
            else
            {
                UserCombo.ItemsSource = _users;
                // Автоматически выбираем первого пользователя, если список не пустой
                if (_users.Count > 0)
                {
                    // Пытаемся выбрать пользователя, вошедшего последним на этом ПК
                    var lastLogin = LoadLastUserLogin();
                    if (!string.IsNullOrWhiteSpace(lastLogin))
                    {
                        var idx = _users.FindIndex(u =>
                            !string.IsNullOrWhiteSpace(u.Login) &&
                            string.Equals(u.Login, lastLogin, StringComparison.OrdinalIgnoreCase));
                        if (idx >= 0)
                            UserCombo.SelectedIndex = idx;
                        else
                            UserCombo.SelectedIndex = 0;
                    }
                    else
                    {
                        UserCombo.SelectedIndex = 0;
                    }

                    // Обновляем отображение роли для выбранного пользователя
                    UserCombo_SelectionChanged(UserCombo, null);

                    // Фокус в поле пароля (см. раздел 2)
                    PasswordBox.Focus();
                }
            }
        }
        catch (Exception ex)
        {
            // Логируем исключение для отладки
            System.Diagnostics.Debug.WriteLine("[LoginWindow] Ошибка: " + ex);
            
            // Отображаем сообщение об ошибке пользователю
            if (ErrorText != null)
            {
                ErrorText.Text = "Ошибка загрузки пользователей: " + ex.Message;
            }
        }
    }

    /// <summary>
    /// Обработчик нажатия кнопки "ОК" в окне авторизации.
    /// Выполняет проверку выбранного пользователя и пароля.
    /// 
    /// Алгоритм:
    /// 1. Проверяет наличие выбранного пользователя
    /// 2. Получает введенный пароль из PasswordBox
    /// 3. Выполняет временную проверку пароля (TODO: улучшить безопасность)
    /// 4. При успехе устанавливает SelectedUser и закрывает окно с результатом true
    /// 5. При неудаче отображает сообщение об ошибке
    /// </summary>
    /// <param name="sender">Источник события (кнопка ОК)</param>
    /// <param name="e">Аргументы события</param>
    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        // Проверяем, что пользователь выбран в комбобоксе
        if (UserCombo.SelectedItem is not AccessUser user)
        {
            ErrorText.Text = "Выберите пользователя.";
            return;
        }

        // Получаем введенный пароль
        string password = PasswordBox.Password ?? string.Empty;

        // Временная реализация: пока принимаем любой пароль для отладки,
        // как это было в WinForms версии
        // TODO: реализовать настоящую проверку пароля по данным из Access
        if (string.IsNullOrWhiteSpace(user.Password) || password == user.Password)
        {
            // Успешная авторизация - сохраняем выбранного пользователя
            SelectedUser = user;

            // Запоминаем логин последнего успешно вошедшего пользователя на этом ПК
            SaveLastUserLogin(user.Login ?? string.Empty);

            // Устанавливаем результат диалога в true и закрываем окно
            DialogResult = true;
        }
        else
        {
            // Неверный пароль - показываем ошибку
            ErrorText.Text = "Неверный пароль.";
        }
    }

    /// <summary>
    /// Обработчик нажатия кнопки "Отмена" в окне авторизации.
    /// Закрывает окно с результатом false, сигнализируя об отмене авторизации.
    ///
    /// Этот метод вызывается, когда пользователь решает не входить в систему,
    /// например, нажимает кнопку "Отмена" или закрывает окно.
    /// </summary>
    /// <param name="sender">Источник события (кнопка Отмена)</param>
    /// <param name="e">Аргументы события</param>
    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    /// <summary>
    /// Обработчик изменения выбора пользователя в комбобоксе.
    /// Обновляет отображение роли выбранного пользователя.
    /// </summary>
    /// <param name="sender">Источник события (UserCombo)</param>
    /// <param name="e">Аргументы события</param>
    private void UserCombo_SelectionChanged(object sender, SelectionChangedEventArgs? e)
    {
        if (UserCombo.SelectedItem is AccessUser user)
        {
            RoleText.Text = user.RoleName;
        }
        else
        {
            RoleText.Text = string.Empty;
        }
    }
}
