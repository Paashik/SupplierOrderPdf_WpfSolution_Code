using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using MimeKit;
using SupplierOrderPdf.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace SupplierOrderPdf.Wpf
{
    /// <summary>
    /// Главное окно приложения SupplierOrderPdf.
    /// 
    /// Основные функции:
    /// - Отображение списка заказов из базы данных Access
    /// - Предпросмотр и генерация PDF заявок на закупку
    /// - Отправка заявок по электронной почте
    /// - Управление настройками приложения
    /// - Интеграция с WebView2 для предпросмотра HTML
    /// - Поддержка темного/светлого интерфейса
    /// 
    /// Архитектура:
    /// - Использует ODBC для подключения к Access БД
    /// - MailKit для отправки email через SMTP
    /// - WebView2 для генерации PDF из HTML шаблонов
    /// - SQLite для хранения настроек и журнала операций
    /// 
    /// Пользовательский интерфейс состоит из:
    /// - Списка заказов (OrdersGrid)
    /// - Панели предпросмотра (PreviewBrowser) 
    /// - Поля для ввода основания заявки (BasisTextBox)
    /// - Выбора адреса доставки (DeliveryAddressText)
    /// - Контактов поставщика (SupplierContactCombo)
    /// - Кнопок операций (PdfButton, SendButton)
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly AppDatabase _db;
        private readonly AppSettings _cfg;
        private readonly AccessUser _currentUser;

    private OdbcConnection? _conn;
    private DataTable? _allOrders;
    private int? _currentOrderId;
    private OrderHeader? _currentOrderHeader;
    private int _currentProviderId;

    // Контакт заказчика для текущего заказа (создатель заявки)
    private string _customerFioForCurrentOrder = "";
    private string _customerEmailForCurrentOrder = "";
    private string _customerPhoneForCurrentOrder = "";

    public MainWindow(AppDatabase db, AppSettings settings, AccessUser currentUser)
    {
        InitializeComponent();

        _db = db;
        _cfg = settings;
        _currentUser = currentUser;

        Title = $"Заявка на закупку → PDF (WPF) — {_currentUser.PersonName}";

        InitYearCombo();
        InitStatusCombo();
        InitOrdersGridColumns();

        // Проверка инициализации элементов управления
        if (PreviewBrowser == null)
        {
            MessageBox.Show("PreviewBrowser не инициализирован");
        }

        // Инициализация WebView2
        PreviewBrowser.CoreWebView2InitializationCompleted += OnWebViewInitialized;

        Loaded += async (_, _) =>
        {
            try
            {
                await EnsureWebViewAsync();
                if (PreviewBrowser != null && PreviewBrowser.CoreWebView2 != null)
                {
                    ApplyThemeToPreview();
                    if (PreviewErrorText != null) PreviewErrorText.Visibility = Visibility.Collapsed;
                }
                else
                {
                    if (PreviewErrorText != null) PreviewErrorText.Visibility = Visibility.Visible;
                }
                LoadAddresses();
                LoadOrders();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке данных:\n" + ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        };

        Closed += (_, _) =>
        {
            try { _conn?.Dispose(); } catch { }
        };

        AttachInvoiceCheckBox.IsChecked = _cfg.AttachInvoiceByDefault;
        OpenAfterExportCheckBox.IsChecked = _cfg.OpenPdfAfterExport;

        SendButton.IsEnabled = IsCurrentUserAdmin;
    }

    private bool IsCurrentUserAdmin =>
        _cfg.AdminRoleId.HasValue &&
        _currentUser.RoleId.HasValue &&
        _cfg.AdminRoleId.Value == _currentUser.RoleId.Value;

    // ---------- UI init ----------

    private void InitYearCombo()
    {
        YearCombo.Items.Clear();
        int currentYear = DateTime.Now.Year;
        for (int y = currentYear; y >= currentYear - 10; y--)
            YearCombo.Items.Add(y.ToString());
        if (YearCombo.Items.Count > 0)
            YearCombo.SelectedIndex = 0;
    }

    private void InitStatusCombo()
    {
        StatusCombo.Items.Clear();
        StatusCombo.Items.Add("Все");
        StatusCombo.Items.Add("Актуальные");
        StatusCombo.SelectedIndex = 1;
    }

    private void InitOrdersGridColumns()
    {
        OrdersGrid.Columns.Clear();

        
        OrdersGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "№",
            Binding = new System.Windows.Data.Binding("Num"),
            Width = DataGridLength.Auto
        });
         OrdersGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Поставщик",
            Binding = new System.Windows.Data.Binding("ProviderName"),
            //Width = new DataGridLength(1, DataGridLengthUnitType.Star)
            Width = DataGridLength.Auto
        });
        OrdersGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Дата",
            Binding = new System.Windows.Data.Binding("OrderDate") { StringFormat = "dd.MM.yy" },
            Width = DataGridLength.Auto
        });

        OrdersGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Статус",
            Binding = new System.Windows.Data.Binding("StateName"),
            Width = DataGridLength.Auto
        });
        OrdersGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Заявка",
            Binding = new System.Windows.Data.Binding("RequestInfo"),
            Width = DataGridLength.Auto
        });
        OrdersGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Кто создал",
            Binding = new System.Windows.Data.Binding("RequestCreatedBy"),
            Width = DataGridLength.Auto
        });
    }

    // ---------- Events ----------

    private void ReloadButton_OnClick(object sender, RoutedEventArgs e) => LoadOrders();

    private void SettingsButton_OnClick(object sender, RoutedEventArgs e)
    {
        var win = new SettingsWindow(_db, _cfg, _currentUser) { Owner = this };
        if (win.ShowDialog() == true)
        {
            // перечитать настройки (могли поменять путь к БД/SMTP и т.п.)
            LoadAddresses();
            LoadOrders();
        }
    }

    private void YearCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e) => LoadOrders();
    private void StatusCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e) => ApplyFilter();
    private void FilterBox_OnTextChanged(object sender, TextChangedEventArgs e) => ApplyFilter();

    private async void OrdersGrid_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        await LoadSelectedOrderAsync();
    }

    private void ChooseAddressButton_OnClick(object sender, RoutedEventArgs e)
    {
        var list = _cfg.DeliveryAddresses ?? AppSettings.Defaults.DeliveryBook;
        var dlg = new Window
        {
            Title = "Выбор адреса",
            Owner = this,
            Width = 600,
            Height = 380,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };
        var lb = new ListBox { Margin = new Thickness(8) };
        foreach (var s in list) lb.Items.Add(s);
        var btnOk = new Button { Content = "Выбрать", Width = 100, Margin = new Thickness(8) };
        btnOk.Click += (_, _) => { dlg.DialogResult = true; dlg.Close(); };
        var root = new DockPanel();
        DockPanel.SetDock(btnOk, Dock.Bottom);
        root.Children.Add(btnOk);
        root.Children.Add(lb);
        dlg.Content = root;

        if (dlg.ShowDialog() == true && lb.SelectedItem is string addr)
        {
            DeliveryAddressText.Text = addr;
            RebuildPreview();
        }
    }

    private void BasisTextBox_OnTextChanged(object sender, TextChangedEventArgs e) => RebuildPreview();

    private void SupplierContactCombo_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateSupplierContactInfo();
        RememberSupplierContactChoice();
        RebuildPreview();
    }

    private void OpenInvoiceButton_OnClick(object sender, RoutedEventArgs e) => OpenInvoiceFile();

    private async void PdfButton_OnClick(object sender, RoutedEventArgs e)
    {
        await CreatePdfOnlyAsync();
    }

    private async void SendButton_OnClick(object sender, RoutedEventArgs e)
    {
        await SendRequestAsync();
    }

    // ---------- DB helpers ----------

    private OdbcConnection EnsureConn()
    {
        if (_conn == null)
        {
            var dbPath = _cfg.GetResolvedDbPath();
            if (!File.Exists(dbPath))
                throw new Exception($"Файл базы не найден: {dbPath}\nОткройте Настройки.");

            var cs = $"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};Dbq={dbPath};Pwd={_cfg.DbPassword};";
            _conn = new OdbcConnection(cs);
            _conn.Open();
        }
        return _conn;
    }

    private int GetSelectedYear()
    {
        if (YearCombo.SelectedItem is string s && int.TryParse(s, out var y))
            return y;
        return DateTime.Now.Year;
    }

    private void LoadOrders()
    {
        try
        {
            OrdersGrid.ItemsSource = null;
            _allOrders = null;
            _currentOrderId = null;
            _currentOrderHeader = null;
            // Не трогаем WebView2 здесь: предварительный текст устанавливается после инициализации,
            // а содержимое будет обновлено при выборе заказа.

            int y = GetSelectedYear();
            var start = new DateTime(y, 1, 1);
            var end = start.AddYears(1);

            var conn = EnsureConn();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT
    o.ID,
    o.Num,
    o.[Data] AS OrderDate,
    p.[Name] AS ProviderName,
    o.[State] AS State,
    IIF(o.[State]=1,'Создан',
        IIF(o.[State]=2,'Заказан',
            IIF(o.[State]=3,'Получен',
                IIF(o.[State]=4,'Архивный','')))) AS StateName
FROM ([Orders] AS o INNER JOIN [Providers] AS p ON p.ID = o.[Provider])
WHERE o.[Data] >= ? AND o.[Data] < ?
ORDER BY o.[Data] DESC, o.ID DESC";
            cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Date, Value = start });
            cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Date, Value = end });

            using var da = new OdbcDataAdapter((OdbcCommand)cmd);
            var dt = new DataTable();
            da.Fill(dt);

            if (!dt.Columns.Contains("RequestInfo"))
            {
                dt.Columns.Add("RequestInfo", typeof(string));
                dt.Columns.Add("RequestCreatedBy", typeof(string));
            }

            foreach (DataRow r in dt.Rows)
            {
                if (r["OrderDate"] is DateTime d)
                    r["OrderDate"] = d;

                int oid = SafeInt(r["ID"]);
                var info = _db.GetRequestInfo(oid);
                if (info != null)
                {
                    if (!string.IsNullOrWhiteSpace(info.CreatedByDisplayName))
                        r["RequestCreatedBy"] = info.CreatedByDisplayName;

                    if (info.SentLocal.HasValue)
                        r["RequestInfo"] = "Отправлена " + info.SentLocal.Value.ToString("dd.MM.yyyy HH:mm");
                    else if (info.CreatedLocal.HasValue)
                        r["RequestInfo"] = "Создана " + info.CreatedLocal.Value.ToString("dd.MM.yyyy HH:mm");
                }
            }

            _allOrders = dt;
            ApplyFilter();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка загрузки заказов:\n" + ex.Message,
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ApplyFilter()
    {
        if (_allOrders == null) return;
        var dv = new DataView(_allOrders);

        var filters = new List<string>();
        var q = (FilterBox.Text ?? string.Empty).Trim().Replace("'", "''");
        if (q.Length > 0)
        {
            filters.Add($"(Convert(Num, 'System.String') LIKE '%{q}%' OR Convert(ProviderName, 'System.String') LIKE '%{q}%')");
        }

        if (StatusCombo.SelectedItem is string status &&
            string.Equals(status, "Актуальные", StringComparison.OrdinalIgnoreCase))
        {
            filters.Add("[State] IN (1,2)");
        }

        dv.RowFilter = filters.Count > 0 ? string.Join(" AND ", filters) : string.Empty;
        OrdersGrid.ItemsSource = dv;
        if (OrdersGrid.Items.Count > 0)
            OrdersGrid.SelectedIndex = 0;
    }

    /// <summary>
    /// Асинхронно загружает и отображает выбранный заказ с предварительной инициализацией WebView2
    /// </summary>
    private async Task LoadSelectedOrderAsync()
    {
        if (OrdersGrid.SelectedItem is not DataRowView drv) return;
        if (!int.TryParse(drv.Row["ID"].ToString(), out var id)) return;

        try
        {
            var (order, items) = FetchOrderDetails(id);
            _currentOrderId = id;
            _currentProviderId = order.ProviderId;
            _currentOrderHeader = order;

            UpdateCustomerContactForOrder(order.ID);

            var basisTemplate = GetCustomerOrderNotesTemplate(id, null);
            BasisTextBox.Text = basisTemplate;

            LoadSupplierContacts(order.ProviderId);

            var (appNoPrev, _) = GetAppNoAndFilePath(order);
            var html = BuildHtml(order, items).Replace("{APP_NO}", appNoPrev);

            // КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Обеспечиваем инициализацию WebView2 перед использованием
            await EnsureWebViewAsync();
            
            // Проверяем, что WebView2 успешно инициализирован
            if (_isWebViewInitialized && PreviewBrowser.CoreWebView2 != null)
            {
                PreviewBrowser.NavigateToString(html);
            }
            else
            {
                throw new InvalidOperationException("Не удалось инициализировать WebView2 для предварительного просмотра");
            }

            UpdateInvoiceUiState();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка получения данных заказа:\n" + ex.Message,
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private (OrderHeader order, DataTable items) FetchOrderDetails(int orderId)
    {
        var conn = EnsureConn();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
SELECT
  o.ID,
  o.[Data]       AS OrderDate,
  o.Num          AS OrderNum,
  o.[Note]       AS OrderNote,
  o.InvoicePath  AS InvoicePath,
  p.ID           AS ProviderID,
  p.[Name]       AS ProviderName,
  p.FullName     AS ProviderFullName,
  p.INN,
  p.KPP,
  p.Address      AS ProviderAddress,
  p.Site, p.Email, p.Phone,
  p.Note         AS ProviderNote
FROM [Orders] AS o
INNER JOIN [Providers] AS p ON p.ID = o.[Provider]
WHERE o.ID = ?";
        cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Int, Value = orderId });
        using var rd = (OdbcDataReader)cmd.ExecuteReader();
        if (!rd.Read()) throw new Exception($"Заказ ID={orderId} не найден.");

        var oh = new OrderHeader
        {
            ID = orderId,
            OrderDate = rd["OrderDate"] is DateTime d1 ? d1 : (DateTime?)null,
            OrderNum = rd["OrderNum"]?.ToString() ?? string.Empty,
            Note = rd["OrderNote"]?.ToString() ?? string.Empty,
            InvoicePath = rd["InvoicePath"]?.ToString() ?? string.Empty,
            ProviderId = SafeInt(rd["ProviderID"]),
            ProviderName = rd["ProviderName"]?.ToString() ?? string.Empty,
            ProviderFullName = rd["ProviderFullName"]?.ToString() ?? string.Empty,
            INN = rd["INN"]?.ToString() ?? string.Empty,
            KPP = rd["KPP"]?.ToString() ?? string.Empty,
            ProviderAddress = rd["ProviderAddress"]?.ToString() ?? string.Empty,
            Site = rd["Site"]?.ToString() ?? string.Empty,
            Email = rd["Email"]?.ToString() ?? string.Empty,
            Phone = rd["Phone"]?.ToString() ?? string.Empty,
            ProviderNote = rd["ProviderNote"]?.ToString() ?? string.Empty,
            CustomerOrdersCsv = GetCustomerOrdersCsv(orderId, conn)
        };
        rd.Close();

        using var cmd2 = conn.CreateCommand();
        bool hasPN = ComponentHasPartNumber();
        var pnExpr = hasPN ? "co.PartNumber AS PartNumber" : "NULL AS PartNumber";
        cmd2.CommandText = @"
SELECT
  op.ID,
  co.[Name] AS ItemName,
  op.[Qty]  AS Qty,
  IIf(Not IsNull(uop.Symbol) And uop.Symbol<>'', uop.Symbol, uco.Symbol) AS UnitSymbol,
  IIf(Not IsNull(uop.[Name]) And uop.[Name]<>'', uop.[Name], uco.[Name]) AS UnitName,
  " + pnExpr + @",
  co.Code AS ItemCode
FROM ((([OrderPos] AS op
    LEFT JOIN [Component] AS co ON co.ID = op.CompID)
    LEFT JOIN [Unit] AS uop ON uop.ID = op.UnitID)
    LEFT JOIN [Unit] AS uco ON uco.ID = co.UnitID)
WHERE op.OrderID = ?
ORDER BY op.ID";
        cmd2.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Int, Value = orderId });
        using var da = new OdbcDataAdapter((OdbcCommand)cmd2);
        var items = new DataTable();
        da.Fill(items);

        return (oh, items);
    }

    private string GetCustomerOrdersCsv(int suppOrderId, OdbcConnection? existingConn)
    {
        var conn = existingConn ?? EnsureConn();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
SELECT cu.[Number]
FROM ([LinkOrder] AS lo
   INNER JOIN [CustomerOrder] AS cu ON cu.ID = lo.CustOrderID)
WHERE lo.SuppOrderID = ?
ORDER BY cu.[Number]";
        cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Int, Value = suppOrderId });

        var list = new List<string>();
        using var rd = (OdbcDataReader)cmd.ExecuteReader();
        while (rd.Read())
        {
            var s = rd[0]?.ToString();
            if (!string.IsNullOrWhiteSpace(s)) list.Add(s!);
        }
        return string.Join(", ", list);
    }

    private string GetCustomerOrderNotesTemplate(int suppOrderId, OdbcConnection? existingConn)
    {
        var conn = existingConn ?? EnsureConn();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
SELECT cu.[Note]
FROM ([LinkOrder] AS lo
   INNER JOIN [CustomerOrder] AS cu ON cu.ID = lo.CustOrderID)
WHERE lo.SuppOrderID = ?
ORDER BY cu.[Number]";
        cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Int, Value = suppOrderId });

        var list = new List<string>();
        using var rd = (OdbcDataReader)cmd.ExecuteReader();
        while (rd.Read())
        {
            var s = rd[0]?.ToString();
            if (!string.IsNullOrWhiteSpace(s)) list.Add(s!.Trim());
        }
        return string.Join(", ", list);
    }

    // --- контакт заказчика (создатель заявки) ---

    private void UpdateCustomerContactForOrder(int orderId)
    {
        var info = _db.GetRequestInfo(orderId);

        if (info != null && !string.IsNullOrWhiteSpace(info.CreatedByDisplayName))
        {
            _customerFioForCurrentOrder = info.CreatedByDisplayName ?? "";
            _customerEmailForCurrentOrder = info.CreatedByEmail ?? "";
            _customerPhoneForCurrentOrder = info.CreatedByPhone ?? "";
        }
        else
        {
            // заявки ещё не было — берём текущего пользователя
            _customerFioForCurrentOrder = _currentUser.ToContactDisplayString();
            _customerEmailForCurrentOrder = _currentUser.Email ?? "";
            _customerPhoneForCurrentOrder = _currentUser.Phone ?? "";
        }

        CustomerContactText.Text =
            $"{_customerFioForCurrentOrder}  {_customerPhoneForCurrentOrder}  {_customerEmailForCurrentOrder}";
    }

    private void LoadSupplierContacts(int providerId)
    {
        SupplierContactCombo.ItemsSource = null;
        SupplierContactInfoText.Text = string.Empty;

        try
        {
            var conn = EnsureConn();
            if (!TableExists(conn, "Contacts")) { UpdateInvoiceUiState(); return; }
            if (!TableHasColumn(conn, "Contacts", "OrgID")) { UpdateInvoiceUiState(); return; }

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT ID, [Name], [Email], [Phone], [Note] FROM [Contacts] WHERE [OrgID] = ? ORDER BY ID";
            cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Int, Value = providerId });

            using var da = new OdbcDataAdapter((OdbcCommand)cmd);
            var dt = new DataTable(); da.Fill(dt);

            var list = new List<ContactItem>();
            foreach (DataRow r in dt.Rows)
            {
                list.Add(new ContactItem
                {
                    Id = SafeInt(r["ID"]),
                    Name = N(r["Name"]),
                    Email = N(r["Email"]),
                    Phone = N(r["Phone"]),
                    Note = N(r["Note"])
                });
            }

            SupplierContactCombo.ItemsSource = list;
            if (list.Count > 0)
            {
                SupplierContactCombo.IsEnabled = true;

                if (_cfg.LastSupplierContact != null &&
                    _cfg.LastSupplierContact.TryGetValue(providerId, out var savedId))
                {
                    SupplierContactCombo.SelectedItem = list.FirstOrDefault(x => x.Id == savedId) ?? list[0];
                }
                else
                {
                    SupplierContactCombo.SelectedIndex = 0;
                    RememberSupplierContactChoice();
                }
            }
            else
            {
                SupplierContactCombo.IsEnabled = false;
            }

            UpdateSupplierContactInfo();
        }
        catch
        {
        }
        finally
        {
            UpdateInvoiceUiState();
        }
    }

    private void UpdateSupplierContactInfo()
    {
        if (SupplierContactCombo.SelectedItem is not ContactItem ci)
        {
            SupplierContactInfoText.Text = string.Empty;
            return;
        }

        var parts = new List<string> { ci.Name };
        if (!string.IsNullOrWhiteSpace(ci.Phone)) parts.Add("тел.: " + ci.Phone);
        if (!string.IsNullOrWhiteSpace(ci.Email)) parts.Add("e-mail: " + ci.Email);
        SupplierContactInfoText.Text = string.Join("; ", parts);

        if (!string.IsNullOrWhiteSpace(ci.Note))
            SupplierContactInfoText.Text += "\n" + ci.Note;
    }

    private void RememberSupplierContactChoice()
    {
        if (SupplierContactCombo.SelectedItem is ContactItem ci && _currentProviderId > 0)
        {
            _cfg.LastSupplierContact[_currentProviderId] = ci.Id;
            _cfg.Save(_db);
        }
    }

    private static bool TableExists(OdbcConnection conn, string table)
    {
        try
        {
            var t = conn.GetSchema("Tables");
            foreach (DataRow r in t.Rows)
            {
                var name = r["TABLE_NAME"]?.ToString();
                if (string.Equals(name, table, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch { }
        return false;
    }

    private static bool TableHasColumn(OdbcConnection conn, string table, string column)
    {
        try
        {
            var cols = conn.GetSchema("Columns", new string?[] { null, null, table, null });
            foreach (DataRow r in cols.Rows)
            {
                var col = r["COLUMN_NAME"]?.ToString();
                if (string.Equals(col, column, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch { }
        return false;
    }

    private bool ComponentHasPartNumber()
    {
        try
        {
            var conn = EnsureConn();
            var cols = conn.GetSchema("Columns", new string?[] { null, null, "Component", null });
            foreach (DataRow r in cols.Rows)
            {
                var col = r["COLUMN_NAME"]?.ToString();
                if (string.Equals(col, "PartNumber", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col, "PartNum", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col, "PartNo", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(col, "PN", StringComparison.OrdinalIgnoreCase))
                    return true;
            }
        }
        catch { }
        return false;
    }

    // ---------- Helpers ----------

    private static string N(object? v) => v?.ToString() ?? string.Empty;
    private static int SafeInt(object? v) { try { return Convert.ToInt32(v); } catch { return 0; } }

    private static string HtmlEncode(string s) => System.Net.WebUtility.HtmlEncode(s ?? string.Empty);

    private static string FmtDate(DateTime? d) => d.HasValue ? d.Value.ToString("dd.MM.yy") : string.Empty;

    private static string FormatQty(object? v)
    {
        if (v == null || v is DBNull) return string.Empty;

        var ru = CultureInfo.GetCultureInfo("ru-RU");

        if (v is decimal dec)
            return FormatDecimal(dec, ru);
        if (v is double dbl && !double.IsNaN(dbl) && !double.IsInfinity(dbl))
            return FormatDecimal((decimal)dbl, ru);
        if (v is float flt && !float.IsNaN(flt) && !float.IsInfinity(flt))
            return FormatDecimal((decimal)flt, ru);
        if (v is int i) return i.ToString(ru);
        if (v is long l) return l.ToString(ru);

        var str = v.ToString()?.Trim();
        if (string.IsNullOrEmpty(str)) return string.Empty;

        if (decimal.TryParse(str, NumberStyles.Any, ru, out dec))
            return FormatDecimal(dec, ru);

        if (decimal.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, out dec))
            return FormatDecimal(dec, ru);

        return str;
    }

    private static string FormatDecimal(decimal value, CultureInfo culture)
    {
        string s = value.ToString("0.###############", culture);
        int decimalSeparatorIndex = s.IndexOf(culture.NumberFormat.NumberDecimalSeparator, StringComparison.Ordinal);
        if (decimalSeparatorIndex >= 0)
        {
            int lastNonZero = -1;
            for (int i = s.Length - 1; i > decimalSeparatorIndex; i--)
            {
                if (s[i] != '0')
                {
                    lastNonZero = i;
                    break;
                }
            }

            if (lastNonZero >= 0)
                s = s[..(lastNonZero + 1)];
            else
                s = s[..decimalSeparatorIndex];
        }

        return s;
    }

    private void LoadAddresses()
    {
        var list = _cfg.DeliveryAddresses ?? AppSettings.Defaults.DeliveryBook;
        DeliveryAddressText.Text = (list.Length > 0) ? list[0] : "Адрес доставки не задан";
    }

    private void ApplyThemeToPreview()
    {
        // WebView2 может быть ещё не инициализирован / инициализация могла упасть.
        // ВАЖНО: сам доступ к свойству CoreWebView2 до инициализации тоже бросает исключение,
        // поэтому здесь проверяем только наш флаг и сам контрол, без обращения к CoreWebView2.
        if (!_isWebViewInitialized || PreviewBrowser == null)
        {
            if (PreviewErrorText != null)
                PreviewErrorText.Visibility = Visibility.Visible;
            return;
        }

        var (bg, fg, _, _) = ThemeCssVars();
        var html = $"<!doctype html><meta charset='utf-8'><body style='margin:24px;font-family:Segoe UI;background:{bg};color:{fg};opacity:.7'>Выберите заказ слева…</body>";
        PreviewBrowser.NavigateToString(html);
    }

    private (string Bg, string Fg, string Border, string ThBg) ThemeCssVars()
    {
        // Определяем эффективную тему (учитываем Auto режим)
        bool isDark = _cfg.ThemeMode == UiThemeMode.Dark ||
                     (_cfg.ThemeMode == UiThemeMode.Auto && ThemeManager.IsSystemDarkTheme());
        
        if (isDark)
            return ("#1f1f1f", "#e6e6e6", "#555", "#2a2a2a");
        return ("#ffffff", "#111", "#999", "#f2f2f2");
    }

    private void RebuildPreview()
    {
        if (_currentOrderId == null) return;
        try
        {
            UpdateCustomerContactForOrder(_currentOrderId.Value);

            var (order, items) = FetchOrderDetails(_currentOrderId.Value);
            _currentOrderHeader = order;
            var (appNoPrev, _) = GetAppNoAndFilePath(order);
            var html = BuildHtml(order, items).Replace("{APP_NO}", appNoPrev);

            if (PreviewBrowser != null && PreviewBrowser.CoreWebView2 != null)
            {
                PreviewBrowser.NavigateToString(html);
            }
        
        
            UpdateInvoiceUiState();
        }
        catch
        {
        }
    }



        private static string TwoDigitYear(DateTime? d) => (d ?? DateTime.Now).ToString("yy");

    private string ComputeNextAppNo(DateTime? date)
    {
        var yy = TwoDigitYear(date);
        var dir = _cfg.GetResolvedOutputDir();
        if (string.IsNullOrWhiteSpace(dir)) dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "out");
        Directory.CreateDirectory(dir);

        int max = 0;
        var re = new Regex(@"\b(\d{3})-(\d{2})\b");
        foreach (var f in Directory.EnumerateFiles(dir, "*Заявка на закупку*.*", SearchOption.TopDirectoryOnly))
        {
            var m = re.Match(Path.GetFileNameWithoutExtension(f));
            if (m.Success && m.Groups[2].Value == yy && int.TryParse(m.Groups[1].Value, out var n) && n > max) max = n;
        }
        return $"{(max + 1):D3}-{yy}";
    }

    private static string SanitizeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new StringBuilder(name.Length);
        foreach (var ch in name) sb.Append(invalid.Contains(ch) ? '_' : ch);
        return sb.ToString();
    }

    private string ComposeDefaultFileName(OrderHeader order, string appNo)
    {
        var prefix = string.IsNullOrWhiteSpace(order.OrderNum) ? string.Empty : (order.OrderNum.Trim() + " ");
        var dt = (order.OrderDate ?? DateTime.Now);
        return SanitizeFileName($"{prefix}Заявка на закупку {appNo} {dt:dd.MM.yy}.pdf");
    }

    private (string appNo, string fullPath) GetAppNoAndFilePath(OrderHeader order)
    {
        var outDir = _cfg.GetResolvedOutputDir();
        if (string.IsNullOrWhiteSpace(outDir))
            outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "out");
        Directory.CreateDirectory(outDir);

        string appNo;
        string fullPath;

        if (!string.IsNullOrWhiteSpace(order.OrderNum))
        {
            var safeOrderNum = SanitizeFileName(order.OrderNum.Trim());
            var pattern = safeOrderNum + " Заявка на закупку*.pdf";

            var candidates = Directory
                .EnumerateFiles(outDir, pattern, SearchOption.TopDirectoryOnly)
                .OrderByDescending(f => File.GetLastWriteTime(f))
                .ToList();

            if (candidates.Count > 0)
            {
                fullPath = candidates[0];
                var m = Regex.Match(Path.GetFileNameWithoutExtension(fullPath), @"\b(\d{3}-\d{2})\b");
                if (m.Success)
                    appNo = m.Groups[1].Value;
                else
                    appNo = ComputeNextAppNo(order.OrderDate);

                return (appNo, fullPath);
            }
        }

        appNo = ComputeNextAppNo(order.OrderDate);
        var fileName = ComposeDefaultFileName(order, appNo);
        fullPath = Path.Combine(outDir, fileName);
        return (appNo, fullPath);
    }

    private void TryOpenFileWithShell(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                MessageBox.Show(this, "Файл не найден:\n" + path, "Открытие файла",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo(path)
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Ошибка при открытии файла:\n" + ex.Message,
                "Открытие файла", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private string? ResolveInvoicePath(string? rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath)) return null;
        var p = Environment.ExpandEnvironmentVariables(rawPath.Trim());

        if (!Path.IsPathRooted(p))
        {
            try
            {
                var dbPath = _cfg.GetResolvedDbPath();
                var dbDir = string.IsNullOrWhiteSpace(dbPath)
                    ? AppDomain.CurrentDomain.BaseDirectory
                    : Path.GetDirectoryName(dbPath);
                if (!string.IsNullOrWhiteSpace(dbDir))
                    p = Path.GetFullPath(Path.Combine(dbDir!, p));
            }
            catch { }
        }

        return p;
    }

    private void OpenInvoiceFile()
    {
        if (_currentOrderHeader == null || string.IsNullOrWhiteSpace(_currentOrderHeader.InvoicePath))
        {
            MessageBox.Show(this, "Для выбранной заявки путь к счёту не указан.",
                "Счёт", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var path = ResolveInvoicePath(_currentOrderHeader.InvoicePath);
        if (string.IsNullOrWhiteSpace(path))
        {
            MessageBox.Show(this, "Не удалось определить путь к счёту.", "Счёт",
                MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        TryOpenFileWithShell(path);
    }

    private void UpdateInvoiceUiState()
    {
        bool hasInvoice = _currentOrderHeader != null &&
                          !string.IsNullOrWhiteSpace(_currentOrderHeader.InvoicePath);
        OpenInvoiceButton.IsEnabled = hasInvoice;
        AttachInvoiceCheckBox.IsEnabled = hasInvoice;
    }

    private bool _isWebViewInitialized = false;

    private async Task EnsureWebViewAsync()
    {
        if (_isWebViewInitialized) return;

        try
        {
            PreviewBrowser.CoreWebView2InitializationCompleted += OnWebViewInitialized;
            await PreviewBrowser.EnsureCoreWebView2Async();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "WebView2 Runtime не найден.\n" + ex.Message,
                "Ошибка WebView2", MessageBoxButton.OK, MessageBoxImage.Error);
            _isWebViewInitialized = false;
        }
    }

    private void OnWebViewInitialized(object? sender, CoreWebView2InitializationCompletedEventArgs e)
    {
        if (!e.IsSuccess)
        {
            MessageBox.Show(this, "Ошибка инициализации WebView2: " + e.InitializationException?.Message,
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        _isWebViewInitialized = true;
        ApplyThemeToPreview();
    }

    private async Task GeneratePdfToFileAsync(string filePath, string html)
    {
        await EnsureWebViewAsync();
        if (!_isWebViewInitialized || PreviewBrowser.CoreWebView2 == null)
        {
            throw new InvalidOperationException("WebView2 не инициализирован");
        }

        var tcs = new TaskCompletionSource<bool>();
        void OnNavCompleted(object? s, CoreWebView2NavigationCompletedEventArgs e)
        {
            tcs.TrySetResult(true);
            PreviewBrowser.CoreWebView2.NavigationCompleted -= OnNavCompleted;
        }

        PreviewBrowser.CoreWebView2.NavigationCompleted += OnNavCompleted;
        PreviewBrowser.NavigateToString(html);
        await tcs.Task;

        await PreviewBrowser.CoreWebView2.PrintToPdfAsync(filePath);
    }

    // ---------- Mail ----------

    private static IEnumerable<string> SplitEmails(string emails)
    {
        if (string.IsNullOrWhiteSpace(emails)) yield break;
        var parts = emails.Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (var p in parts)
        {
            var s = p.Trim();
            if (s.Length > 0) yield return s;
        }
    }

    private async Task SendEmailAsync(string toEmails, string subject, string bodyText, IReadOnlyList<string> attachmentPaths)
    {
        // From: сначала берём из карточки сотрудника, fallback — настройки
        var fromEmail = !string.IsNullOrWhiteSpace(_currentUser.Email)
            ? _currentUser.Email
            : _cfg.FromEmail;

        if (string.IsNullOrWhiteSpace(fromEmail))
            throw new InvalidOperationException("Не настроен e-mail отправителя (карточка сотрудника или настройки).");

        if (string.IsNullOrWhiteSpace(_cfg.SmtpHost))
            throw new InvalidOperationException("Не настроен SMTP сервер в настройках.");

        var msg = new MimeMessage();

        var fromName = !string.IsNullOrWhiteSpace(_currentUser.PersonName)
            ? _currentUser.PersonName
            : (!string.IsNullOrWhiteSpace(_cfg.FromDisplayName) ? _cfg.FromDisplayName : fromEmail);

        msg.From.Add(new MailboxAddress(fromName, fromEmail));

        foreach (var addr in SplitEmails(toEmails))
            msg.To.Add(MailboxAddress.Parse(addr));

        msg.Subject = subject ?? string.Empty;

        var bodyBuilder = new BodyBuilder
        {
            TextBody = bodyText ?? string.Empty
        };

        foreach (var path in attachmentPaths ?? Array.Empty<string>())
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    bodyBuilder.Attachments.Add(path);
            }
            catch
            {
            }
        }

        msg.Body = bodyBuilder.ToMessageBody();

        using (var smtp = new SmtpClient())
        {
            smtp.Timeout = 120000; // 120 секунд

            var secure = _cfg.SmtpUseSsl ? SecureSocketOptions.StartTlsWhenAvailable : SecureSocketOptions.None;
            await smtp.ConnectAsync(_cfg.SmtpHost, _cfg.SmtpPort, secure);
            if (!string.IsNullOrWhiteSpace(_cfg.SmtpLogin))
                await smtp.AuthenticateAsync(_cfg.SmtpLogin, _cfg.SmtpPassword);
            await smtp.SendAsync(msg);
            await smtp.DisconnectAsync(true);
        }

        await AppendToSentAsync(msg);
    }

    private async Task AppendToSentAsync(MimeMessage msg)
    {
        if (string.IsNullOrWhiteSpace(_cfg.ImapHost))
            return;

        try
        {
            using var imap = new ImapClient();
            await imap.ConnectAsync(_cfg.ImapHost, _cfg.ImapPort, _cfg.ImapUseSsl);

            // логин/пароль — те же, что для SMTP
            if (!string.IsNullOrWhiteSpace(_cfg.SmtpLogin))
                await imap.AuthenticateAsync(_cfg.SmtpLogin, _cfg.SmtpPassword);

            var folderName = string.IsNullOrWhiteSpace(_cfg.ImapSentFolder) ? "Sent" : _cfg.ImapSentFolder;
            var folder = await imap.GetFolderAsync(folderName);
            if (folder != null)
            {
                await folder.OpenAsync(MailKit.FolderAccess.ReadWrite);
                await folder.AppendAsync(msg, MailKit.MessageFlags.Seen);
                await folder.CloseAsync();
            }

            await imap.DisconnectAsync(true);
        }
        catch
        {
        }
    }

    private async Task TrySendNotifyOnCreateAsync(OrderHeader order, string appNo)
    {
        var to = !string.IsNullOrWhiteSpace(_cfg.NotifyOnCreateTo)
            ? _cfg.NotifyOnCreateTo
            : _cfg.AdminNotifyEmail;

        if (string.IsNullOrWhiteSpace(to)) return;

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Создана новая заявка на закупку.");
            sb.AppendLine();
            sb.AppendLine($"Заявка № {appNo}");
            sb.AppendLine($"Поставщик: {order.ProviderName}");
            sb.AppendLine($"Инициатор: {_currentUser.PersonName}");
            sb.AppendLine($"Дата: {FmtDate(order.OrderDate)}");

            await SendEmailAsync(to,
                $"[Заявка создана] № {appNo} ({order.ProviderName})",
                sb.ToString(),
                Array.Empty<string>());
        }
        catch
        {
        }
    }

    private async Task TrySendNotifyOnSendAsync(OrderHeader order, string appNo)
    {
        // уведомление — создателю заявки
        var info = _db.GetRequestInfo(order.ID);
        var to = info?.CreatedByEmail;
        if (string.IsNullOrWhiteSpace(to)) return;

        try
        {
            var sb = new StringBuilder();
            sb.AppendLine("Ваша заявка на закупку проверена и отправлена в отдел.");
            sb.AppendLine();
            sb.AppendLine($"Заявка № {appNo}");
            sb.AppendLine($"Поставщик: {order.ProviderName}");
            sb.AppendLine($"Проверил и отправил: {_currentUser.PersonName}");
            sb.AppendLine($"Дата отправки: {DateTime.Now:dd.MM.yyyy HH:mm}");

            await SendEmailAsync(to,
                $"[Заявка отправлена] № {appNo} ({order.ProviderName})",
                sb.ToString(),
                Array.Empty<string>());
        }
        catch
        {
        }
    }

    // ---------- Create PDF / Send ----------

    private async Task CreatePdfOnlyAsync()
    {
        if (_currentOrderId == null)
        {
            MessageBox.Show(this, "Сначала выберите заказ.", "Нет данных",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            ToggleButtons(false);

            var (order, items) = FetchOrderDetails(_currentOrderId.Value);
            var (appNo, suggestedFullPath) = GetAppNoAndFilePath(order);
            var html = BuildHtml(order, items).Replace("{APP_NO}", appNo);

            var outDir = Path.GetDirectoryName(suggestedFullPath)!;
            Directory.CreateDirectory(outDir);

            var sfd = new SaveFileDialog
            {
                Filter = "PDF (*.pdf)|*.pdf",
                FileName = Path.GetFileName(suggestedFullPath),
                InitialDirectory = outDir
            };
            if (sfd.ShowDialog(this) != true)
                return;

            await GeneratePdfToFileAsync(sfd.FileName, html);
            _db.MarkRequestCreated(order.ID, _currentUser, sfd.FileName);

            if (OpenAfterExportCheckBox.IsChecked == true)
                TryOpenFileWithShell(sfd.FileName);

            await TrySendNotifyOnCreateAsync(order, appNo);

            LoadOrders();

            MessageBox.Show(this, "PDF сформирован:\n" + sfd.FileName,
                "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                "Ошибка при формировании PDF:\n" + ex.Message,
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ToggleButtons(true);
        }
    }

    private async Task SendRequestAsync()
    {
        if (_currentOrderId == null)
        {
            MessageBox.Show(this, "Сначала выберите заказ.", "Нет данных",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }
        if (!IsCurrentUserAdmin)
        {
            MessageBox.Show(this, "Отправка заявки доступна только пользователю с ролью администратора.",
                "Нет прав", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(_cfg.RequestToEmail))
        {
            MessageBox.Show(this,
                "В настройках не указан e-mail адресата заявки (RequestToEmail).\nОткройте настройки и заполните поле.",
                "Отправка письма", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            ToggleButtons(false);

            var (order, items) = FetchOrderDetails(_currentOrderId.Value);
            var (appNo, fullPath) = GetAppNoAndFilePath(order);
            var html = BuildHtml(order, items).Replace("{APP_NO}", appNo);

            var outDirSend = Path.GetDirectoryName(fullPath)!;
            Directory.CreateDirectory(outDirSend);

            await GeneratePdfToFileAsync(fullPath, html);
            _db.MarkRequestCreated(order.ID, _currentUser, fullPath);

            if (OpenAfterExportCheckBox.IsChecked == true)
                TryOpenFileWithShell(fullPath);

            string? invoicePathToAttach = null;
            if (AttachInvoiceCheckBox.IsChecked == true &&
                !string.IsNullOrWhiteSpace(order.InvoicePath))
            {
                var resolved = ResolveInvoicePath(order.InvoicePath);
                if (!string.IsNullOrWhiteSpace(resolved) && File.Exists(resolved))
                {
                    invoicePathToAttach = resolved;
                }
                else
                {
                    MessageBox.Show(this,
                        "Путь к счёту указан, но файл не найден:\n" + resolved,
                        "Прикрепление счёта", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }

            var attachments = new List<string> { fullPath };
            if (!string.IsNullOrWhiteSpace(invoicePathToAttach))
                attachments.Add(invoicePathToAttach);

            string toEmails = _cfg.RequestToEmail;
            string subject = $"Заявка на закупку № {appNo} ({order.ProviderName})";

            var body = new StringBuilder();
            body.AppendLine("Здравствуйте,");
            body.AppendLine();
            body.AppendLine("Направляем вам заявку на закупку во вложении.");
            body.AppendLine();

            if (!string.IsNullOrWhiteSpace(BasisTextBox.Text))
            {
                body.AppendLine("Основание: " + BasisTextBox.Text);
                body.AppendLine();
            }

            body.AppendLine("С уважением,");
            body.AppendLine(_customerFioForCurrentOrder);
            if (!string.IsNullOrWhiteSpace(_cfg.EmailSignature))
                body.AppendLine(_cfg.EmailSignature);

            string bodyText = body.ToString();

            var preview = new EmailPreviewWindow(toEmails, subject, bodyText, attachments);
            if (preview.ShowDialog() != true)
            {
                MessageBox.Show(this, "Отправка отменена пользователем.",
                    "Отправка", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await SendEmailAsync(preview.ToText, preview.SubjectText, preview.BodyText, attachments);

            _db.MarkRequestSent(order.ID, _currentUser, fullPath, preview.ToText);

            await TrySendNotifyOnSendAsync(order, appNo);

            await UpdateOrderStateToOrderedAsync(order.ID);

            LoadOrders();

            MessageBox.Show(this, "Письмо с заявкой отправлено.", "Готово",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this,
                "Ошибка при формировании PDF и/или отправке письма:\n" + ex.Message,
                "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            ToggleButtons(true);
        }
    }

    private async Task UpdateOrderStateToOrderedAsync(int orderId)
    {
        try
        {
            var conn = EnsureConn();
            using var cmd = conn.CreateCommand();
            cmd.CommandText = "UPDATE Orders SET [State] = 2 WHERE ID = ?";
            cmd.Parameters.Add(new OdbcParameter { OdbcType = OdbcType.Int, Value = orderId });
            await Task.Run(() => cmd.ExecuteNonQuery());
        }
        catch
        {
        }
    }

    private void ToggleButtons(bool enabled)
    {
        PdfButton.IsEnabled = enabled;
        SendButton.IsEnabled = enabled && IsCurrentUserAdmin;
    }

    // ---------- HTML ----------

    private string BuildHtml(OrderHeader order, DataTable items)
    {
        var sbRows = new StringBuilder();
        int i = 0;
        foreach (DataRow r in items.Rows)
        {
            i++;
            var itemName = HtmlEncode(N(r["ItemName"]));
            var qty = HtmlEncode(FormatQty(r["Qty"]));
            var unit = !string.IsNullOrWhiteSpace(N(r["UnitSymbol"]))
                           ? HtmlEncode(N(r["UnitSymbol"]))
                           : HtmlEncode(N(r["UnitName"]));
            var pn = HtmlEncode(N(r["PartNumber"]));
            var itemCode = HtmlEncode(N(r["ItemCode"]));

            sbRows.AppendLine("<tr>");
            sbRows.AppendLine($"  <td class='num'>{i}</td>");
            sbRows.AppendLine($"  <td class='code-nom'>{itemCode}</td>");
            sbRows.AppendLine($"  <td class='name'>{itemName}</td>");
            sbRows.AppendLine($"  <td class='code-supp'>{pn}</td>");
            sbRows.AppendLine($"  <td class='qty'>{qty}</td>");
            sbRows.AppendLine($"  <td class='unit'>{unit}</td>");
            sbRows.AppendLine("</tr>");
        }

        var sc = SupplierContactCombo.SelectedItem as ContactItem;
        var delivery = DeliveryAddressText.Text ?? "";
        var basis = BasisTextBox.Text ?? "";

        var suppNo = (order.OrderNum ?? "").Trim();
        var custList = (order.CustomerOrdersCsv ?? "").Trim();
        var suppAndCust = string.IsNullOrWhiteSpace(custList)
            ? suppNo
            : $"{suppNo} / {custList}";

        var (bg, fg, border, thBg) = ThemeCssVars();
        string themeVars = $"--paper-bg:{bg};--paper-fg:{fg};--b:{border};--thbg:{thBg};";

        // Контакт поставщика — человекочитаемо
        string suplContactText = "";
        if (sc != null)
        {
            var parts = new List<string> { sc.Name };
            if (!string.IsNullOrWhiteSpace(sc.Phone))
                parts.Add("тел.: " + sc.Phone);
            if (!string.IsNullOrWhiteSpace(sc.Email))
                parts.Add("e-mail: " + sc.Email);
            suplContactText = string.Join("; ", parts);
        }

        string suplContactNoteText = sc?.Note ?? "";

        string requestToRaw = _cfg.RequestToName ?? string.Empty;
        string requestToEncoded = HtmlEncode(requestToRaw)
            .Replace("\r\n", "<br/>")
            .Replace("\n", "<br/>");

        return LoadTemplate()
            .Replace("{THEME_VARS}", themeVars)
            .Replace("{ORDER_DATE}", FmtDate(order.OrderDate))

            .Replace("{COMPANY_NAME}", HtmlEncode(_cfg.CompanyName ?? ""))
            .Replace("{CUSTOMER_FIO}", HtmlEncode(_customerFioForCurrentOrder))
            .Replace("{CUSTOMER_PHONE}", HtmlEncode(_customerPhoneForCurrentOrder))
            .Replace("{CUSTOMER_EMAIL}", HtmlEncode(_customerEmailForCurrentOrder))
            .Replace("{DELIVERY_ADDRESS}", HtmlEncode(delivery))
            .Replace("{SUPP_AND_CUST}", HtmlEncode(suppAndCust))
            .Replace("{ORDER_NOTE}", HtmlEncode(order.Note ?? ""))
            .Replace("{BASIS}", HtmlEncode(basis))

            .Replace("{PROVIDER_NAME}", HtmlEncode(
                string.IsNullOrWhiteSpace(order.ProviderFullName)
                    ? order.ProviderName
                    : order.ProviderFullName))
            .Replace("{PROVIDER_INN}", HtmlEncode(order.INN))
            .Replace("{PROVIDER_ADDR}", HtmlEncode(order.ProviderAddress))
            .Replace("{PROVIDER_SITE}", HtmlEncode(order.Site))
            .Replace("{PROVIDER_EMAIL}", HtmlEncode(order.Email))
            .Replace("{PROVIDER_PHONE}", HtmlEncode(order.Phone))
            .Replace("{PROVIDER_NOTE}", HtmlEncode(order.ProviderNote))

            .Replace("{SUPL_CONTACT}", HtmlEncode(suplContactText))
            .Replace("{SUPL_CONTACT_NOTE}", HtmlEncode(suplContactNoteText))

            .Replace("{REQUEST_TO}", requestToEncoded)
            .Replace("{REQUEST_TO_EMAIL}", HtmlEncode(_cfg.RequestToEmail ?? ""))

            .Replace("{ROWS}", sbRows.ToString());
    }

    private string LoadTemplate()
    {
        return @"<!doctype html>
<html lang='ru'>
<head>
<meta charset='utf-8'>
<title>Заявка на закупку № {APP_NO}</title>
<style>
:root{
  --font:'Segoe UI',Arial,sans-serif;
  --fs:12pt; --fs-h1:18pt; --fs-to:13pt; --line:1.28; --gap:10pt;
  --left-col:220px; --pad:6px;
  {THEME_VARS}
}
*{box-sizing:border-box}
body{font-family:var(--font);font-size:var(--fs);line-height:var(--line);color:var(--paper-fg);background:var(--paper-bg);margin:12mm;}
h1{font-size:var(--fs-h1);text-align:center;margin:0 0 4pt 0}
.meta{margin:0 0 var(--gap) 0;text-align:center;opacity:.8}
.section-title{margin-top:var(--gap);font-weight:700;text-transform:uppercase}
.kv{display:grid;grid-template-columns:var(--left-col) 1fr;column-gap:12pt;row-gap:2pt;margin-top:4pt}
.kv .k{font-weight:600}
table{border-collapse:collapse;width:100%;margin-top:8pt;table-layout:fixed}
th,td{border:1px solid var(--b);padding:var(--pad);vertical-align:top}
th{background:var(--thbg);text-align:left}
.num{text-align:center;width:30px}
.code-nom{width:90px}
.code-supp{width:140px}
.qty{text-align:right;width:60px}
.unit{text-align:center;width:40px}
.small{font-size:10pt;opacity:.7;margin-top:10pt}
.header-grid{display:grid;grid-template-columns:1fr auto;column-gap:20pt;margin-bottom:6pt;}
.header-to{text-align:right;font-size:var(--fs-to);}
.header-to .to-label{font-weight:600;text-transform:uppercase;margin-bottom:2pt;}
.header-to .to-name{font-weight:600;}
.header-to .to-email{font-size:10pt;opacity:.8;}
</style>
</head>
<body>
  <div class='header-grid'>
    <div></div>
    <div class='header-to'>
      
      <div class='to-name'>{REQUEST_TO}</div>
      <div class='to-email'>{REQUEST_TO_EMAIL}</div>
    </div>
  </div>

  <h1>ЗАЯВКА НА ЗАКУПКУ № {APP_NO}</h1>
  <div class='meta'>Дата: {ORDER_DATE}</div>

  <div class='section-title'>Заказчик</div>
  <div class='kv'>
    <div class='k'>Подразделение:</div><div>{COMPANY_NAME}</div>
    <div class='k'>Контакт:</div><div>{CUSTOMER_FIO}; тел.: {CUSTOMER_PHONE}; e-mail: {CUSTOMER_EMAIL}</div>
    <div class='k'>Адрес доставки:</div><div>{DELIVERY_ADDRESS}</div>
    <div class='k'>Вн. № заказа Пост./Кл.:</div><div>{SUPP_AND_CUST}</div>
    <div class='k'>Примечание:</div><div>{ORDER_NOTE}</div>
    <div class='k'>Основание:</div><div>{BASIS}</div>
  </div>

  <div class='section-title'>Поставщик</div>
  <div class='kv'>
    <div class='k'>Наименование:</div><div>{PROVIDER_NAME}</div>
    <div class='k'>ИНН:</div><div>{PROVIDER_INN}</div>
    <div class='k'>Юр. адрес:</div><div>{PROVIDER_ADDR}</div>
    <div class='k'>Сайт:</div><div>{PROVIDER_SITE}</div>
    <div class='k'>E-mail:</div><div>{PROVIDER_EMAIL}</div>
    <div class='k'>Тел.:</div><div>{PROVIDER_PHONE}</div>
    <div class='k'>Примечание поставщика:</div><div>{PROVIDER_NOTE}</div>
    <div class='k'>Контакт поставщика:</div><div>{SUPL_CONTACT}</div>
    <div class='k'>Заметка контакта:</div><div>{SUPL_CONTACT_NOTE}</div>
  </div>

  <table>
    <thead><tr>
      <th class='num'>№</th>
      <th class='code-nom'>Ном. №</th>
      <th class='name'>Наименование</th>
      <th class='code-supp'>Код поставщика</th>
      <th class='qty'>Кол.</th>
      <th class='unit'>Ед.</th>
    </tr></thead>
    <tbody>{ROWS}</tbody>
  </table>

  <div class='small'>Документ сформирован автоматически.</div>
</body>
</html>";
        }
    }
}
