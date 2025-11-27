using System;
using System.Globalization;
using System.IO;
using Microsoft.Data.Sqlite;

namespace SupplierOrderPdf.Core;

/// <summary>
/// ��������� ���� (SQLite): ��������� + ������ ������.
/// </summary>
public sealed class AppDatabase : System.IDisposable
{
    private readonly string _path;
    private readonly SqliteConnection _conn;

    public SqliteConnection Connection => _conn;

    public AppDatabase(string path)
    {
        _path = path ?? throw new ArgumentNullException(nameof(path));

        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrWhiteSpace(dir))
            Directory.CreateDirectory(dir);

        _conn = new SqliteConnection($"Data Source={_path}");
        _conn.Open();
    }

    /// <summary>
    /// Инициализирует структуру БД:
    /// - включает журналирование WAL для лучшей устойчивости к сбоям;
    /// - создаёт таблицу Settings (если её ещё нет);
    /// - создаёт таблицу RequestLog (если её ещё нет);
    /// - проверяет наличие всех нужных колонок в RequestLog (на случай старых версий БД).
    /// Вызывается один раз при старте приложения.
    /// </summary>
    public void Initialize()
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = @"
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS Settings (
    Key   TEXT PRIMARY KEY,
    Value TEXT
);

CREATE TABLE IF NOT EXISTS RequestLog (
    OrderId              INTEGER PRIMARY KEY,
    CreatedUtc           TEXT,
    SentUtc              TEXT,
    CreatedByLogin       TEXT,
    CreatedByDisplayName TEXT,
    CreatedByEmail       TEXT,
    CreatedByPhone       TEXT,
    SentByLogin          TEXT,
    SentByDisplayName    TEXT,
    PdfPath              TEXT,
    LastEmailTo          TEXT
);";
        cmd.ExecuteNonQuery();

        // На случай старых БД, созданных до добавления новых полей,
        // гарантируем наличие всех нужных колонок в RequestLog.
        EnsureRequestLogColumns();
    }

    /// <summary>
    /// Гарантирует, что в таблице RequestLog есть все необходимые текстовые колонки,
    /// которые используются кодом приложения. Это позволяет «мигрировать» старые БД,
    /// созданные до добавления новых полей, без отдельного SQL-скрипта миграций.
    /// </summary>
    private void EnsureRequestLogColumns()
    {
        // Все текстовые колонки, которые мы читаем/пишем в RequestLog.
        string[] neededColumns =
        {
            "CreatedByLogin",
            "CreatedByDisplayName",
            "CreatedByEmail",
            "CreatedByPhone",
            "SentByLogin",
            "SentByDisplayName",
            "PdfPath",
            "LastEmailTo"
        };

        foreach (var column in neededColumns)
        {
            // Если нужной колонки нет — добавляем её через ALTER TABLE.
            if (!ColumnExists("RequestLog", column))
            {
                using var cmd = _conn.CreateCommand();
                cmd.CommandText = $"ALTER TABLE RequestLog ADD COLUMN {column} TEXT;";
                cmd.ExecuteNonQuery();
            }
        }
    }

    /// <summary>
    /// Проверяет, существует ли колонка <paramref name="column"/> в таблице <paramref name="table"/>.
    /// Используется через PRAGMA table_info(...) — это безопасный и быстрый способ introspection в SQLite.
    /// </summary>
    private bool ColumnExists(string table, string column)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = "PRAGMA table_info(" + table + ");";
        using var rd = cmd.ExecuteReader();
        while (rd.Read())
        {
            var name = rd["name"] as string;
            if (string.Equals(name, column, StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    // ---------- Settings ----------

    /// <summary>
    /// Возвращает строковое значение настройки по ключу или <c>null</c>,
    /// если такой записи нет в таблице Settings.
    /// Никаких преобразований типов здесь не выполняется.
    /// </summary>
    public string? GetSetting(string key)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = "SELECT Value FROM Settings WHERE Key = $key";
        cmd.Parameters.AddWithValue("$key", key);
        return cmd.ExecuteScalar() as string;
    }

    public string GetSetting(string key, string defaultValue)
    {
        var value = GetSetting(key);
        return value ?? defaultValue;
    }

    public void SetSetting(string key, string? value)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO Settings(Key, Value) VALUES($key,$value)
ON CONFLICT(Key) DO UPDATE SET Value = excluded.Value;";
        cmd.Parameters.AddWithValue("$key", key);
        cmd.Parameters.AddWithValue("$value", (object?)value ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    // ---------- RequestLog ----------

    /// <summary>
    /// Возвращает информацию о заявке по её идентификатору заказа из журнала RequestLog.
    /// Если запись ещё не создавалась (заявка не формировалась/не отправлялась) — возвращает <c>null</c>.
    /// Даты создаются в виде UTC-значений и дополнительно конвертируются в локальное время
    /// через вычисляемые свойства <see cref="RequestLogInfo.CreatedLocal"/> и
    /// <see cref="RequestLogInfo.SentLocal"/>.
    /// </summary>
    public RequestLogInfo? GetRequestInfo(int orderId)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM RequestLog WHERE OrderId = $id";
        cmd.Parameters.AddWithValue("$id", orderId);

        using var rd = cmd.ExecuteReader();
        if (!rd.Read()) return null;

        // Локальная функция парсинга даты/времени в UTC-формате (ISO 8601).
        DateTime? ParseUtc(string col)
        {
            var s = rd[col] as string;
            if (string.IsNullOrWhiteSpace(s)) return null;
            if (DateTime.TryParse(s, null, DateTimeStyles.AdjustToUniversal, out var dt))
                return dt;
            return null;
        }

        return new RequestLogInfo
        {
            OrderId = orderId,
            CreatedUtc = ParseUtc("CreatedUtc"),
            SentUtc = ParseUtc("SentUtc"),
            CreatedByLogin = rd["CreatedByLogin"] as string,
            CreatedByDisplayName = rd["CreatedByDisplayName"] as string,
            CreatedByEmail = rd["CreatedByEmail"] as string,
            CreatedByPhone = rd["CreatedByPhone"] as string,
            SentByLogin = rd["SentByLogin"] as string,
            SentByDisplayName = rd["SentByDisplayName"] as string,
            PdfPath = rd["PdfPath"] as string,
            LastEmailTo = rd["LastEmailTo"] as string
        };
    }

    /// <summary>
    /// Регистрирует в журнале факт формирования PDF-заявки для указанного заказа.
    /// Если запись по <paramref name="orderId"/> уже существует, она обновляется (UPSERT по ключу OrderId).
    /// В журнал записываются:
    /// - время создания (UTC, ISO 8601);
    /// - данные пользователя, сформировавшего заявку;
    /// - путь к PDF-файлу.
    /// </summary>
    public void MarkRequestCreated(int orderId, AccessUser user, string pdfPath)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO RequestLog (
    OrderId,
    CreatedUtc,
    CreatedByLogin,
    CreatedByDisplayName,
    CreatedByEmail,
    CreatedByPhone,
    PdfPath
) VALUES (
    $id,
    $createdUtc,
    $login,
    $display,
    $email,
    $phone,
    $pdf
)
ON CONFLICT(OrderId) DO UPDATE SET
    CreatedUtc           = excluded.CreatedUtc,
    CreatedByLogin       = excluded.CreatedByLogin,
    CreatedByDisplayName = excluded.CreatedByDisplayName,
    CreatedByEmail       = excluded.CreatedByEmail,
    CreatedByPhone       = excluded.CreatedByPhone,
    PdfPath              = excluded.PdfPath;";
        cmd.Parameters.AddWithValue("$id", orderId);
        cmd.Parameters.AddWithValue("$createdUtc", DateTime.UtcNow.ToString("o"));
        cmd.Parameters.AddWithValue("$login", user.Login ?? "");
        cmd.Parameters.AddWithValue("$display", user.PersonName ?? user.Login ?? "");
        cmd.Parameters.AddWithValue("$email", user.Email ?? "");
        cmd.Parameters.AddWithValue("$phone", user.Phone ?? "");
        cmd.Parameters.AddWithValue("$pdf", pdfPath ?? "");
        cmd.ExecuteNonQuery();
    }

    /// <summary>
    /// Регистрирует в журнале факт отправки заявки по электронной почте.
    /// - Если записи ещё нет, она создаётся и поля "Created*" берутся из текущего пользователя/времени.
    /// - Если запись уже существует, данные о создании сохраняются,
    ///   а поля "Sent*" и "LastEmailTo" обновляются.
    /// Такой подход позволяет восстанавливать историю: кто создал и кто отправил заявку.
    /// </summary>
    public void MarkRequestSent(int orderId, AccessUser user, string pdfPath, string toEmails)
    {
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO RequestLog (
    OrderId,
    CreatedUtc,
    CreatedByLogin,
    CreatedByDisplayName,
    CreatedByEmail,
    CreatedByPhone,
    SentUtc,
    SentByLogin,
    SentByDisplayName,
    PdfPath,
    LastEmailTo
) VALUES (
    $id,
    COALESCE((SELECT CreatedUtc           FROM RequestLog WHERE OrderId = $id), $createdUtc),
    COALESCE((SELECT CreatedByLogin       FROM RequestLog WHERE OrderId = $id), $login),
    COALESCE((SELECT CreatedByDisplayName FROM RequestLog WHERE OrderId = $id), $display),
    COALESCE((SELECT CreatedByEmail       FROM RequestLog WHERE OrderId = $id), $email),
    COALESCE((SELECT CreatedByPhone       FROM RequestLog WHERE OrderId = $id), $phone),
    $sentUtc,
    $sentLogin,
    $sentDisplay,
    $pdf,
    $to
)
ON CONFLICT(OrderId) DO UPDATE SET
    SentUtc           = excluded.SentUtc,
    SentByLogin       = excluded.SentByLogin,
    SentByDisplayName = excluded.SentByDisplayName,
    PdfPath           = excluded.PdfPath,
    LastEmailTo       = excluded.LastEmailTo;";
        cmd.Parameters.AddWithValue("$id", orderId);
        cmd.Parameters.AddWithValue("$createdUtc", DateTime.UtcNow.ToString("o"));
        cmd.Parameters.AddWithValue("$login", user.Login ?? "");
        cmd.Parameters.AddWithValue("$display", user.PersonName ?? user.Login ?? "");
        cmd.Parameters.AddWithValue("$email", user.Email ?? "");
        cmd.Parameters.AddWithValue("$phone", user.Phone ?? "");
        cmd.Parameters.AddWithValue("$sentUtc", DateTime.UtcNow.ToString("o"));
        cmd.Parameters.AddWithValue("$sentLogin", user.Login ?? "");
        cmd.Parameters.AddWithValue("$sentDisplay", user.PersonName ?? user.Login ?? "");
        cmd.Parameters.AddWithValue("$pdf", pdfPath ?? "");
        cmd.Parameters.AddWithValue("$to", toEmails ?? "");
        cmd.ExecuteNonQuery();
    }

    public void Dispose()
    {
        _conn.Dispose();
    }
}

/// <summary>
/// ���� ������ ������� ������.
/// ������������� � RequestInfo -> RequestLogInfo, ����� �� �������������
/// � ����� ������ ����� RequestInfo.
/// </summary>
public sealed class RequestLogInfo
{
    public int OrderId { get; set; }
    public DateTime? CreatedUtc { get; set; }
    public DateTime? SentUtc { get; set; }

    public string? CreatedByLogin { get; set; }
    public string? CreatedByDisplayName { get; set; }
    public string? CreatedByEmail { get; set; }
    public string? CreatedByPhone { get; set; }

    public string? SentByLogin { get; set; }
    public string? SentByDisplayName { get; set; }

    public string? PdfPath { get; set; }
    public string? LastEmailTo { get; set; }

    public DateTime? CreatedLocal => CreatedUtc?.ToLocalTime();
    public DateTime? SentLocal => SentUtc?.ToLocalTime();
}
