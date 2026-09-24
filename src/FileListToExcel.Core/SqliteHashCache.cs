using Microsoft.Data.Sqlite;

namespace FileListToExcel.Core;

/// <summary>Optional persistent SHA-256 cache. All database work is serialized; any failure falls back to reading source contents.</summary>
public sealed class SqliteHashCache : IDuplicateHashCache, IDisposable
{
    private const int SchemaVersion = 1;
    private readonly object gate = new();
    private SqliteConnection? connection;
    private bool disposed;
    public string DatabasePath { get; }
    public bool IsAvailable { get { lock (gate) return connection is not null && !disposed; } }
    public string? LastError { get; private set; }
    public string? RecoveredDatabasePath { get; private set; }
    public static string DefaultPath => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "FileListToExcel", "hash_cache.sqlite");

    public SqliteHashCache(string? databasePath = null)
    {
        DatabasePath = databasePath ?? DefaultPath;
        try
        {
            DatabasePath = Path.GetFullPath(DatabasePath);
            Initialize();
        }
        catch (Exception ex) when (IsCacheFailure(ex))
        {
            CloseConnection();
            LastError = ex.Message;
            // A cache is disposable derived data, but preserve a damaged database for diagnosis.
            if ((ex is SqliteException { SqliteErrorCode: 1 or 11 or 26 } || ex is InvalidDataException) && TryQuarantine())
            {
                try { Initialize(); LastError = null; }
                catch (Exception retry) when (IsCacheFailure(retry)) { Disable(retry); }
            }
        }
    }

    public bool OwnsPath(string path)
    {
        try
        {
            string fullPath = Path.GetFullPath(path);
            var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            return fullPath.Equals(DatabasePath, comparison) ||
                new[] { "-wal", "-shm", "-journal" }.Any(suffix => fullPath.Equals(DatabasePath + suffix, comparison)) ||
                fullPath.StartsWith(DatabasePath + ".corrupt-", comparison);
        }
        catch (Exception ex) when (IsCacheFailure(ex)) { return false; }
    }

    public bool TryGet(HashCacheKey key, out HashCacheValue? value)
    {
        ArgumentNullException.ThrowIfNull(key);
        lock (gate)
        {
            value = null;
            if (connection is null || disposed) return false;
            try
            {
                using var command = connection.CreateCommand();
                command.CommandText = "SELECT quick_hash, full_sha256, last_checked_utc_ticks FROM hashes WHERE absolute_path=$path AND hash_algorithm=$algorithm AND size_bytes=$size AND modified_utc_ticks=$modified";
                AddKey(command, key);
                using var reader = command.ExecuteReader();
                if (!reader.Read()) return false;
                string? quick = ReadHash(reader, 0), full = ReadHash(reader, 1);
                if (quick is null && full is null) return false;
                long checkedTicks = reader.GetInt64(2);
                if (checkedTicks < DateTime.MinValue.Ticks || checkedTicks > DateTime.MaxValue.Ticks) return false;
                value = new(quick, full, new DateTime(checkedTicks, DateTimeKind.Utc));
                return true;
            }
            catch (Exception ex) when (IsCacheFailure(ex)) { Disable(ex); return false; }
        }
    }

    public void Store(HashCacheKey key, HashCacheValue value)
    {
        ArgumentNullException.ThrowIfNull(key);
        ArgumentNullException.ThrowIfNull(value);
        lock (gate)
        {
            if (connection is null || disposed) return;
            string? quick = NormalizeHash(value.QuickHash), full = NormalizeHash(value.FullSha256);
            if (quick is null && full is null) return;
            try
            {
                using var command = connection.CreateCommand();
                // One SQLite statement is an atomic transaction. Keep the current version of each path only.
                command.CommandText = """
                    INSERT INTO hashes(absolute_path,hash_algorithm,size_bytes,modified_utc_ticks,quick_hash,full_sha256,last_checked_utc_ticks)
                    VALUES($path,$algorithm,$size,$modified,$quick,$full,$checked)
                    ON CONFLICT(absolute_path,hash_algorithm) DO UPDATE SET
                      quick_hash=CASE WHEN hashes.size_bytes=excluded.size_bytes AND hashes.modified_utc_ticks=excluded.modified_utc_ticks THEN COALESCE(excluded.quick_hash,hashes.quick_hash) ELSE excluded.quick_hash END,
                      full_sha256=CASE WHEN hashes.size_bytes=excluded.size_bytes AND hashes.modified_utc_ticks=excluded.modified_utc_ticks THEN COALESCE(excluded.full_sha256,hashes.full_sha256) ELSE excluded.full_sha256 END,
                      size_bytes=excluded.size_bytes,
                      modified_utc_ticks=excluded.modified_utc_ticks,
                      last_checked_utc_ticks=excluded.last_checked_utc_ticks
                    """;
                AddKey(command, key);
                command.Parameters.AddWithValue("$quick", (object?)quick ?? DBNull.Value);
                command.Parameters.AddWithValue("$full", (object?)full ?? DBNull.Value);
                command.Parameters.AddWithValue("$checked", value.CheckedAtUtc.ToUniversalTime().Ticks);
                command.ExecuteNonQuery();
            }
            catch (Exception ex) when (IsCacheFailure(ex)) { Disable(ex); }
        }
    }

    private void Initialize()
    {
        EnsureSafeCachePath();
        Directory.CreateDirectory(Path.GetDirectoryName(DatabasePath)!);
        EnsureSafeCachePath();
        var opened = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = DatabasePath,
            Mode = SqliteOpenMode.ReadWriteCreate,
            Pooling = false,
            DefaultTimeout = 1
        }.ToString());
        connection = opened;
        opened.Open();
        using var command = opened.CreateCommand();
        command.CommandText = "PRAGMA user_version";
        int version = Convert.ToInt32(command.ExecuteScalar(), System.Globalization.CultureInfo.InvariantCulture);
        if (version > SchemaVersion) throw new InvalidOperationException("This hash cache uses a newer schema; caching is disabled for this run.");
        command.CommandText = """
            PRAGMA journal_mode=WAL;
            PRAGMA synchronous=NORMAL;
            CREATE TABLE IF NOT EXISTS hashes(
                absolute_path TEXT NOT NULL,
                hash_algorithm TEXT NOT NULL,
                size_bytes INTEGER NOT NULL,
                modified_utc_ticks INTEGER NOT NULL,
                quick_hash TEXT,
                full_sha256 TEXT,
                last_checked_utc_ticks INTEGER NOT NULL,
                PRIMARY KEY(absolute_path, hash_algorithm));
            PRAGMA user_version=1;
            """;
        command.ExecuteNonQuery();
        command.CommandText = "SELECT absolute_path,hash_algorithm,size_bytes,modified_utc_ticks,quick_hash,full_sha256,last_checked_utc_ticks FROM hashes LIMIT 0";
        using (command.ExecuteReader()) { }
        command.CommandText = "PRAGMA table_info(hashes)";
        using var schema = command.ExecuteReader();
        var keys = new Dictionary<string, int>();
        while (schema.Read()) keys[schema.GetString(1)] = schema.GetInt32(5);
        if (keys.Count != 7 || keys.GetValueOrDefault("absolute_path") != 1 || keys.GetValueOrDefault("hash_algorithm") != 2)
            throw new InvalidDataException("The hash cache schema is invalid.");
    }

    private void EnsureSafeCachePath()
    {
        // SQLite opens its sidecars for writing. A safe database path alone is insufficient:
        // a pre-existing sidecar must not redirect those writes or recall remote contents.
        foreach (string suffix in new[] { "-wal", "-shm", "-journal" })
        {
            string companion = DatabasePath + suffix;
            FileAttributes attributes;
            try { attributes = File.GetAttributes(companion); }
            catch (FileNotFoundException) { continue; }
            catch (DirectoryNotFoundException) { continue; }
            var reason = new LocalFileContentPolicy().GetSkipReason(companion, attributes);
            if (reason is not null || (attributes & FileAttributes.Directory) != 0)
                throw new IOException("Hash cache sidecar is not a local regular file: " + companion);
        }
        for (string? current = DatabasePath; current is not null; current = Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(current)))
        {
            FileAttributes attributes;
            try { attributes = File.GetAttributes(current); }
            catch (FileNotFoundException) { continue; }
            catch (DirectoryNotFoundException) { continue; }
            var reason = new LocalFileContentPolicy().GetSkipReason(current, attributes);
            if (reason is not null) throw new IOException("Hash cache path is not a local regular path: " + reason.ErrorType);
        }
    }

    private static void AddKey(SqliteCommand command, HashCacheKey key)
    {
        string path = Path.GetFullPath(key.AbsolutePath);
        if (OperatingSystem.IsWindows()) path = path.ToUpperInvariant();
        command.Parameters.AddWithValue("$path", path);
        command.Parameters.AddWithValue("$algorithm", key.Algorithm);
        command.Parameters.AddWithValue("$size", key.SizeBytes);
        command.Parameters.AddWithValue("$modified", key.ModifiedUtcTicks);
    }

    private static string? ReadHash(SqliteDataReader reader, int ordinal) => reader.IsDBNull(ordinal) ? null : NormalizeHash(reader.GetString(ordinal));
    private static string? NormalizeHash(string? value) => value is { Length: 64 } && value.All(Uri.IsHexDigit) ? value.ToUpperInvariant() : null;

    private bool TryQuarantine()
    {
        try
        {
            EnsureSafeCachePath();
            if (!File.Exists(DatabasePath)) return false;
            // Do not rename a database currently in use by another running application instance.
            using (new FileStream(DatabasePath, FileMode.Open, FileAccess.Read, FileShare.None)) { }
            string suffix = ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss", System.Globalization.CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N");
            string recovered = DatabasePath + suffix;
            File.Move(DatabasePath, recovered);
            foreach (string companion in new[] { "-wal", "-shm", "-journal" })
                if (File.Exists(DatabasePath + companion)) File.Move(DatabasePath + companion, recovered + companion);
            RecoveredDatabasePath = recovered;
            return true;
        }
        catch (Exception ex) when (IsCacheFailure(ex)) { LastError = ex.Message; return false; }
    }

    private static bool IsCacheFailure(Exception ex) => ex is SqliteException or IOException or UnauthorizedAccessException or
        ArgumentException or NotSupportedException or InvalidOperationException or System.Security.SecurityException or
        FormatException or OverflowException or InvalidCastException or DllNotFoundException or BadImageFormatException or TypeInitializationException;

    private void Disable(Exception error) { LastError = error.Message; CloseConnection(); }
    private void CloseConnection()
    {
        try { connection?.Dispose(); }
        catch (Exception ex) when (IsCacheFailure(ex)) { LastError = ex.Message; }
        connection = null;
    }

    public void Dispose()
    {
        lock (gate)
        {
            if (disposed) return;
            disposed = true;
            CloseConnection();
        }
    }
}
