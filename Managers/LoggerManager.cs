﻿using System;
using System.IO;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace SpravkoBot_AsSapfir
{
internal class LoggerManager
{
    private readonly string _logPath;
    private readonly int _daysToKeepLogs;

    public LoggerManager(string logPath, int daysToKeepLogs)
    {
        _logPath = logPath ?? throw new ArgumentNullException(nameof(logPath));
        _daysToKeepLogs = daysToKeepLogs;
    }

    public void ConfigureLogger()
    {
        try
        {
            var config = new LoggingConfiguration();
            var fileTarget = new FileTarget("fileTarget") {
                FileName = System.IO.Path.Combine(_logPath, $"Логирование за {DateTime.Now:yyyy-MM-dd}.log"),
                Layout = "${longdate} | ${uppercase:${level}} | ${logger} | ${callsite} | ${message} " +
                         "${exception:format=ToString}"
            };

            config.AddTarget(fileTarget);
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, fileTarget);

            LogManager.Configuration = config;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            throw new InvalidOperationException("Ошибка при настройке логгера", ex);
        }
    }

    public void ClearOldLogs()
    {

        try
        {
            if (!Directory.Exists(_logPath))
            {
                Console.WriteLine($"Папка {_logPath} не существует. Невозможно очистить логи.");
                return;
            }

            var files = Directory.GetFiles(_logPath);
            var cutoffDate = DateTime.Now.AddDays(-_daysToKeepLogs);

            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);
                if (fileInfo.LastWriteTime < cutoffDate)
                {
                    fileInfo.Delete();
                    Console.WriteLine(
                        $"Удаление старого лога: {fileInfo.Name} ({fileInfo.Length} байт) Последнее изменение: {fileInfo.LastWriteTime}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при удалении старых логов: {ex.Message}");
        }
    }
}
}
