using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpravkoBot_AsSapfir
{
    internal class AppInitManager
    {
        private readonly string _baseAppPath;
        private readonly string _configFilePath;
        private readonly string _requestsPath;
        private readonly string[] _subdirectories = new string[] { "input", "error", "output", "logs", "temp" };
        private readonly LoggerManager _loggerManager;
        public static Logger log;

        public AppInitManager()
        {
            _baseAppPath = AppDomain.CurrentDomain.BaseDirectory;
            _configFilePath = Path.Combine(_baseAppPath, "config.json");
            _requestsPath = Path.Combine(_baseAppPath, "data", "requests");
            _loggerManager = new LoggerManager(
                Path.Combine(_baseAppPath, _requestsPath, "logs"), // Путь до папки с логами
                30                                                 // Количество дней хранения логов    
            );
        }

        public void Init()
        {
            _loggerManager.ConfigureLogger();
            _loggerManager.ClearOldLogs();
            log = LogManager.GetCurrentClassLogger();

            CreateDirectory();
            CreateConfigFile(_configFilePath);

            string tempFolder = Path.Combine(_requestsPath, "temp");
            ClearFilesFolders(tempFolder);
        }

        private void CreateDirectory()
        {
            try
            {
                if (!Directory.Exists(_requestsPath))
                {
                    Directory.CreateDirectory(_requestsPath);
                }

                foreach (var subdirectory in _subdirectories)
                {
                    string fullPath = Path.Combine(_requestsPath, subdirectory);

                    if (!Directory.Exists(fullPath))
                    {
                        Directory.CreateDirectory(fullPath);
                    }
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                log.Error($"Ошибка доступа: {ex.Message}");
            }
            catch (IOException ex)
            {
                log.Error($"Ошибка ввода-вывода: {ex.Message}");
            }
            catch (Exception ex)
            {
                log.Error($"Неизвестная ошибка: {ex.Message}");
            }
        }

        private void CreateConfigFile(string path)
        {
            try
            {
                // Проверка существования файла и его содержимого
                if (!File.Exists(path) || new FileInfo(path).Length == 0)
                {
                    File.WriteAllText(path, ConfigManager.GetTemplateConfig());
                    log.Info("Файл конфигурации config.json создан и заполнен данными.");
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                log.Error($"Ошибка доступа при создании config.json: {ex.Message}");
            }
            catch (IOException ex)
            {
                log.Error($"Ошибка ввода-вывода при создании config.json: {ex.Message}");
            }
            catch (Exception ex)
            {
                log.Error($"Неизвестная ошибка при создании config.json: {ex.Message}");
            }
        }

        public Dictionary<string, string> GetAppFolders()
        {
            return _subdirectories.ToDictionary(subdir => subdir, subdir => Path.Combine(_requestsPath, subdir));
        }

        public string GetPathConfigFile()
        {
            return _configFilePath.ToString();
        }

        public void ClearFilesFolders(string pathFolder)
        {
            try
            {
                if (!Directory.Exists(pathFolder))
                {
                    log.Error($"Не найдена папка {pathFolder}.");
                }

                string[] files = Directory.GetFiles(pathFolder);

                foreach (string file in files)
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch (Exception e)
                    {
                        log.Error($"Ошибка при удалении файла: {e.ToString()}");
                    }
                }

            }
            catch (Exception e)
            {
                log.Error($"Ошибка при удалении файлов в папке {pathFolder}. Ошибка: {pathFolder}");
            }
        }

    }
}
