using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;

namespace SpravkoBot_AsSapfir
{
    internal class JsonManager
    {
        private JObject _jsonObject;
        private static Logger log = LogManager.GetCurrentClassLogger();

        // Конструктор для загрузки JSON из файла или строки
        public JsonManager(string filePath, bool isFile = true)
        {
            try
            {
                if (isFile)
                {
                    // Считываем JSON из файла
                    if (!File.Exists(filePath))
                    {
                        log.Error($"Файл не найден: {filePath}");
                        throw new FileNotFoundException($"Файл не найден: {filePath}");
                    }

                    var json = File.ReadAllText(filePath);
                    _jsonObject = JObject.Parse(json);
                }
                else
                {
                    // Считываем JSON строку
                    _jsonObject = JObject.Parse(filePath);
                }
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка при загрузке JSON: {ex.Message}");
                throw new Exception($"Ошибка при загрузке JSON: {ex.Message}");
            }
        }

        // Метод для получения значения по ключу (поддерживает вложенные ключи)
        public T GetValue<T>(string key)
        {
            try
            {
                var token = _jsonObject.SelectToken(key);
                if (token != null)
                {
                    return token.ToObject<T>();
                }
                else
                {
                    log.Error($"Ключ '{key}' не найден.");
                    throw new Exception($"Ключ '{key}' не найден.");
                }
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка при получении значения по ключу '{key}': {ex.Message}");
                throw new Exception($"Ошибка при получении значения по ключу '{key}': {ex.Message}");
            }
        }

        // Метод для записи значения по ключу (если ключа нет, добавит новый)
        public void SetValue(string key, object value)
        {
            try
            {
                var token = _jsonObject.SelectToken(key);

                if (token != null)
                {
                    // Если ключ найден, заменяем его значение
                    token.Replace(JToken.FromObject(value));
                }
                else
                {
                    // Если ключ не найден, добавляем новый ключ в корневой объект JSON
                    if (_jsonObject is JObject jObject)
                    {
                        jObject.Add(new JProperty(key, value));
                    }
                    else
                    {
                        // Если _jsonObject не является JObject, выбрасываем исключение
                        log.Error("Корневой объект не является JObject.");
                        throw new Exception("Корневой объект не является JObject.");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Ошибка при записи значения по ключу '{key}': {ex.Message}");
            }
        }

        // Метод для добавления нового ключа (если ключа нет)
        public void AddKey(string key, object value)
        {
            try
            {
                if (_jsonObject is JObject jObject)
                {
                    // Добавляем новый ключ
                    jObject.Add(new JProperty(key, value));
                }
                else
                {
                    log.Error("Корневой объект не является JObject.");
                    throw new Exception("Корневой объект не является JObject.");
                }
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка при добавлении ключа '{key}': {ex.Message}");
                throw new Exception($"Ошибка при добавлении ключа '{key}': {ex.Message}");
            }
        }

        // Метод для удаления ключа
        public void RemoveKey(string key)
        {
            try
            {
                var token = _jsonObject.SelectToken(key);
                if (token != null)
                {
                    token.Parent.Remove();
                }
                else
                {
                    log.Error($"Ключ '{key}' не найден для удаления.");
                    throw new Exception($"Ключ '{key}' не найден для удаления.");
                }
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка при удалении ключа '{key}': {ex.Message}");
                throw new Exception($"Ошибка при удалении ключа '{key}': {ex.Message}");
            }
        }

        // Метод для сохранения JSON в файл
        public void SaveToFile(string filePath)
        {
            try
            {
                string jsonString = _jsonObject.ToString();
                File.WriteAllText(filePath, jsonString);
            }
            catch (Exception ex)
            {
                log.Error($"Ошибка при сохранении в файл: {ex.Message}");
                throw new Exception($"Ошибка при сохранении в файл: {ex.Message}");
            }
        }

        // Метод для преобразования JSON в строку
        public string ToJsonString()
        {
            return _jsonObject.ToString();
        }
    }
}
