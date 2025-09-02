using NLog;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SpravkoBot_AsSapfir
{
internal class ConfigManager
{
    private static Logger log = LogManager.GetCurrentClassLogger();
    private string _configPath;
    public Config Config { get; private set; }
    private JsonManager jsonManager;

    // Конструктор, который принимает путь конфигурационного файла
    public ConfigManager(string configPath)
    {
        _configPath = configPath;
        jsonManager = new JsonManager(_configPath);
        LoadConfig();
    }

    // Загрузка конфигурации из JSON
    public void LoadConfig()
    {
        try
        {
            Config = jsonManager.GetValue<Config>("$");
            log.Info($"Конфигурация успешно загружена: {Config.Organizations.Count} организаций.");
        }
        catch (Exception ex)
        {
            log.Error($"Ошибка загрузки конфигурации: {ex.Message}");
            throw new Exception($"Ошибка загрузки конфигурации: {ex.Message}");
            // Если конфигурация не загружается, сохраняем шаблон
            /*                Config = JsonConvert.DeserializeObject<Config>(GetTemplateConfig());
                            SaveConfig();*/
        }
    }

    // Сохранение конфигурации в файл
    public void SaveConfig()
    {
        try
        {
            jsonManager.SetValue("$", Config);
            jsonManager.SaveToFile(_configPath);
        }
        catch (Exception ex)
        {
            log.Error($"Ошибка сохранения конфигурации: {ex.Message}");
        }
    }

    // Метод для получения шаблона конфигурации
    public static string GetTemplateConfig()
    {
        return @"
    {
      ""ApplicationPath"": ""C:\\Program Files\\1C\\1C.exe"",
      ""Organizations"": [
        {
          ""Name"": ""Организация 1"",
          ""Database"": {
            ""Host"": ""db1.example.com"",
            ""Port"": 5432,
            ""Username"": ""user1"",
            ""Password"": ""password1"",
            ""DatabaseName"": ""org1_db""
          },
          ""Branches"": [
            {
              ""Name"": ""Филиал 1"",
              ""Signers"": [
                {
                  ""FullName"": ""Иванов Иван Иванович"",
                  ""Alias"": ""Иванов И.И.""
                },
                {
                  ""FullName"": ""Петров Петр Петрович"",
                  ""Alias"": ""Петров П.П.""
                }
              ]
            },
            {
              ""Name"": ""Филиал 2"",
              ""Signers"": [
                {
                  ""FullName"": ""Сидоров Сидор Сидорович"",
                  ""Alias"": ""Сидоров С.С.""
                }
              ]
            }
          ]
        },
        {
          ""Name"": ""Организация 2"",
          ""Database"": {
            ""Host"": ""db2.example.com"",
            ""Port"": 1433,
            ""Username"": ""user2"",
            ""Password"": ""password2"",
            ""DatabaseName"": ""org2_db""
          },
          ""Branches"": [
            {
              ""Name"": ""Филиал A"",
              ""Signers"": [
                {
                  ""FullName"": ""Алексеева Анна Алексеевна"",
                  ""Alias"": ""Алексеева А.А.""
                }
              ]
            }
          ]
        }
      ]
    }";
    }

    // Метод для получения конфигурации базы данных организации по имени
    public DatabaseConfig GetOrganizationConfig(string organizationName)
    {
        log.Info($"Ищем организацию с именем: {organizationName}");
        var organization = Config.Organizations.FirstOrDefault(
            o => o.Name.Equals(organizationName, StringComparison.OrdinalIgnoreCase));

        if (organization != null)
        {
            log.Info($"Организация с именем {organizationName} найдена.");
            return organization.Database;
        }

        log.Error($"Организация с именем {organizationName} не найдена в конфигурации.");
        return null;
    }

    // Метод для получения филиалов и подписантов по организации
    public Dictionary<string, List<(string FullName, string Alias)>> GetBranchesAndSigners(string organizationName)
    {
        var organization = Config.Organizations.FirstOrDefault(
            o => o.Name.Equals(organizationName, StringComparison.OrdinalIgnoreCase));

        if (organization == null)
        {
            log.Error($"Организация с именем {organizationName} не найдена в конфигурации.");
            return null;
        }

        var branchesInfo = new Dictionary<string, List<(string FullName, string Alias)>>();

        foreach (var branch in organization.Branches)
        {
            var signersList = branch.Signers.Select(s => (s.FullName, s.Alias)).ToList();

            branchesInfo[branch.Name] = signersList;
        }

        return branchesInfo;
    }
}

// Классы для хранения конфигурационных данных
public class Config
{
    public string ApplicationPath { get; set; }
    public List<Organization> Organizations { get; set; } = new List<Organization>();
}

public class Organization
{
    public string Name { get; set; }
    public DatabaseConfig Database { get; set; }
    public List<Branch> Branches { get; set; } = new List<Branch>();
}

public class DatabaseConfig
{
    public string Host { get; set; }
    public int Port { get; set; }
    public string Username { get; set; }
    public string Password { get; set; }
    public string DatabaseName { get; set; }
}

public class Branch
{
    public string Name { get; set; }
    public List<Signer> Signers { get; set; } = new List<Signer>();
}

public class Signer
{
    public string FullName { get; set; }
    public string Alias { get; set; }
}
}
