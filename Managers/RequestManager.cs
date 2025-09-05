using NLog;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SpravkoBot_AsSapfir
{
internal static class RequestManager
{
    private static Logger log = LogManager.GetCurrentClassLogger();

    public static Request FromJson(JsonManager jsonManager)
    {
        try
        {
            string RequestType = jsonManager.GetValue<string>("formType.title");
            string RequestBE = jsonManager.GetValue<string>("orgfilial.title");
            string RequestUUID = jsonManager.GetValue<string>("title");
            string Organization = jsonManager.GetValue<string>("organiz.title");
            DateTime DateStart = jsonManager.GetValue<DateTime>("startPeriod");
            DateTime DateEnd = jsonManager.GetValue<DateTime>("endPeriod");
            string INN = jsonManager.GetValue<string>("innString");
            string KPP = jsonManager.GetValue<string>("kppString");
            string Service = jsonManager.GetValue<string>("service.title");
            var AgreementNumbers = jsonManager.GetValue<List<string>>("regNumbDoc");
            var CounterpartyNumbers = jsonManager.GetValue<List<string>>("contragent");
            string status = jsonManager.GetValue<string>("status");
            string message = jsonManager.GetValue<string>("message");

            if (string.IsNullOrEmpty(status))
            {
                jsonManager.RemoveKey("status");
                jsonManager.AddKey("status", null);
            }

            if (string.IsNullOrEmpty(message))
            {
                jsonManager.RemoveKey("message");
                jsonManager.AddKey("message", null);
            }

            return new Request { RequestUUID = RequestUUID,
                                 RequestType = RequestType,
                                 RequestBE = RequestBE,
                                 Organization = Organization,
                                 DateStart = DateStart,
                                 DateEnd = DateEnd,
                                 INN = INN,
                                 KPP = KPP,
                                 Service = Service,
                                 AgreementNumbers = AgreementNumbers,
                                 CounterpartyNumbers = CounterpartyNumbers,
                                 status = status,
                                 message = message };
        }
        catch (Exception ex)
        {
            log.Error($"Ошибка при извлечении данных из файла заявки: {ex.Message}");
            throw new Exception($"Ошибка при извлечении данных из файла заявки: {ex.Message}");
        }
    }

    // Метод для нормализации строк (удаление лишних пробелов и приведение к единому формату)
    private static string NormalizeString(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return input;

        // Удаление лишних пробелов
        input = input.Trim();
        input = Regex.Replace(input, @"\s+", " ");

        return input;
    }

    // Метод для нормализации ФИО
    private static string NormalizeEmployeeName(string employeeName)
    {
        if (string.IsNullOrWhiteSpace(employeeName))
            return employeeName;

        // Удаление лишних пробелов
        employeeName = employeeName.Trim();
        employeeName = Regex.Replace(employeeName, @"\s+", " ");

        // Разделение на части (фамилия, имя, отчество)
        var parts = employeeName.Split(' ');
        if (parts.Length < 3)
            return employeeName; // Если частей меньше трёх, возвращаем как есть

        // Приведение каждой части к формату "С Заглавной Буквы"
        TextInfo textInfo = CultureInfo.CurrentCulture.TextInfo;
        for (int i = 0; i < parts.Length; i++)
        {
            parts[i] = textInfo.ToTitleCase(parts[i].ToLower());
        }

        return string.Join(" ", parts);
    }

    // Метод для нормализации номера сотрудника
    private static string NormalizeEmployeeNumber(string employeeNumber)
    {
        if (string.IsNullOrWhiteSpace(employeeNumber))
            return employeeNumber;

        return Regex.Replace(employeeNumber, @"\s+", "").Trim();
    }

    // Метод для проверки корректности ФИО
    private static bool IsValidEmployeeName(string employeeName)
    {
        if (string.IsNullOrWhiteSpace(employeeName))
            return false;

        // Проверка на наличие трёх частей (фамилия, имя, отчество)
        var parts = employeeName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 3)
            return false;

        // Проверка на допустимые символы (буквы и дефисы)
        var regex = new Regex(@"^[а-яА-ЯёЁa-zA-Z\-]+$");
        foreach (var part in parts)
        {
            if (!regex.IsMatch(part))
                return false;
        }

        return true;
    }
}
}
