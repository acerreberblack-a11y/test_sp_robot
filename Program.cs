using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using NLog;
using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace SpravkoBot_AsSapfir
{
internal class Program
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();
    private static JsonManager jsonManager;
    private static ConfigManager configManager;
    private static Dictionary<string, string> appFolders;
    private static Dictionary<string, string> beCodes;

    private static string signatory1 = string.Empty;
    private static string signatory2 = string.Empty;

    [STAThread]
    static void Main(string[] args)
    {
        try
        {
            AppInitManager initializer = new AppInitManager();
            initializer.Init();
            log.Info("========================  Запускаю SpravkoBot_asSapfir. Робот начинает работу.  " +
                     "========================");

            appFolders = initializer.GetAppFolders();
            log.Info("Инициализация успешно завершена.");

            string configPath = initializer.GetPathConfigFile();
            configManager = new ConfigManager(configPath);
            var config = configManager.Config;
            string excelPath = config.ExcelPath;
            string sapLogonPath = config.SapLogonPath;
            string stage = config.SapStage;
            string testStage = config.SapTestStage;
            string sapUser = config.SapUser;
            string sapPassword = config.SapPassword;
            beCodes = config.BeCodes;

            if (!appFolders.TryGetValue("input", out string inputFolder) || string.IsNullOrWhiteSpace(inputFolder))
            {
                log.Error("Инициализация входной папки input произошла с ошибкой.");
                return;
            }

            var files = Directory.EnumerateFiles(inputFolder, "SD*.txt")
                            .Where(file => Path.GetFileName(file).Contains("+"))
                            .ToList();

            if (!files.Any())
            {
                log.Info("Нет заявок для обработки. Робот завершает работу.");
                return;
            }

            log.Info($"В папке {inputFolder} найдено {files.Count} входящих файл(-а)(-ов) для обработки.");

            foreach (var file in files)
            {
                string fileName = Path.GetFileName(file);
                log.Info($"Обработка заявки: {fileName}. Начинаю извлечение данных из заявки.");

                try
                {
                    jsonManager = new JsonManager(file, true);
                    RequestManager requestManager = RequestManager.FromJson(jsonManager);

                    log.Info("Данные из заявки извлечены:");
                    log.Info("********************************************");
                    log.Info($"* Номер заявки => {requestManager.RequestUUID}");
                    log.Info($"* Тип заявки => {requestManager.RequestType}");
                    log.Info($"* Услуга => {requestManager.Service}");
                    log.Info($"* Организация => {requestManager.Organization}");
                    log.Info($"*..........................................*");
                    log.Info($"* БЕ => {requestManager.RequestBE}");
                    log.Info($"* ИНН => {requestManager.INN}");
                    log.Info($"* КПП => {requestManager.KPP}");
                    log.Info($"* Номер (-а) договора в системе => {requestManager.AgreementNumber}");
                    log.Info($"* Начало периода => {requestManager.DateStart:dd.MM.yyyy}");
                    log.Info($"* Конец периода => {requestManager.DateEnd:dd.MM.yyyy}");
                    log.Info("**********************************************");

                    string folderRequest = "";
                    if (appFolders.TryGetValue("output", out string outputFolder))
                    {
                        folderRequest = CreateRequestFolder(requestManager.RequestUUID, outputFolder, file);
                    }
                    else
                    {
                        log.Error("Не удалось получить путь к папке output из appFolders.");
                        throw new Exception("Не удалось получить путь к папке output.");
                    }

                    string jsonFile = Path.Combine(folderRequest, $"{requestManager.RequestUUID}+.txt");
                    jsonManager = new JsonManager(jsonFile, true);
                    requestManager = RequestManager.FromJson(jsonManager);

                    log.Info($"Начинаю конвертирование excel файла в csv: {excelPath}");
                    var converter = new ExcelConverter(excelPath);
                    string csvPath = converter.ConvertToCsv();

                    if (!File.Exists(csvPath))
                    {
                        throw new FileNotFoundException("CSV-файл не найден!", csvPath);
                    }

                    Console.WriteLine(
                        $"Выполняю фильтрацию по параметрам из заявки: БЕ:{requestManager.RequestBE}, ИНН:{requestManager.INN}, КПП:{requestManager.KPP}");
                    var searchNames = new HashSet<string> { requestManager.RequestBE.ToString() };
                    var searchInn = new HashSet<string> { requestManager.INN.ToString() };
                    var searchKpp = new HashSet<string> { requestManager.KPP.ToString() };

                    var filteredRows =
                        File.ReadLines(csvPath)
                            .Skip(1)
                            .Select(line => line.Split(',').Select(c => c.Trim('"').Trim()).ToArray())
                            .Where(columns => columns.Length > 10 && !string.IsNullOrWhiteSpace(columns[1]) &&
                                              !string.IsNullOrWhiteSpace(columns[4]) &&
                                              !string.IsNullOrWhiteSpace(columns[5]) &&
                                              searchNames.Contains(columns[1]) && searchInn.Contains(columns[4]) &&
                                              searchKpp.Contains(columns[5]))
                            .Select(columns => new { Number = columns.ElementAtOrDefault(0) ?? "",
                                                     Branch = columns.ElementAtOrDefault(1) ?? "",
                                                     CompanyName = columns.ElementAtOrDefault(2) ?? "",
                                                     CompanyNumberSap = columns.ElementAtOrDefault(3) ?? "",
                                                     INN = columns.ElementAtOrDefault(4) ?? "",
                                                     KPP = columns.ElementAtOrDefault(5) ?? "",
                                                     Status = columns.ElementAtOrDefault(6) ?? "",
                                                     SignatoryLanDocs = columns.ElementAtOrDefault(7) ?? "",
                                                     PersonnelNumber = columns.ElementAtOrDefault(8) ?? "",
                                                     SignatoryLandocs = columns.ElementAtOrDefault(9) ?? "",
                                                     VGO = columns.ElementAtOrDefault(10) ?? "" })
                            .ToList<dynamic>();

                    log.Info($"Найдено строк: {filteredRows.Count}. ");

                    if (filteredRows.Count >= 2)
                    {
                        log.Error("При выборке по ИНН и КПП нашлось несколько результатов. Просьба проверить " +
                                  "корректность данных.");
                        foreach (var row in filteredRows)
                        {
                            log.Error(string.Join(" | ", row));
                        }
                        throw new Exception("При выборке по ИНН и КПП нашлось несколько результатов. Просьба " +
                                            "проверить корректность данных.");
                    }

                    foreach (var row in filteredRows)
                    {
                        Console.WriteLine(string.Join(" | ", row));
                    }

                    string id = filteredRows.FirstOrDefault()?.Number ?? "";
                    string branch = filteredRows.FirstOrDefault()?.Branch ?? "";
                    string companyName = filteredRows.FirstOrDefault()?.CompanyName ?? "";
                    string companyNumberSap = filteredRows.FirstOrDefault()?.CompanyNumberSap ?? "";
                    string status = filteredRows.FirstOrDefault()?.Status ?? "";
                    string personnelNumber = filteredRows.FirstOrDefault()?.PersonnelNumber ?? "";

                    // Проверяем, что значения не пустые и не null
                    if (string.IsNullOrWhiteSpace(companyNumberSap))
                    {
                        throw new Exception("Ошибка: CompanyNumberSap не может быть пустым или null.");
                    }

                    if (string.IsNullOrWhiteSpace(personnelNumber))
                    {
                        throw new Exception("Ошибка: PersonnelNumber не может быть пустым или null.");
                    }

                    log.Info($"Результаты фильтрации:");
                    log.Info("********************************************");
                    log.Info($"Номер: {id}");
                    log.Info($"Филиал ПАО РГ: {branch}");
                    log.Info($"Наименование КА: {companyName}");
                    log.Info($"Номер ДП в САП: {companyNumberSap}");
                    log.Info($"Статус: {status}");
                    log.Info($"Табельный номер подписантов: {personnelNumber}");
                    log.Info("********************************************");

                    if (!string.IsNullOrWhiteSpace(personnelNumber))
                    {
                        string[] signatories =
                            personnelNumber.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                        signatory1 = signatories[0].Trim();
                        signatory2 = signatories.Length > 1 ? signatories[1].Trim() : signatory1;

                        Console.WriteLine(
                            $"Выбраны подписанты: Подписант 1 — {signatory1}, Подписант 2 — {signatory2}");
                    }
                    else
                    {
                        Console.WriteLine("Значение personnelNumber пустое или null.");
                    }

                    try
                    {
                        SapfirManager sapfir = new SapfirManager(sapLogonPath);
                        sapfir.LaunchSAP();
                        log.Info("SAP Logon успешно запущен.");

                        // 1. Продуктивная среда САПФИР
                        // ER2 - среда тестирования САПФИР
                        GuiSession session = sapfir.GetSapSession(stage);
                        if (session == null)
                        {
                            throw new Exception("Не удалось получить сессию SAP.");
                        }
                        log.Info("Сессия SAP успешно получена.");

                        if (stage == testStage)
                        {
                            sapfir.LoginToSAP(session, sapUser, sapPassword);
                        }
                        string statusBarValue = sapfir.GetStatusMessage();

                        if (statusBarValue.Contains("Этот мандант сейчас блокирован для регистрации в нём."))
                        {
                            log.Error("Возникла ошибка при попытке входа в САП. Ошибка: Этот мандант сейчас " +
                                      "блокирован для регистрации в нём.");
                            throw new Exception("Возникла ошибка при попытке входа в САП. Ошибка: Этот мандант " +
                                                "сейчас блокирован для регистрации в нём.");
                        }

                        log.Info("Успешно выполнен вход в SAP.");

                        // Добавляем задержку для стабилизации сессии
                        Thread.Sleep(2000);

                        try
                        {
                            session.StartTransaction("ZTSF_AKT_SVERKI");
                            log.Info("Транзакция ZTSF_AKT_SVERKI успешно запущена.");
                        }
                        catch (Exception ex)
                        {
                            log.Error($"Ошибка при запуске транзакции ZTSF_AKT_SVERKI: {ex.Message}");
                            throw;
                        }

                        log.Info("Успешно выполнен вход и запущена транзакция ZTSF_AKT_SVERKI");

                        string be =
                            GetValueByName(requestManager.RequestBE) ??
                            throw new Exception(
                                "Не смог сопоставить БЕ с заявки из списка БЕ в конфиге, так как она отсутствует");

                        /*
                         * Описание полей SAP
                         *
                         * wnd[0]/usr/ctxtP_BUKRS - Поле ввода БЕ
                         * wnd[0]/usr/ctxtS_BUDAT-LOW - Поле ввода начальной даты
                         * wnd[0]/usr/ctxtS_BUDAT-HIGH - Поле ввода конечной даты
                         */

                        sapfir.SetText("wnd[0]/usr/ctxtP_BUKRS", be);
                        Thread.Sleep(500);
                        sapfir.SetText("wnd[0]/usr/ctxtS_BUDAT-LOW", requestManager.DateStart.ToString("dd.MM.yyyy"));
                        Thread.Sleep(500);
                        sapfir.SetText("wnd[0]/usr/ctxtS_BUDAT-HIGH", requestManager.DateEnd.ToString("dd.MM.yyyy"));

                        Thread.Sleep(1000);
                        sapfir.RadioButton("wnd[0]/usr/radP_PROCH", true);
                        Thread.Sleep(1000);

                        /*
                         * Описание полей SAP
                         *
                         * wnd[0]/usr/radP_ELEC - Радиокнопка Электроэнергия
                         * wnd[0]/usr/radP_CAP - Радиокнопка Мощность
                         * wnd[0]/usr/radP_PROCH - Радиокнопка Прочие акты
                         */
                        /*
                         * Пункты Прочие акты
                         *
                         * wnd[0]/usr/chkP_PRSLD - Развернутое сальдо
                         * wnd[0]/usr/chkP_PRSPP - Развернутое сальдо с СПП
                         * wnd[0]/usr/chkP_PRALL - Акт сверки по всем договорам
                         * wnd[0]/usr/chkP_PRHKT - Выбор счетов
                         * wnd[0]/usr/chkP_WAERS - Акт сверки в Валюте, УЕ
                         *
                         */

                        sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSLD", false);
                        sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRSPP", false);
                        Thread.Sleep(500);
                        sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRHKT", true);
                        sapfir.GuiCheckBox("wnd[0]/usr/chkP_WAERS", false);

                        /*
                         * Номер договора
                         *
                         * wnd[0]/usr/ctxtS_ZUONR-LOW - Поле ввода С
                         * wnd[0]/usr/txtS_ZUONR-HIGH - Поле ввода ПО
                         * wnd[0]/usr/btn%_S_ZUONR_%_APP_%-VALU_PUSH - Кнопка многократный ввод
                         */
                        /*
                         * Контрагент
                         *
                         * wnd[0]/usr/ctxtS_PARTN-LOW - Поле ввода С
                         * wnd[0]/usr/ctxtS_PARTN-HIGH - Поле ввода ПО
                         * wnd[0]/usr/btn%_S_PARTN_%_APP_%-VALU_PUSH - Кнопка многократный ввод
                         */
                        Thread.Sleep(500);
                        sapfir.SetText("wnd[0]/usr/ctxtS_PARTN-LOW", companyNumberSap);
                        /*
                         * Счёт
                         *
                         * wnd[0]/usr/ctxtS_HKONT-LOW - Поле ввода С
                         * wnd[0]/usr/ctxtS_HKONT-HIGH - Поле ввода ПО
                         * wnd[0]/usr/btn%_S_HKONT_%_APP_%-VALU_PUSH - Кнопка многократный ввод
                         */
                        // Нажатие кнопки для открытия окна выбора счета
                        sapfir.PressButton("wnd[0]/usr/btn%_S_HKONT_%_APP_%-VALU_PUSH");
                        Thread.Sleep(2000);
                        // Установка "*" для выбора всех счетов на вкладке SIVA
                        sapfir.SetText("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" +
                                           "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]",
                                       "*");
                        // Переключение на вкладку NOSV
                        sapfir.SelectTab("wnd[1]/usr/tabsTAB_STRIP/tabpNOSV");
                        Thread.Sleep(1000);
                        // Ввод значений для счетов 62
                        string[] accounts62 = new string[] { "6201010101", "6201010201", "6201010301", "6201010401",
                                                             "6201010501", "6201010601", "6201020101", "6201030101" };

                        for (int i = 0; i < accounts62.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts62[i]);
                        }

                        sapfir.SetVerticalScrollPosition(
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E", 7);
                        Thread.Sleep(1000);

                        string[] accounts62_2 = new string[] { "6201030201", "6201030301", "6201030401", "6201030501",
                                                               "6201030601", "6201030701", "6201040201" };

                        for (int i = 1; i <= accounts62_2.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts62_2[i - 1]);
                        }

                        // Прокрутка таблицы до позиции 14
                        sapfir.SetVerticalScrollPosition(
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                            14);
                        Thread.Sleep(1000);

                        string[] accounts62_3 = new string[] { "6201040301", "6201110101", "6201130101", "6202010101",
                                                               "6202010201", "6202010301", "6202010401" };

                        for (int i = 1; i <= accounts62_3.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts62_3[i - 1]);
                        }

                        // Прокрутка таблицы до позиции 21
                        sapfir.SetVerticalScrollPosition(
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                            21);
                        Thread.Sleep(1000);

                        string[] accounts62_4 = new string[] { "6202010501", "6202010601", "6202020101", "6202030101",
                                                               "6202030201", "6202030301", "6202030401" };

                        for (int i = 1; i <= accounts62_4.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts62_4[i - 1]);
                        }

                        // Прокрутка таблицы до позиции 28
                        sapfir.SetVerticalScrollPosition(
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                            28);
                        Thread.Sleep(1000);

                        string[] accounts62_5 = new string[] { "6202030501", "6202030601", "6202030701", "6202040201",
                                                               "6202040301", "6202110101", "6202120101" };

                        for (int i = 1; i <= accounts62_5.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts62_5[i - 1]);
                        }

                        // Прокрутка таблицы до позиции 35
                        sapfir.SetVerticalScrollPosition(
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                            35);
                        Thread.Sleep(1000);

                        string[] accounts76_1 = new string[] { "6202130101", "7602020101", "7602040101", "7611010101",
                                                               "7615020101", "7602010102", "7602020102" };

                        for (int i = 1; i <= accounts76_1.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts76_1[i - 1]);
                        }

                        // Прокрутка таблицы до позиции 42
                        sapfir.SetVerticalScrollPosition(
                            "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E",
                            42);
                        Thread.Sleep(1000);

                        string[] accounts76_2 =
                            new string[] { "7602030102", "7602010101", "7602040102", "7615020102", "760903*" };

                        for (int i = 1; i <= accounts76_2.Length; i++)
                        {
                            sapfir.SetText(
                                $"wnd[1]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,{i}]",
                                accounts76_2[i - 1]);
                        }
                        sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                        Thread.Sleep(2000);

                        string serviceType = requestManager.RequestType.ToString();

                        if (serviceType == "По одному контрагенту по всем договорам")
                        {
                            sapfir.GuiCheckBox("wnd[0]/usr/chkP_DETAIL", true);
                            sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRALL", true);
                        }

                        string agreementNumber = requestManager.AgreementNumber;

                        if (serviceType == "По одному договору")
                        {
                            sapfir.GuiCheckBox("wnd[0]/usr/chkP_PRALL", false);

                            string[] agreementNumbers =
                                requestManager.AgreementNumber
                                    ?.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(n => n.Trim())
                                    .ToArray() ??
                                Array.Empty<string>();

                            sapfir.PressButton("wnd[0]/usr/btn%_S_ZUONR_%_APP_%-VALU_PUSH");
                            Thread.Sleep(2000);

                            foreach (var number in agreementNumbers)
                            {
                                if (!Regex.IsMatch(number.ToString(), @"^[\d/\-\.]+$"))
                                    throw new ArgumentException(
                                        $"Номер договора '{number}' должен содержать только цифры и символы /, -, .");
                            }

                            Clipboard.SetText(string.Join("\n", agreementNumbers));

                            sapfir.PressButton("wnd[1]/tbar[0]/btn[24]");
                            Thread.Sleep(3000);
                            sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");
                        }
                        Thread.Sleep(1000);
                        sapfir.GuiCheckBox("wnd[0]/usr/chkP_AGREE", true);
                        sapfir.GuiCheckBox("wnd[0]/usr/chkP_NULL", false);

                        log.Info($"Селекционный экран заполнен.");
                        sapfir.PressButton("wnd[0]/tbar[1]/btn[8]");
                        Thread.Sleep(3000);
                        string windowActive = sapfir.GetFrameText();
                        statusBarValue = sapfir.GetStatusMessage();

                        if (statusBarValue.Contains("не найдены"))

                            log.Info(
                                $"Ожидаю появления окна \"Акт сверки расчетов с контрагентами: ALV отчет\", текущее окно: {windowActive}");
                        if (windowActive != "Акт сверки расчетов с контрагентами: ALV отчет" &&
                            !string.IsNullOrWhiteSpace(statusBarValue))
                        {
                            log.Error(
                                $"Данные по контрагенту не найдены или допущены ошибки в заполнении. Статус окна: {statusBarValue}. Перехожу к следующей заявке.");
                            throw new Exception();
                        }

                        // Получение объекта shell и приведение к GuiGridView
                        GuiShell shell = session.FindById("wnd[0]/shellcont/shell") as GuiShell;
                        GuiGridView grid = shell as GuiGridView;
                        if (grid != null)
                        {
                            // Установка текущей ячейки в таблице (-1, "" означает первую ячейку)
                            grid.SetCurrentCell(-1, "");

                            // Выбор всех строк/столбцов в таблице
                            grid.SelectAll();
                        }
                        else
                        {
                            throw new Exception("Не удалось привести shell к типу GuiGridView");
                        }

                        // Нажатие кнопки "Выбрать выделенные записи"
                        sapfir.PressButton("wnd[0]/tbar[1]/btn[16]");

                        // Пауза 2 секунды
                        Thread.Sleep(2000);

                        // Нажатие кнопки "Сформировать документы"
                        sapfir.PressButton("wnd[0]/tbar[1]/btn[13]");

                        // Пауза 5 секунд для завершения формирования
                        Thread.Sleep(5000);

                        bool buttonGenerateDocuments = false;

                        // Проверка наличия кнопки и окна повторного формирования документов.
                        try
                        {
                            GuiButton button = session.FindById("wnd[1]/usr/btnBUTTON_1") as GuiButton;
                            if (button != null)
                            {
                                sapfir.PressButton(
                                    "wnd[1]/usr/btnBUTTON_1"); // Нажатие кнопки "Сформировать еще раз", если она есть
                                buttonGenerateDocuments = true;

                                log.Warn("По данному контрагенту уже есть сформированные документы. Выполнено " +
                                         "повторное формирование.");

                                Thread.Sleep(2000);
                            }
                        }
                        catch
                        {
                            buttonGenerateDocuments = false; // Устанавливаем ошибку, если элемент недоступен
                        }

                        // Первая часть: Установка первого подписанта
                        sapfir.ShowContextText("wnd[1]/usr/ctxtP_RUKOV1"); // Установка фокуса
                        sapfir.SendKey("wnd[1]", 4);                       // Отправка F4
                        sapfir.PressButton("wnd[2]/tbar[0]/btn[17]");      // Нажатие кнопки
                        Thread.Sleep(1000);                                // Пауза 1 секунда

                        // Заполнение полей для signatory1
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]",
                                       signatory1);
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[1,24]",
                                       "");
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]",
                                       "");
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]",
                                       "");
                        sapfir.ShowContextText(
                            "wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                            "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]"); // Установка фокуса
                        Thread.Sleep(1000);                                  // Пауза 1 секунда

                        // Подтверждение выбора
                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");
                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");

                        sapfir.ShowContextText("wnd[1]/usr/ctxtP_RUKOV1");

                        // Вторая часть: Установка второго подписанта
                        sapfir.ShowContextText("wnd[1]/usr/ctxtP_BUGAL2"); // Установка фокуса
                        Thread.Sleep(1000);                                // Пауза 1 секунда

                        sapfir.SendKey("wnd[1]", 4);                  // Отправка F4
                        sapfir.PressButton("wnd[2]/tbar[0]/btn[17]"); // Нажатие кнопки

                        // Заполнение полей для signatory2
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]",
                                       signatory2);
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[1,24]",
                                       "");
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]",
                                       "");
                        sapfir.SetText("wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                                           "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]",
                                       "");
                        sapfir.ShowContextText(
                            "wnd[2]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/" +
                            "sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[6,24]"); // Установка фокуса
                        Thread.Sleep(1000);                                  // Пауза 1 секунда

                        // Подтверждение выбора второго подписанта
                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");
                        Thread.Sleep(1000); // Пауза 1 секунда
                        sapfir.PressButton("wnd[2]/tbar[0]/btn[0]");
                        Thread.Sleep(1000); // Пауза 1 секунда

                        log.Info(sapfir.GetTextFromField("wnd[1]/usr/ctxtP_BUGAL2"));

                        // Финальные действия
                        sapfir.PressButton("wnd[1]/tbar[0]/btn[8]");  // Зеленая галочка
                        Thread.Sleep(1000);                           // Пауза 1 секунда
                        sapfir.PressButton("wnd[1]/usr/btnBUTTON_1"); // Альбомная ориентация
                        Thread.Sleep(1000);                           // Пауза 1 секунда

                        statusBarValue = sapfir.GetStatusMessage();

                        Console.WriteLine(statusBarValue);

                        // Фильтр на ошибки

                        // Снятие выделения
                        sapfir.PressButton("wnd[0]/tbar[1]/btn[17]"); // Кнопка снятия выделения

                        // Установка текущей ячейки в таблице на первую (-1, "")
                        shell = session.FindById("wnd[0]/shellcont/shell") as GuiShell;
                        grid = shell as GuiGridView;
                        if (grid != null)
                        {
                            grid.SetCurrentCell(-1, "");

                            // Выделение всех столбцов
                            grid.SelectAll();

                            // Нажатие на кнопку фильтра "&MB_FILTER"
                            grid.PressToolbarButton("&MB_FILTER");

                            // Нажатие кнопки для ввода значений фильтра
                            sapfir.PressButton(
                                "wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH");

                            // Выбор вкладки NOSV
                            sapfir.SelectTab("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV");

                            // Установка значения "@5b@" в поле исключений
                            sapfir.SetText("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/" +
                                               "tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]",
                                           "@5b@");

                            // Подтверждение ввода значений (кнопка btn[8])
                            sapfir.PressButton("wnd[2]/tbar[0]/btn[8]");

                            // Применение фильтра (кнопка btn[0])
                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");
                            Thread.Sleep(5000); // Ожидание 10 секунд, как в оригинальном скрипте

                            // Повторное выделение всех столбцов
                            grid.SelectAll();

                            // Нажатие кнопки "Выбрать выделенные записи"
                            sapfir.PressButton("wnd[0]/tbar[1]/btn[16]");

                            // Открытие контекстного меню экспорта "&MB_EXPORT"
                            grid.PressToolbarContextButton("&MB_EXPORT");

                            // Выбор пункта "&PC" в контекстном меню
                            grid.SelectContextMenuItem("&PC");

                            // Выбор опции "Электронная таблица" и установка фокуса
                            sapfir.RadioButton("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/" +
                                                   "radSPOPLI-SELFLAG[1,0]",
                                               true);

                            // Подтверждение выбора
                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");

                            //
                            appFolders.TryGetValue("temp", out string tempFolder);

                            // Установка пути и имени файла
                            sapfir.SetText("wnd[1]/usr/ctxtDY_PATH", tempFolder);         // Путь для сохранения
                            sapfir.SetText("wnd[1]/usr/ctxtDY_FILENAME", "errorsAC.xls"); // Имя файла

                            // Финальное подтверждение сохранения
                            sapfir.PressButton("wnd[1]/tbar[0]/btn[0]");

                            Thread.Sleep(1000); // Ожидание 10 секунд, как в оригинальном скрипте

                            sapfir.PressButton("wnd[0]/tbar[1]/btn[20]");
                            log.Info(@"Нажали на кнопку ""Сформировать документы""");
                            string savePath = @"C:\Users\RobinSapAC\Desktop\";

                            // Добавляем логику сохранения файла
                            savePath = @"C:\Users\RobinSapAC\Desktop"; // Путь к папке Desktop
                            // string fileNameToSave =
                            // $"AktSverki_{requestManager.RequestBE}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf"; // Пример
                            // имени файла
                            string fullPath = Path.Combine(savePath);

                            // Ожидаем появления окна сохранения
                            Thread.Sleep(2000);

                            const int MaxAttempts = 3;
                            const int TimeoutSeconds = 30;

                            using (var automation = new UIA3Automation())
                            {
                                // Поиск главного окна SAP
                                bool mainWindowFound = false;
                                for (int attempt = 1; attempt <= MaxAttempts && !mainWindowFound; attempt++)
                                {
                                    try
                                    {
                                        // Получаем десктоп
                                        var desktop = automation.GetDesktop();

                                        // Ожидание появления главного окна SAP "Акт сверки расчетов с контрагентами:
                                        // ALV отчет"
                                        var mainWindow =
                                            desktop
                                                .FindFirstChild(
                                                    cf => cf.ByName("Акт сверки расчетов с контрагентами: ALV отчет")
                                                              .Or(cf.ByControlType(
                                                                  FlaUI.Core.Definitions.ControlType.Window)))
                                                ?.AsWindow();

                                        DateTime startTime = DateTime.Now;
                                        while (mainWindow == null &&
                                               (DateTime.Now - startTime).TotalSeconds < TimeoutSeconds)
                                        {

                                            Thread.Sleep(1000); // Проверяем каждую секунду
                                            mainWindow =
                                                desktop
                                                    .FindFirstChild(
                                                        cf =>
                                                            cf.ByName("Акт сверки расчетов с контрагентами: ALV отчет")
                                                                .Or(cf.ByControlType(
                                                                    FlaUI.Core.Definitions.ControlType.Window)))
                                                    ?.AsWindow();
                                        }

                                        if (mainWindow == null)
                                        {
                                            log.Info(
                                                $"Попытка {attempt}: Окно 'Акт сверки расчетов с контрагентами: ALV отчет' не найдено");
                                            if (attempt == MaxAttempts)
                                            {
                                                throw new Exception("Главное окно SAP не найдено после всех попыток.");
                                            }
                                            continue;
                                        }

                                        log.Info("Окно 'Акт сверки расчетов с контрагентами: ALV отчет' найдено");

                                        // Ожидание дочернего окна внутри главного окна
                                        var childWindow =
                                            mainWindow
                                                .FindFirstChild(
                                                    cf => cf.ByName("Акт сверки расчетов с контрагентами: ALV отчет")
                                                              .Or(cf.ByControlType(
                                                                  FlaUI.Core.Definitions.ControlType.Window)))
                                                ?.AsWindow();

                                        startTime = DateTime.Now;
                                        while (childWindow == null &&
                                               (DateTime.Now - startTime).TotalSeconds < TimeoutSeconds)
                                        {
                                            Thread.Sleep(1000); // Проверяем каждую секунду
                                            childWindow =
                                                mainWindow
                                                    .FindFirstChild(
                                                        cf =>
                                                            cf.ByName("Акт сверки расчетов с контрагентами: ALV отчет")
                                                                .Or(cf.ByControlType(
                                                                    FlaUI.Core.Definitions.ControlType.Window)))
                                                    ?.AsWindow();
                                        }

                                        if (childWindow == null)
                                        {
                                            log.Info(
                                                $"Попытка {attempt}: Дочернее окно 'Акт сверки расчетов с контрагентами: ALV отчет' не найдено");
                                            if (attempt == MaxAttempts)
                                            {
                                                throw new Exception("Дочернее окно SAP не найдено после всех попыток.");
                                            }
                                            continue;
                                        }

                                        log.Info(
                                            "Дочернее окно 'Акт сверки расчетов с контрагентами: ALV отчет' найдено");

                                        // Отправка клавиши Enter
                                        SendKeys.SendWait("{ENTER}");
                                        mainWindowFound = true; // Успешно нашли и обработали окно
                                        log.Info("Клавиша Enter успешно отправлена в окно SAP.");
                                    }
                                    catch (Exception ex)
                                    {
                                        log.Error(
                                            $"Попытка {attempt}: Ошибка при поиске окна SAP или отправке Enter - {ex.Message}");
                                        if (attempt == MaxAttempts)
                                        {
                                            throw new Exception(
                                                "Не удалось обработать главное окно SAP после всех попыток.", ex);
                                        }
                                        Thread.Sleep(1000); // Пауза перед следующей попыткой
                                    }
                                }

                                // Поиск окна сохранения
                                for (int attempt = 1; attempt <= MaxAttempts; attempt++)
                                {
                                    try
                                    {
                                        // Получаем десктоп
                                        var desktop = automation.GetDesktop();

                                        // Ожидание появления окна "Browse for Files or Folders"
                                        var browseWindow =
                                            desktop
                                                .FindFirstChild(cf =>
                                                                    cf.ByName("Browse for Files or Folders")
                                                                        .Or(cf.ByControlType(
                                                                            FlaUI.Core.Definitions.ControlType.Window)))
                                                ?.AsWindow();

                                        DateTime startTime = DateTime.Now;
                                        while (browseWindow == null &&
                                               (DateTime.Now - startTime).TotalSeconds < TimeoutSeconds)
                                        {
                                            Thread.Sleep(1000); // Проверяем каждую секунду
                                            browseWindow =
                                                desktop
                                                    .FindFirstChild(
                                                        cf => cf.ByName("Browse for Files or Folders")
                                                                  .Or(cf.ByControlType(
                                                                      FlaUI.Core.Definitions.ControlType.Window)))
                                                    ?.AsWindow();
                                        }

                                        if (browseWindow == null)
                                        {
                                            log.Info(
                                                $"Попытка {attempt}: Окно 'Browse for Files or Folders' не найдено");
                                            if (attempt == MaxAttempts)
                                            {
                                                throw new Exception("Окно сохранения не найдено после всех попыток.");
                                            }
                                            continue;
                                        }

                                        log.Info("Окно 'Browse for Files or Folders' найдено");

                                        // Ищем поле ввода "Folder:"
                                        var folderEdit =
                                            browseWindow
                                                .FindFirstDescendant(
                                                    cf => cf.ByControlType(FlaUI.Core.Definitions.ControlType.Edit)
                                                              .And(cf.ByName("Folder:")))
                                                ?.AsTextBox();

                                        if (folderEdit == null)
                                        {
                                            log.Info($"Попытка {attempt}: Поле ввода 'Folder:' не найдено");
                                            if (attempt == MaxAttempts)
                                            {
                                                throw new Exception(
                                                    "Поле ввода пути сохранения не найдено после всех попыток.");
                                            }
                                            continue;
                                        }

                                        log.Info("Поле ввода 'Folder:' найдено");

                                        // Очищаем поле и вводим путь
                                        folderEdit.Text = "";       // Очистка
                                        folderEdit.Text = fullPath; // Предполагается, что fullPath определен ранее
                                        log.Info($"Установлен путь сохранения: {fullPath}");

                                        // Ищем кнопку "OK"
                                        var okButton =
                                            browseWindow
                                                .FindFirstDescendant(cf => cf.ByName("OK").And(cf.ByControlType(
                                                                         FlaUI.Core.Definitions.ControlType.Button)))
                                                ?.AsButton();

                                        if (okButton == null)
                                        {
                                            log.Info($"Попытка {attempt}: Кнопка 'OK' не найдена");
                                            if (attempt == MaxAttempts)
                                            {
                                                throw new Exception("Кнопка 'OK' не найдена после всех попыток.");
                                            }
                                            continue;
                                        }

                                        // Нажимаем кнопку "OK"
                                        okButton.Click();
                                        log.Info("Акт(-ы) сверки успешно сохранены.");
                                        break; // Выходим из цикла после успешного сохранения
                                    }
                                    catch (Exception ex)
                                    {
                                        log.Error($"Попытка {attempt}: Ошибка при сохранении файла - {ex.Message}");
                                        if (attempt == MaxAttempts)
                                        {
                                            throw new Exception("Не удалось сохранить файл после всех попыток.", ex);
                                        }
                                        Thread.Sleep(1000); // Пауза перед следующей попыткой
                                    }
                                }
                            }
                        }
                        else
                        {
                            throw new Exception(
                                "Не удалось привести shell к типу GuiGridView для установки текущей ячейки.");
                        }

                        sapfir.CloseSAPWindowByNex(session);
                    }
                    catch (Exception ex)
                    {

                        Console.WriteLine($"Ошибка SAP: {ex.Message}");
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    log.Error($"Ошибка при обработке входного файла {fileName}. Переход к следующему файлу.");
                    jsonManager.SetValue("status", "ERROR");
                    jsonManager.SetValue("message", ex.Message);
                    jsonManager.SaveToFile(file);
                    continue;
                }
            }

            log.Info("Файлов для обработки нет. Робот завершает свою работу.");
        }
        catch (Exception ex)
        {
            log.Fatal($"Ошибка при инициализации программы: {ex.Message}");
            throw;
        }
    }

    static string GetValueByName(string name)
    {
        return beCodes != null && beCodes.TryGetValue(name, out string value) ? value : null;
    }

    static string CreateRequestFolder(string UUIDRequest, string pathRootFolder, string pathJsonFileRequest)
    {
        // Проверяем существование корневой папки (исправлена логика условия)
        if (!Directory.Exists(pathRootFolder))
        {
            log.Error($"Папка {pathRootFolder} отсутствует, либо нет доступа к ней.");
            throw new Exception();
        }

        // Проверяем существование исходного JSON файла (исправлена логика условия)
        if (!File.Exists(pathJsonFileRequest))
        {
            log.Error($"Файл {pathJsonFileRequest} отсутствует, либо нет доступа к нему.");
            throw new Exception();
        }

        // Формируем полный путь до папки запроса
        string fullPathFolderRequest = Path.Combine(pathRootFolder, UUIDRequest);

        // Определяем подпапки
        string[] _subdirectories = { "ЭДО", "НЕ ЭДО" };

        // Создаем папку запроса, если она не существует
        if (!Directory.Exists(fullPathFolderRequest))
        {
            Directory.CreateDirectory(fullPathFolderRequest);

            // Создаем подпапки
            foreach (var subdirectory in _subdirectories)
            {
                string folder = Path.Combine(fullPathFolderRequest, subdirectory);
                if (!Directory.Exists(folder))
                {
                    Directory.CreateDirectory(folder);
                }
            }
        }

        // Формируем конечный путь для перемещения файла
        string fileName = Path.GetFileName(pathJsonFileRequest);
        string destinationPath = Path.Combine(fullPathFolderRequest, fileName);

        // Перемещаем файл
        // File.Move(pathJsonFileRequest, destinationPath);

        // Возвращаем путь до папки запроса
        return fullPathFolderRequest;
    }
}
}