using FlaUI.UIA3;
using SAPFEWSELib;
using SapROTWr;
using System;
using System.Diagnostics;
using System.Threading;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using NLog;

namespace SpravkoBot_AsSapfir
{
public class SapfirManager
{
    private static readonly Logger log = LogManager.GetCurrentClassLogger();
    public GuiApplication SapApp { get; private set; }
    public GuiConnection SapConnection { get; private set; }
    public GuiSession SapSession { get; private set; }

    private readonly string _pathSAPLoginApp;

    public SapfirManager(string pathSAPLoginApp)
    {
        _pathSAPLoginApp = pathSAPLoginApp ?? throw new ArgumentNullException(nameof(pathSAPLoginApp),
                                                                              "Путь к SAP Logon не может быть null.");
    }

    private T FindElement<T>(string id, bool raiseExceptionIfNotFound = true, int timeoutMs = 5000,
                              int retryIntervalMs = 500) where T : class
    {
        return FindElement<T>(SapSession, id, raiseExceptionIfNotFound, timeoutMs, retryIntervalMs);
    }

    private T FindElement<T>(GuiSession session, string id, bool raiseExceptionIfNotFound = true, int timeoutMs = 5000,
                              int retryIntervalMs = 500) where T : class
    {
        if (session == null)
        {
            if (raiseExceptionIfNotFound)
                throw new Exception("Сессия SAP не инициализирована.");
            return null;
        }

        var start = DateTime.Now;
        while ((DateTime.Now - start).TotalMilliseconds < timeoutMs)
        {
            try
            {
                var component = session.FindById(id, false) as T;
                if (component != null)
                    return component;
            }
            catch
            {
                // игнорируем и повторяем попытку
            }
            Thread.Sleep(retryIntervalMs);
        }

        if (raiseExceptionIfNotFound)
            throw new Exception($"Компонент с ID {id} не найден за отведенное время.");

        return null;
    }

    public void LaunchSAP()
    {
        using (var automation = new UIA3Automation())
        {
            /*                // Если процесс уже запущен, просто выходим
                            if (Process.GetProcessesByName("saplogon").Length > 0)
                            {
                                return;
                            }*/

            // Запуск через FlaUI
            var app = Application.Launch(_pathSAPLoginApp);

            try
            {
                var desktop = automation.GetDesktop();
                Window mainWindow = null;
                var startTime = DateTime.Now;

                // Ожидание окна до 20 секунд
                while (mainWindow == null && (DateTime.Now - startTime).TotalSeconds < 20)
                {
                    mainWindow =
                        desktop
                            .FindFirstChild(cf => cf.ByName("SAP Logon 750")
                                                      .And(cf.ByControlType(FlaUI.Core.Definitions.ControlType.Window)))
                            ?.AsWindow();

                    if (mainWindow == null)
                    {
                        Thread.Sleep(1000); // Пауза перед следующей попыткой
                    }
                }

                if (mainWindow == null)
                {
                    throw new TimeoutException("Не удалось запустить SAP Logon в течение заданного времени.");
                }

                // Дополнительное ожидание для стабилизации
                Thread.Sleep(2000);
            }
            catch (Exception)
            {
                app.Close(); // Закрываем приложение при ошибке
                throw;
            }
        }
    }

    public void KillSAP()
    {
        foreach (var process in Process.GetProcessesByName("saplogon"))
        {
            process.Kill();
            process.WaitForExit(5000); // Ожидание завершения процесса до 5 секунд
        }
    }

    public GuiSession GetSapSession(string sapSystemName)
    {
        if (string.IsNullOrEmpty(sapSystemName))
            throw new ArgumentNullException(nameof(sapSystemName), "Название системы SAP не может быть пустым.");

        CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
        object sapGuiRot = sapROTWrapper.GetROTEntry("SAPGUI");
        if (sapGuiRot == null)
            throw new Exception("Не удалось получить ROT Entry для SAPGUI.");

        object engine = sapGuiRot.GetType().InvokeMember(
            "GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, sapGuiRot, null);
        if (engine == null)
            throw new Exception("Не удалось получить Scripting Engine.");

        SapApp =
            engine as GuiApplication ?? throw new Exception("Не удалось привести Scripting Engine к GuiApplication.");
        SapConnection =
            SapApp.OpenConnection(sapSystemName, false, true); // false - не асинхронно, true - ждать завершения
        SapSession =
            SapConnection.Children.ElementAt(0) as GuiSession ?? throw new Exception("Не удалось получить сессию.");

        return SapSession;
    }

    public GuiSession TryGetCurrentWindowSession()
    {
        try
        {
            CSapROTWrapper sapROTWrapper = new CSapROTWrapper();
            object sapGuiRot = sapROTWrapper.GetROTEntry("SAPGUI");
            if (sapGuiRot == null)
                return null;

            object engine = sapGuiRot.GetType().InvokeMember(
                "GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, sapGuiRot, null);
            if (engine == null)
                return null;

            GuiApplication sapApp = engine as GuiApplication;
            if (sapApp?.Connections.Count == 0)
                return null;

            GuiConnection connection = sapApp.Connections.ElementAt(0) as GuiConnection;
            if (connection?.Children.Count == 0)
                return null;

            return connection.Children.ElementAt(0) as GuiSession;
        }
        catch (Exception ex)
        {
            log.Error(ex, "Ошибка при получении текущей сессии");
            return null;
        }
    }

    public void LoginToSAP(GuiSession session, string user, string password)
    {
        if (session == null)
            throw new ArgumentNullException(nameof(session), "Сессия не может быть null.");
        if (string.IsNullOrEmpty(user))
            throw new ArgumentNullException(nameof(user), "Имя пользователя не может быть пустым.");
        if (string.IsNullOrEmpty(password))
            throw new ArgumentNullException(nameof(password), "Пароль не может быть пустым.");

        Thread.Sleep(1000); // Ожидание для стабилизации интерфейса
        GuiMainWindow mainWindow =
            FindElement<GuiMainWindow>(session, "wnd[0]") ?? throw new Exception("Главное окно wnd[0] не найдено.");
        GuiTextField userField = FindElement<GuiTextField>(session, "wnd[0]/usr/txtRSYST-BNAME") ??
                                 throw new Exception("Поле логина не найдено.");
        GuiTextField pwdField = FindElement<GuiTextField>(session, "wnd[0]/usr/pwdRSYST-BCODE") ??
                                throw new Exception("Поле пароля не найдено.");

        userField.Text = user;
        pwdField.Text = password;
        mainWindow.Maximize();
        mainWindow.SendVKey(0); // Enter
    }

    public void CloseSAPWindowByNex(GuiSession session)
    {
        if (session == null)
        {
            throw new ArgumentNullException(nameof(session), "Сессия не может быть null.");
        }

        try
        {
            var mainWindow = FindElement<GuiMainWindow>(session, "wnd[0]") ??
                             throw new InvalidOperationException("Главное окно 'wnd[0]' не найдено.");

            // Выполняем команду /nex напрямую через HardCopy
            session.SendCommand("/nex");
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Ошибка при закрытии окна SAP: {ex.Message}", ex);
        }
    }

    public void ToolBarButton(string id)
    {
        GuiToolbarControl toolbar = FindElement<GuiToolbarControl>(id) ??
                                    throw new Exception($"Кнопка панели инструментов с ID {id} не найдена.");
        toolbar.SetFocus();
        toolbar.PressButton(id); // Исправлено на PressButton вместо SelectMenuItem
    }

    public void SendKey(string id, int numberKey)
    {
        GuiFrameWindow window =
            FindElement<GuiFrameWindow>(id) ?? throw new Exception($"Окно с ID {id} не найдено.");
        window.SetFocus();
        window.SendVKey(numberKey);
    }

    public void SetText(string id, string text)
    {
        GuiComponent component =
            FindElement<GuiComponent>(id) ?? throw new Exception($"Компонент с ID {id} не найден.");
        if (component.Type == "GuiTextField" || component.Type == "GuiCTextField")
        {
            dynamic textField = component; // Динамическое приведение для поддержки обоих типов
            textField.SetFocus();
            textField.Text = text ?? string.Empty;
        }
        else
        {
            throw new Exception($"Компонент с ID {id} не является текстовым полем (тип: {component.Type}).");
        }
    }

    public void GuiShellPress(string id, string context, bool raiseExceptionIfNotFound = true)
    {
        GuiShell shell = FindElement<GuiShell>(id, raiseExceptionIfNotFound);
        if (shell != null && shell.Type == "GuiGridView")
        {
            GuiGridView grid = shell as GuiGridView;
            grid.PressToolbarContextButton(context);
        }
        else if (raiseExceptionIfNotFound)
        {
            throw new Exception($"GuiShell с ID {id} не найден или не является GuiGridView.");
        }
    }

    public void PressButton(string id)
    {
        GuiButton button = FindElement<GuiButton>(id) ?? throw new Exception($"Кнопка с ID {id} не найдена.");
        button.SetFocus();
        button.Press();
    }

    public void ContextMenu(string id, string key)
    {
        GuiGridView grid =
            FindElement<GuiGridView>(id) ?? throw new Exception($"GridView с ID {id} не найден.");
        grid.ContextMenu();
        grid.SelectContextMenuItem(key);
    }

    public bool GuiCheckBox(string id, bool status)
    {
        GuiCheckBox checkbox =
            FindElement<GuiCheckBox>(id) ?? throw new Exception($"Чекбокс с ID {id} не найден.");
        checkbox.SetFocus();
        checkbox.Selected = status;
        return true;
    }

    public bool RadioButton(string id, bool status)
    {
        GuiRadioButton radio =
            FindElement<GuiRadioButton>(id) ?? throw new Exception($"Радиокнопка с ID {id} не найдена.");
        radio.SetFocus();
        radio.Selected = status;
        return true;
    }

    public string GetLabel(string id)
    {
        GuiLabel label = FindElement<GuiLabel>(id) ?? throw new Exception($"Метка с ID {id} не найдена.");
        label.SetFocus();
        return label.Text ?? string.Empty;
    }

    public void SelectTab(string id)
    {
        GuiTab tab = FindElement<GuiTab>(id) ?? throw new Exception($"Вкладка с ID {id} не найдена.");
        tab.SetFocus();
        tab.Select();
    }

    public string ShowContextText(string id)
    {
        GuiCTextField textField = FindElement<GuiCTextField>(id) ??
                                  throw new Exception($"Поле с ID {id} не найдено или не является GuiCTextField.");
        textField.SetFocus();
        string text = textField.Text;
        textField.ShowContextMenu();
        return text;
    }

    public string GetComponentType(string id)
    {
        GuiComponent component = FindElement<GuiComponent>(id, raiseExceptionIfNotFound: false);
        return component?.Type ?? string.Empty;
    }

    public void SetComboBox(string id, string key)
    {
        GuiComboBox combo =
            FindElement<GuiComboBox>(id) ?? throw new Exception($"Комбобокс с ID {id} не найден.");
        combo.SetFocus();
        combo.Key = key; // Используем string вместо int для большей гибкости
    }

    public void CloseFrameWindow(string id)
    {
        GuiFrameWindow window =
            FindElement<GuiFrameWindow>(id) ?? throw new Exception($"Окно с ID {id} не найдено.");
        window.SetFocus();
        window.Close();
    }

    public string GetFrameText()
    {
        GuiFrameWindow window = SapSession.ActiveWindow as GuiFrameWindow;
        return window?.Text ?? string.Empty;
    }

    public void CloseFrame()
    {
        GuiFrameWindow window = SapSession.ActiveWindow as GuiFrameWindow;
        window?.Close();
    }

    public void SetVerticalScrollPosition(string tableId, int position)
    {
        GuiTableControl table = FindElement<GuiTableControl>(tableId) ??
                                throw new Exception($"Таблица с ID {tableId} не найдена.");
        table.VerticalScrollbar.Position = position;
    }

    public string GetStatusMessage()
    {
        GuiStatusbar statusbar = FindElement<GuiStatusbar>("wnd[0]/sbar", raiseExceptionIfNotFound: false);
        return statusbar?.Text ?? "Статус-бар не найден.";
    }

    public string CheckPopupMessage()
    {
        if (SapSession.Children.Count > 1)
        {
            GuiModalWindow popup = FindElement<GuiModalWindow>("wnd[1]", raiseExceptionIfNotFound: false);
            return popup?.Text ?? "Всплывающее окно не содержит текста.";
        }
        return "Нет всплывающих окон.";
    }

    public int GetWindowIndexByTitle(string windowTitle)
    {
        if (string.IsNullOrEmpty(windowTitle))
            throw new ArgumentNullException(nameof(windowTitle), "Заголовок окна не может быть пустым.");

        for (int i = 0; i < SapSession.Children.Count; i++)
        {
            GuiMainWindow window = FindElement<GuiMainWindow>($"wnd[{i}]", raiseExceptionIfNotFound: false);
            if (window != null && window.Text.Contains(windowTitle))
            {
                return i;
            }
        }
        return -1;
    }

    public string GetTextFromField(string id)
    {
        GuiComponent component =
            FindElement<GuiComponent>(id) ?? throw new Exception($"Компонент с ID {id} не найден.");
        if (component.Type == "GuiCTextField" || component.Type == "GuiTextField")
        {
            dynamic textField = component;
            return textField.Text ?? string.Empty;
        }
        throw new Exception($"Компонент с ID {id} не является текстовым полем (тип: {component.Type}).");
    }
}
}