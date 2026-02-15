using System.Globalization;
using ClosedXML.Excel;
using Spectre.Console;
using System.Text;

namespace PortfolioCli;

internal static class Program
{
    private const string DefaultWorkbook = "my_portfolio.xlsx";
    private const string SheetPrefix = "Data Over time ";
    private const int RecentDayRows = 25;
    private static readonly bool UseUnicodeSymbols = DetermineUnicodeSupport();
    private static readonly string CurrencyPrefix = UseUnicodeSymbols ? "£" : "GBP ";
    private static readonly char BarChar = UseUnicodeSymbols ? '█' : '#';

    public static int Main(string[] args)
    {
        if (UseUnicodeSymbols)
        {
            Console.OutputEncoding = Encoding.UTF8;
        }

        Dictionary<string, string> options;
        string workbookPath;

        if (args.Length == 0)
        {
            options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            workbookPath = DefaultWorkbook;
            return RunInteractive(workbookPath, options);
        }

        if (IsHelp(args[0]))
        {
            PrintHelp();
            return 0;
        }

        var command = args[0].Trim().ToLowerInvariant();
        options = command.StartsWith("--", StringComparison.Ordinal)
            ? ParseOptions(args)
            : ParseOptions(args.Skip(1).ToArray());

        workbookPath = options.TryGetValue("--file", out var file) ? file : DefaultWorkbook;

        try
        {
            if (command.StartsWith("--", StringComparison.Ordinal))
            {
                return RunInteractive(workbookPath, options);
            }

            return command switch
            {
                "view" => RunView(workbookPath, options),
                "add" => RunAdd(workbookPath, options),
                "interactive" => RunInteractive(workbookPath, options),
                _ => PrintUnknown(command)
            };
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
            return 1;
        }
    }

    private static int PrintUnknown(string command)
    {
        AnsiConsole.MarkupLine($"[red]Unknown command:[/] {Markup.Escape(command)}");
        PrintHelp();
        return 1;
    }

    private static void PrintHelp()
    {
        var text = new Panel(new Markup(
            "[bold]my-portfolio-cli[/]\n" +
            "Usage:\n" +
            "  [grey]dotnet run --project PortfolioCli -- view [--month 2026-02] [--date 2026-02-15] [--file my_portfolio.xlsx][/]\n" +
            "  [grey]dotnet run --project PortfolioCli -- add  [--date 2026-02-15] [--file my_portfolio.xlsx][/]\n\n" +
            "  [grey]dotnet run --project PortfolioCli -- interactive [--month 2026-02] [--date 2026-02-15] [--file my_portfolio.xlsx][/]\n\n" +
            "Commands (optional):\n" +
            "  [bold]view[/]  Show the latest snapshot (or a specific month/date)\n" +
            "  [bold]add[/]   Add a daily snapshot (prompts for account values)\n" +
            "  [bold]interactive[/]  Interactive dashboard (arrows to navigate)\n\n" +
            "If no command is provided, interactive mode starts by default.\n"))
        {
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 1, 1, 1)
        };
        AnsiConsole.Write(text);
    }

    private static bool IsHelp(string arg)
        => arg.Equals("-h", StringComparison.OrdinalIgnoreCase)
           || arg.Equals("--help", StringComparison.OrdinalIgnoreCase)
           || arg.Equals("help", StringComparison.OrdinalIgnoreCase);

    private static Dictionary<string, string> ParseOptions(string[] args)
    {
        var options = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < args.Length; i++)
        {
            var token = args[i];
            if (!token.StartsWith("--", StringComparison.Ordinal))
            {
                continue;
            }

            var value = "true";
            if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
            {
                value = args[i + 1];
                i++;
            }

            options[token] = value;
        }

        return options;
    }

    private static int RunView(string workbookPath, Dictionary<string, string> options)
    {
        if (!File.Exists(workbookPath))
        {
            AnsiConsole.MarkupLine($"[red]Workbook not found:[/] {Markup.Escape(workbookPath)}");
            return 1;
        }

        using var workbook = new XLWorkbook(workbookPath);
        var sheet = SelectSheet(workbook, options);
        if (sheet == null)
        {
            AnsiConsole.MarkupLine("[red]No matching month sheet found.[/]");
            return 1;
        }

        var info = ParseSheet(sheet);
        if (info.AccountRows.Count == 0 || info.DateColumns.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]Sheet does not contain recognizable portfolio data.[/]");
            return 1;
        }

        var dateOverride = options.TryGetValue("--date", out var dateValue)
            && TryParseDate(dateValue, out var parsedDate)
            ? parsedDate.Date
            : (DateTime?)null;

        var current = FindLatestSnapshot(info, dateOverride);
        if (current == null)
        {
            AnsiConsole.MarkupLine("[red]No data found for that date/month.[/]");
            return 1;
        }

        var fySummary = TryGetFySummary(workbook);
        RenderSnapshot(workbookPath, sheet.Name, current, current.Date, null, false, fySummary);
        return 0;
    }

    private static int RunAdd(string workbookPath, Dictionary<string, string> options)
    {
        if (!File.Exists(workbookPath))
        {
            var state = InitializeWorkbookFromScratch(workbookPath, options, null);
            return state == null ? 1 : 0;
        }

        using var workbook = new XLWorkbook(workbookPath);
        if (!HasPortfolioSheets(workbook))
        {
            var state = InitializeWorkbookFromScratch(workbookPath, options, workbook);
            return state == null ? 1 : 0;
        }

        DateTime date;
        if (options.TryGetValue("--date", out var dateInput))
        {
            if (!TryParseDate(dateInput, out date))
            {
                AnsiConsole.MarkupLine("[red]Invalid date. Use yyyy-MM-dd.[/]");
                return 1;
            }
        }
        else
        {
            date = PromptForDate();
        }

        var sheet = EnsureMonthSheet(workbook, date);
        var info = ParseSheet(sheet);
        if (info.AccountRows.Count == 0 || info.DateColumns.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]Sheet does not contain recognizable portfolio data.[/]");
            return 1;
        }

        var targetColumn = FindDateColumn(info, date);
        if (targetColumn == null)
        {
            AnsiConsole.MarkupLine($"[red]Date column not found in sheet:[/] {date:yyyy-MM-dd}");
            return 1;
        }

        var previousColumn = FindPreviousDateColumn(info, targetColumn.Value);
        var hasExisting = info.AccountRows.Any(row => !sheet.Cell(row.Row, targetColumn.Value).IsEmpty());
        if (hasExisting && !AnsiConsole.Confirm("Values already exist for this date. Overwrite?"))
        {
            return 0;
        }

        foreach (var account in info.AccountRows)
        {
            var prevValue = previousColumn.HasValue
                ? GetDecimalOrZero(sheet.Cell(account.Row, previousColumn.Value))
                : 0m;

            var prompt = new TextPrompt<decimal>($"{account.Name} value")
                .DefaultValue(prevValue)
                .ShowDefaultValue()
                .Validate(value => value >= 0m);

            var value = AnsiConsole.Prompt(prompt);
            sheet.Cell(account.Row, targetColumn.Value).Value = value;
        }

        workbook.SaveAs(workbookPath);
        AnsiConsole.MarkupLine($"[green]Saved:[/] {Markup.Escape(workbookPath)}");
        return 0;
    }

    private static int RunInteractive(string workbookPath, Dictionary<string, string> options)
    {
        if (!AnsiConsole.Profile.Capabilities.Interactive)
        {
            AnsiConsole.MarkupLine("[red]Interactive mode requires a real console.[/]");
            return 1;
        }

        var state = InitializeInteractiveState(workbookPath, options);
        if (state == null)
        {
            return 1;
        }

        while (true)
        {
            using var workbook = new XLWorkbook(workbookPath);
            var selection = SelectSheetForInteractive(workbook, state.SelectedDate);
            if (selection.Sheet == null)
            {
                AnsiConsole.MarkupLine("[red]No portfolio sheets found.[/]");
                return 1;
            }

            state.IsMonthMatched = selection.MonthMatched;
            var sheet = selection.Sheet;
            var info = ParseSheet(sheet);
            if (info.AccountRows.Count == 0 || info.DateColumns.Count == 0)
            {
                AnsiConsole.MarkupLine("[red]Sheet does not contain recognizable portfolio data.[/]");
                return 1;
            }

            var displayDate = ClampToAvailableDate(info, selection.DisplayDate);
            if (selection.MonthMatched && displayDate != selection.DisplayDate)
            {
                state.SelectedDate = displayDate;
            }

            var selectedColumn = FindDateColumn(info, displayDate) ?? info.DateColumns.Last().Column;

            var snapshot = BuildSnapshot(info, selectedColumn, true, out var hasData);
            var fySummary = TryGetFySummary(workbook);
            var status = state.StatusMessage ?? selection.StatusMessage;
            if (!hasData)
            {
                var noData = $"No entries for {state.SelectedDate:yyyy-MM-dd}. Showing previous values.";
                status = string.IsNullOrWhiteSpace(status) ? noData : $"{status} {noData}";
            }

            RenderSnapshot(workbookPath, sheet.Name, snapshot, state.SelectedDate, status, true, fySummary);
            state.StatusMessage = null;

            var key = Console.ReadKey(true);
            switch (key.Key)
            {
                case ConsoleKey.Q:
                case ConsoleKey.Escape:
                    return 0;
                case ConsoleKey.LeftArrow:
                    state.SelectedDate = state.SelectedDate.AddMonths(-1);
                    break;
                case ConsoleKey.RightArrow:
                    state.SelectedDate = state.SelectedDate.AddMonths(1);
                    break;
                case ConsoleKey.UpArrow:
                    state.SelectedDate = state.IsMonthMatched
                        ? MoveByDay(info, state.SelectedDate, 1)
                        : state.SelectedDate.AddDays(1);
                    break;
                case ConsoleKey.DownArrow:
                    state.SelectedDate = state.IsMonthMatched
                        ? MoveByDay(info, state.SelectedDate, -1)
                        : state.SelectedDate.AddDays(-1);
                    break;
                case ConsoleKey.A:
                    AddValuesInteractive(workbookPath, state.SelectedDate, out var message);
                    if (!string.IsNullOrWhiteSpace(message))
                    {
                        state.StatusMessage = message;
                    }
                    break;
            }
        }
    }

    private static bool AddValuesInteractive(string workbookPath, DateTime date, out string message)
    {
        message = string.Empty;
        using var workbook = new XLWorkbook(workbookPath);
        var sheet = EnsureMonthSheet(workbook, date);
        var info = ParseSheet(sheet);
        if (info.AccountRows.Count == 0 || info.DateColumns.Count == 0)
        {
            message = "Sheet does not contain recognizable portfolio data.";
            return false;
        }

        var targetColumn = FindDateColumn(info, date);
        if (targetColumn == null)
        {
            message = $"Date column not found: {date:yyyy-MM-dd}";
            return false;
        }

        if (AnsiConsole.Profile.Capabilities.Interactive)
        {
            AnsiConsole.Clear();
        }

        var panel = new Panel(new Markup(
            $"[bold]Add Values[/]\n" +
            $"[grey]Sheet:[/] {Markup.Escape(sheet.Name)}\n" +
            $"[grey]Date:[/] {date:yyyy-MM-dd}"))
        {
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0)
        };
        AnsiConsole.Write(panel);
        AnsiConsole.WriteLine();

        var previousColumn = FindPreviousDateColumn(info, targetColumn.Value);
        var hasExisting = info.AccountRows.Any(row => !sheet.Cell(row.Row, targetColumn.Value).IsEmpty());
        if (hasExisting && !AnsiConsole.Confirm("Values already exist for this date. Overwrite?"))
        {
            message = "Add canceled.";
            return false;
        }

        foreach (var account in info.AccountRows)
        {
            var prevValue = previousColumn.HasValue
                ? GetDecimalOrZero(sheet.Cell(account.Row, previousColumn.Value))
                : 0m;

            var prompt = new TextPrompt<decimal>($"{account.Name} value")
                .DefaultValue(prevValue)
                .ShowDefaultValue()
                .Validate(value => value >= 0m);

            var value = AnsiConsole.Prompt(prompt);
            sheet.Cell(account.Row, targetColumn.Value).Value = value;
        }

        workbook.SaveAs(workbookPath);
        message = $"Saved values for {date:yyyy-MM-dd}.";
        return true;
    }

    private static DateTime PromptForDate()
    {
        var today = DateTime.Today;
        while (true)
        {
            var input = AnsiConsole.Prompt(
                new TextPrompt<string>("Date (yyyy-MM-dd)")
                    .DefaultValue(today.ToString("yyyy-MM-dd"))
                    .ShowDefaultValue());

            if (TryParseDate(input, out var date))
            {
                return date.Date;
            }

            AnsiConsole.MarkupLine("[red]Invalid date. Try again.[/]");
        }
    }

    private static UiState? InitializeWorkbookFromScratch(string workbookPath, Dictionary<string, string> options, XLWorkbook? existingWorkbook)
    {
        var shouldExpand = AnsiConsole.Profile.Capabilities.Interactive;
        if (shouldExpand)
        {
            AnsiConsole.Clear();
        }

        AnsiConsole.MarkupLine("[yellow]No portfolio data found. Let's create your first month.[/]");

        DateTime date;
        if (options.TryGetValue("--date", out var dateInput))
        {
            if (!TryParseDate(dateInput, out date))
            {
                AnsiConsole.MarkupLine("[red]Invalid date. Use yyyy-MM-dd.[/]");
                return null;
            }
        }
        else
        {
            date = PromptForDate();
        }

        var accounts = PromptForAccounts(date);
        if (accounts.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]At least one account is required.[/]");
            return null;
        }

        var workbook = existingWorkbook ?? new XLWorkbook();
        var sheetName = MonthSheetName(date);
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
            ?? workbook.AddWorksheet(sheetName);

        BuildSheetFromScratch(sheet, date, accounts);
        workbook.SaveAs(workbookPath);

        if (existingWorkbook == null)
        {
            workbook.Dispose();
        }

        return new UiState(date)
        {
            StatusMessage = $"Created {sheetName} with {accounts.Count} accounts.",
            IsMonthMatched = true
        };
    }

    private static List<AccountSeed> PromptForAccounts(DateTime date)
    {
        var accounts = new List<AccountSeed>();
        while (true)
        {
            var name = AnsiConsole.Prompt(
                new TextPrompt<string>("Account name (blank to finish)")
                    .AllowEmpty());

            if (string.IsNullOrWhiteSpace(name))
            {
                if (accounts.Count == 0)
                {
                    AnsiConsole.MarkupLine("[red]Please add at least one account.[/]");
                    continue;
                }

                break;
            }

            var value = AnsiConsole.Prompt(
                new TextPrompt<decimal>($"{name} value ({date:yyyy-MM-dd})")
                    .Validate(v => v >= 0m));

            accounts.Add(new AccountSeed(name.Trim(), value));
        }

        return accounts;
    }

    private static void BuildSheetFromScratch(IXLWorksheet sheet, DateTime date, IReadOnlyList<AccountSeed> accounts)
    {
        var firstDay = new DateTime(date.Year, date.Month, 1);
        var lastDay = firstDay.AddMonths(1).AddDays(-1);
        var prevDay = firstDay.AddDays(-1);

        sheet.Clear();
        sheet.Cell(1, 2).Value = prevDay;
        var col = 3;
        var current = firstDay;
        while (current <= lastDay)
        {
            sheet.Cell(1, col).Value = current;
            col++;
            current = current.AddDays(1);
        }

        var startRow = 4;
        var totalRow = startRow + accounts.Count;
        var changeRow = totalRow + 1;
        var targetCol = 2 + date.Day;

        for (var i = 0; i < accounts.Count; i++)
        {
            var row = startRow + i;
            var account = accounts[i];
            sheet.Cell(row, 1).Value = account.Name;
            sheet.Cell(row, 2).Value = account.Value;
            sheet.Cell(row, targetCol).Value = account.Value;
        }

        sheet.Cell(totalRow, 1).Value = "Total";
        for (var c = 2; c < col; c++)
        {
            var columnLetter = XLHelper.GetColumnLetterFromNumber(c);
            sheet.Cell(totalRow, c).FormulaA1 = $"SUM({columnLetter}{startRow}:{columnLetter}{totalRow - 1})";
        }

        for (var c = 3; c < col; c++)
        {
            var columnLetter = XLHelper.GetColumnLetterFromNumber(c);
            var prevLetter = XLHelper.GetColumnLetterFromNumber(c - 1);
            sheet.Cell(changeRow, c).FormulaA1 = $"IF({columnLetter}{totalRow}=0,0,{columnLetter}{totalRow}-{prevLetter}{totalRow})";
        }
    }

    private static bool HasPortfolioSheets(XLWorkbook workbook)
        => workbook.Worksheets.Any(ws => TryGetMonthFromSheetName(ws.Name).HasValue);

    private static UiState? InitializeInteractiveState(string workbookPath, Dictionary<string, string> options)
    {
        if (!File.Exists(workbookPath))
        {
            return InitializeWorkbookFromScratch(workbookPath, options, null);
        }

        using var workbook = new XLWorkbook(workbookPath);
        if (!HasPortfolioSheets(workbook))
        {
            return InitializeWorkbookFromScratch(workbookPath, options, workbook);
        }
        DateTime requestedDate;
        string? statusMessage = null;

        if (options.TryGetValue("--date", out var dateValue))
        {
            if (!TryParseDate(dateValue, out requestedDate))
            {
                AnsiConsole.MarkupLine("[red]Invalid date. Use yyyy-MM-dd.[/]");
                return null;
            }
        }
        else if (options.TryGetValue("--month", out var monthValue))
        {
            if (!TryParseMonth(monthValue, out var month))
            {
                AnsiConsole.MarkupLine("[red]Invalid month. Use yyyy-MM.[/]");
                return null;
            }

            requestedDate = month;
        }
        else
        {
            requestedDate = DateTime.Today;
        }

        var selection = SelectSheetForInteractive(workbook, requestedDate);
        if (selection.Sheet == null)
        {
            AnsiConsole.MarkupLine("[red]No portfolio sheets found.[/]");
            return null;
        }

        statusMessage = selection.StatusMessage;
        var infoForMonth = ParseSheet(selection.Sheet);
        var adjustedDate = ClampToAvailableDate(infoForMonth, selection.DisplayDate);
        var initialDate = selection.MonthMatched ? adjustedDate : requestedDate.Date;

        return new UiState(initialDate)
        {
            StatusMessage = statusMessage,
            IsMonthMatched = selection.MonthMatched
        };
    }

    private static bool TryParseDate(string input, out DateTime date)
        => DateTime.TryParseExact(input.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture,
            DateTimeStyles.None, out date);

    private static IXLWorksheet? SelectSheet(XLWorkbook workbook, Dictionary<string, string> options)
    {
        if (options.TryGetValue("--month", out var monthValue) && TryParseMonth(monthValue, out var month))
        {
            return workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, MonthSheetName(month), StringComparison.OrdinalIgnoreCase));
        }

        if (options.TryGetValue("--date", out var dateValue) && TryParseDate(dateValue, out var date))
        {
            var name = MonthSheetName(date);
            return workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        return FindLatestMonthSheet(workbook);
    }

    private static SheetSelection SelectSheetForInteractive(XLWorkbook workbook, DateTime requestedDate)
    {
        var name = MonthSheetName(requestedDate);
        var sheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));
        if (sheet != null)
        {
            return new SheetSelection(sheet, requestedDate.Date, null, true);
        }

        var targetMonth = new DateTime(requestedDate.Year, requestedDate.Month, 1);
        var candidates = workbook.Worksheets
            .Select(ws => (ws, month: TryGetMonthFromSheetName(ws.Name)))
            .Where(x => x.month.HasValue)
            .Select(x => (x.ws, month: x.month!.Value))
            .OrderBy(x => x.month)
            .ToList();

        if (candidates.Count == 0)
        {
            return new SheetSelection(null, requestedDate.Date, null, false);
        }

        var previous = candidates.Where(x => x.month < targetMonth).OrderByDescending(x => x.month).FirstOrDefault();
        var next = candidates.Where(x => x.month > targetMonth).OrderBy(x => x.month).FirstOrDefault();
        var chosen = previous.ws != null ? previous : next;

        if (chosen.ws == null)
        {
            return new SheetSelection(null, requestedDate.Date, null, false);
        }

        var daysInMonth = DateTime.DaysInMonth(chosen.month.Year, chosen.month.Month);
        var adjustedDate = new DateTime(chosen.month.Year, chosen.month.Month, Math.Min(requestedDate.Day, daysInMonth));
        var statusMessage = $"Month {requestedDate:yyyy-MM} not found. Showing {chosen.month:yyyy-MM}. Press A to create it.";
        return new SheetSelection(chosen.ws, adjustedDate, statusMessage, false);
    }

    private static bool TryParseMonth(string input, out DateTime month)
    {
        var formats = new[] { "yyyy-MM", "yyyy-M" };
        foreach (var format in formats)
        {
            if (DateTime.TryParseExact(input.Trim(), format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            {
                month = new DateTime(dt.Year, dt.Month, 1);
                return true;
            }
        }

        month = default;
        return false;
    }

    private static IXLWorksheet? FindLatestMonthSheet(XLWorkbook workbook)
    {
        var candidates = workbook.Worksheets
            .Select(ws => (ws, month: TryGetMonthFromSheetName(ws.Name)))
            .Where(x => x.month.HasValue)
            .OrderBy(x => x.month)
            .ToList();

        return candidates.LastOrDefault().ws;
    }

    private static DateTime ClampToAvailableDate(SheetInfo info, DateTime date)
    {
        var dates = info.DateColumns.Select(dc => dc.Date).OrderBy(d => d).ToList();
        if (dates.Count == 0)
        {
            return date.Date;
        }

        var exactIndex = dates.FindIndex(d => d == date.Date);
        if (exactIndex >= 0)
        {
            return dates[exactIndex];
        }

        var previous = dates.LastOrDefault(d => d < date.Date);
        if (previous != default)
        {
            return previous;
        }

        return dates[0];
    }

    private static DateTime MoveByDay(SheetInfo info, DateTime current, int delta)
    {
        var dates = info.DateColumns.Select(dc => dc.Date).OrderBy(d => d).ToList();
        if (dates.Count == 0)
        {
            return current.Date;
        }

        var index = dates.FindIndex(d => d == current.Date);
        if (index < 0)
        {
            index = dates.FindLastIndex(d => d < current.Date);
            if (index < 0)
            {
                index = 0;
            }
        }

        var nextIndex = Math.Clamp(index + delta, 0, dates.Count - 1);
        return dates[nextIndex];
    }

    private static DateTime? TryGetMonthFromSheetName(string name)
    {
        if (!name.StartsWith(SheetPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var monthPart = name.Substring(SheetPrefix.Length).Trim();
        if (DateTime.TryParseExact(monthPart, "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
        {
            return new DateTime(dt.Year, dt.Month, 1);
        }

        return null;
    }

    private static string MonthSheetName(DateTime date)
        => $"{SheetPrefix}{date:MMMM yyyy}";

    private static IXLWorksheet EnsureMonthSheet(XLWorkbook workbook, DateTime date)
    {
        var name = MonthSheetName(date);
        var existing = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            return existing;
        }

        var previousSheet = FindPreviousMonthSheet(workbook, date);
        if (previousSheet == null)
        {
            throw new InvalidOperationException("No previous month sheet found to copy from.");
        }

        var newSheet = previousSheet.CopyTo(name);
        InitializeNewMonthSheet(newSheet, previousSheet, date);
        return newSheet;
    }

    private static IXLWorksheet? FindPreviousMonthSheet(XLWorkbook workbook, DateTime date)
    {
        var targetMonth = new DateTime(date.Year, date.Month, 1);
        var candidates = workbook.Worksheets
            .Select(ws => (ws, month: TryGetMonthFromSheetName(ws.Name)))
            .Where(x => x.month.HasValue && x.month.Value < targetMonth)
            .OrderBy(x => x.month)
            .ToList();

        return candidates.LastOrDefault().ws;
    }

    private static void InitializeNewMonthSheet(IXLWorksheet sheet, IXLWorksheet previousSheet, DateTime date)
    {
        var targetMonth = new DateTime(date.Year, date.Month, 1);
        var previousMonthEnd = targetMonth.AddDays(-1);
        var lastDay = targetMonth.AddMonths(1).AddDays(-1);

        var previousInfo = ParseSheet(previousSheet);
        if (previousInfo.DateColumns.Count == 0)
        {
            throw new InvalidOperationException("Previous sheet does not contain date columns.");
        }

        var prevLastCol = previousInfo.DateColumns.Last().Column;
        var prevDataCol = FindLatestDataColumn(previousInfo) ?? prevLastCol;
        var lastUsedCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 2;

        // Set header dates
        var col = 2; // column B
        sheet.Cell(1, col).Value = previousMonthEnd;
        var currentDate = targetMonth;
        while (currentDate <= lastDay)
        {
            col++;
            sheet.Cell(1, col).Value = currentDate;
            currentDate = currentDate.AddDays(1);
        }

        var lastDateCol = col;
        for (var c = lastDateCol + 1; c <= lastUsedCol; c++)
        {
            sheet.Cell(1, c).Clear();
        }

        // Carry baseline values from previous month end, clear new month columns.
        var totalRow = FindTotalRow(sheet);
        for (var r = 2; r < totalRow; r++)
        {
            sheet.Cell(r, 2).Value = previousSheet.Cell(r, prevDataCol).Value;
            for (var c = 3; c <= lastUsedCol; c++)
            {
                sheet.Cell(r, c).Clear();
            }
        }
    }

    private static SheetInfo ParseSheet(IXLWorksheet sheet)
    {
        var dateColumns = new List<DateColumn>();
        var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        for (var col = 2; col <= lastCol; col++)
        {
            var cell = sheet.Cell(1, col);
            if (cell.IsEmpty())
            {
                continue;
            }

            if (cell.TryGetValue<DateTime>(out var date))
            {
                dateColumns.Add(new DateColumn(date.Date, col));
            }
            else if (cell.TryGetValue<double>(out var oaDate) && oaDate is > 20000 and < 60000)
            {
                dateColumns.Add(new DateColumn(DateTime.FromOADate(oaDate).Date, col));
            }
            else if (DateTime.TryParse(cell.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            {
                dateColumns.Add(new DateColumn(parsed.Date, col));
            }
        }

        dateColumns = NormalizeDateColumns(dateColumns);

        var totalRow = FindTotalRow(sheet);
        var accountRows = new List<AccountRow>();
        var firstDateCol = dateColumns.Count > 0 ? dateColumns.OrderBy(dc => dc.Column).First().Column : 2;

        for (var row = 4; row < totalRow; row++)
        {
            var name = sheet.Cell(row, 1).GetString().Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var valueCell = sheet.Cell(row, firstDateCol);
            if (valueCell.IsEmpty())
            {
                continue;
            }

            if (valueCell.TryGetValue<double>(out var numeric) && numeric >= 1d)
            {
                accountRows.Add(new AccountRow(name, row));
            }
        }

        return new SheetInfo(sheet.Name, dateColumns, accountRows, totalRow, sheet);
    }

    private static List<DateColumn> NormalizeDateColumns(List<DateColumn> dateColumns)
    {
        var ordered = dateColumns.OrderBy(dc => dc.Column).ToList();
        if (ordered.Count == 0)
        {
            return ordered;
        }

        var contiguous = new List<DateColumn> { ordered[0] };
        for (var i = 1; i < ordered.Count; i++)
        {
            var previous = contiguous[^1];
            var candidate = ordered[i];
            if (candidate.Column != previous.Column + 1)
            {
                break;
            }

            if (candidate.Date != previous.Date.AddDays(1))
            {
                break;
            }

            contiguous.Add(candidate);
        }

        return contiguous;
    }

    private static int FindTotalRow(IXLWorksheet sheet)
    {
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 1;
        for (var row = 1; row <= lastRow; row++)
        {
            var name = sheet.Cell(row, 1).GetString().Trim();
            if (name.Equals("Total", StringComparison.OrdinalIgnoreCase))
            {
                return row;
            }
        }

        throw new InvalidOperationException("Could not locate Total row in the sheet.");
    }

    private static int? FindDateColumn(SheetInfo info, DateTime date)
    {
        var match = info.DateColumns.FirstOrDefault(dc => dc.Date == date.Date);
        return match == null ? null : match.Column;
    }

    private static int? FindPreviousDateColumn(SheetInfo info, int targetColumn)
    {
        var previous = info.DateColumns
            .Where(dc => dc.Column < targetColumn)
            .OrderByDescending(dc => dc.Column)
            .FirstOrDefault();

        return previous == null ? null : previous.Column;
    }

    private static Snapshot? FindLatestSnapshot(SheetInfo info, DateTime? overrideDate)
    {
        var dateColumns = info.DateColumns;
        if (overrideDate.HasValue)
        {
            var overrideCol = FindDateColumn(info, overrideDate.Value);
            if (!overrideCol.HasValue)
            {
                return null;
            }

            return BuildSnapshot(info, overrideCol.Value, false, out _);
        }

        for (var i = dateColumns.Count - 1; i >= 0; i--)
        {
            var column = dateColumns[i].Column;
            if (info.AccountRows.Any(row => !info.Worksheet.Cell(row.Row, column).IsEmpty()))
            {
                return BuildSnapshot(info, column, false, out _);
            }
        }

        return null;
    }

    private static Snapshot BuildSnapshot(SheetInfo info, int column, bool carryForwardIfEmpty, out bool hasData)
    {
        hasData = ColumnHasData(info, column);
        var previousColumn = FindPreviousDateColumn(info, column);
        var effectiveColumn = hasData || !carryForwardIfEmpty
            ? column
            : FindPreviousDataColumn(info, column) ?? column;

        var accountSnapshots = new List<AccountSnapshot>();

        foreach (var account in info.AccountRows)
        {
            var current = GetDecimalOrZero(info.Worksheet.Cell(account.Row, effectiveColumn));
            var previous = previousColumn.HasValue
                ? GetDecimalOrZero(info.Worksheet.Cell(account.Row, previousColumn.Value))
                : 0m;

            if (!hasData && carryForwardIfEmpty)
            {
                previous = current;
            }

            var change = current - previous;
            var changePct = previous != 0m ? change / previous : 0m;
            accountSnapshots.Add(new AccountSnapshot(account.Name, current, previous, change, changePct));
        }

        var currentTotal = accountSnapshots.Sum(a => a.Current);
        var previousTotal = accountSnapshots.Sum(a => a.Previous);
        var totalChange = currentTotal - previousTotal;
        var totalChangePct = previousTotal != 0m ? totalChange / previousTotal : 0m;

        var baselineColumn = info.DateColumns.OrderBy(dc => dc.Column).First().Column;
        var baselineTotal = info.AccountRows.Sum(row => GetDecimalOrZero(info.Worksheet.Cell(row.Row, baselineColumn)));
        var monthToDate = currentTotal - baselineTotal;
        var monthToDatePct = baselineTotal != 0m ? monthToDate / baselineTotal : 0m;

        var recentChanges = new List<DailyChange>();
        var recentDates = info.DateColumns
            .Where(dc => dc.Column <= column)
            .OrderByDescending(dc => dc.Column)
            .Take(RecentDayRows + 1)
            .OrderBy(dc => dc.Column)
            .ToList();

        for (var i = 1; i < recentDates.Count; i++)
        {
            var currentCol = recentDates[i].Column;
            var previousCol = recentDates[i - 1].Column;
            var dayTotal = GetEffectiveTotal(info, currentCol, carryForwardIfEmpty);
            var prevDayTotal = GetEffectiveTotal(info, previousCol, carryForwardIfEmpty);
            var diff = dayTotal - prevDayTotal;
            var diffPct = prevDayTotal != 0m ? diff / prevDayTotal : 0m;
            recentChanges.Add(new DailyChange(recentDates[i].Date, diff, diffPct));
        }

        var date = info.DateColumns.First(dc => dc.Column == column).Date;
        return new Snapshot(date, accountSnapshots, currentTotal, totalChange, totalChangePct, monthToDate, monthToDatePct, recentChanges, info.Worksheet.Name);
    }

    private static bool ColumnHasData(SheetInfo info, int column)
        => info.AccountRows.Any(row => !info.Worksheet.Cell(row.Row, column).IsEmpty());

    private static int? FindLatestDataColumn(SheetInfo info)
    {
        for (var i = info.DateColumns.Count - 1; i >= 0; i--)
        {
            var column = info.DateColumns[i].Column;
            if (ColumnHasData(info, column))
            {
                return column;
            }
        }

        return null;
    }

    private static int? FindPreviousDataColumn(SheetInfo info, int column)
    {
        var previous = FindPreviousDateColumn(info, column);
        while (previous.HasValue)
        {
            if (ColumnHasData(info, previous.Value))
            {
                return previous;
            }

            previous = FindPreviousDateColumn(info, previous.Value);
        }

        return null;
    }

    private static decimal GetEffectiveTotal(SheetInfo info, int column, bool carryForwardIfEmpty)
    {
        if (!carryForwardIfEmpty || ColumnHasData(info, column))
        {
            return info.AccountRows.Sum(row => GetDecimalOrZero(info.Worksheet.Cell(row.Row, column)));
        }

        var previous = FindPreviousDataColumn(info, column);
        if (!previous.HasValue)
        {
            return 0m;
        }

        return info.AccountRows.Sum(row => GetDecimalOrZero(info.Worksheet.Cell(row.Row, previous.Value)));
    }

    private static FySummary? TryGetFySummary(XLWorkbook workbook)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals("Dashboard", StringComparison.OrdinalIgnoreCase));
        if (sheet == null)
        {
            return null;
        }

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var lastCol = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        if (lastRow == 0 || lastCol == 0)
        {
            return null;
        }

        var (headerRow, monthColumns, totalColumn) = FindMonthHeader(sheet, lastRow, lastCol);
        if (headerRow == 0 || monthColumns.Count == 0)
        {
            return null;
        }

        var returnRow = FindRowByLabel(sheet, lastRow, lastCol, "Return");
        var cashRow = FindRowByLabel(sheet, lastRow, lastCol, "Cash");
        var pnlRow = FindRowByLabel(sheet, lastRow, lastCol, "PnL");

        if (returnRow == 0 && cashRow == 0 && pnlRow == 0)
        {
            return null;
        }

        var months = new List<FyMonth>();
        foreach (var column in monthColumns.OrderBy(mc => mc.Column))
        {
            var ret = returnRow == 0 ? null : GetDecimalOrNull(sheet.Cell(returnRow, column.Column));
            var cash = cashRow == 0 ? null : GetDecimalOrNull(sheet.Cell(cashRow, column.Column));
            var pnl = pnlRow == 0 ? null : GetDecimalOrNull(sheet.Cell(pnlRow, column.Column));
            months.Add(new FyMonth(column.Label, column.MonthNumber, ret, cash, pnl));
        }

        decimal? totalReturn = null;
        decimal? totalCash = null;
        decimal? totalPnL = null;
        if (totalColumn.HasValue)
        {
            if (returnRow != 0)
            {
                totalReturn = GetDecimalOrNull(sheet.Cell(returnRow, totalColumn.Value));
            }

            if (cashRow != 0)
            {
                totalCash = GetDecimalOrNull(sheet.Cell(cashRow, totalColumn.Value));
            }

            if (pnlRow != 0)
            {
                totalPnL = GetDecimalOrNull(sheet.Cell(pnlRow, totalColumn.Value));
            }
        }

        var title = FindTitle(sheet, lastRow, lastCol) ?? "Financial Year";
        return new FySummary(title, months, totalReturn, totalCash, totalPnL);
    }

    private static (int HeaderRow, List<MonthColumn> MonthColumns, int? TotalColumn) FindMonthHeader(IXLWorksheet sheet, int lastRow, int lastCol)
    {
        for (var row = 1; row <= Math.Min(lastRow, 25); row++)
        {
            var monthColumns = new List<MonthColumn>();
            int? totalColumn = null;
            for (var col = 1; col <= lastCol; col++)
            {
                var text = sheet.Cell(row, col).GetString().Trim();
                if (TryNormalizeMonth(text, out var monthNumber, out var label))
                {
                    monthColumns.Add(new MonthColumn(col, monthNumber, label));
                }
                else if (text.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    totalColumn = col;
                }
            }

            if (monthColumns.Count >= 6)
            {
                return (row, monthColumns, totalColumn);
            }
        }

        return (0, new List<MonthColumn>(), null);
    }

    private static int FindRowByLabel(IXLWorksheet sheet, int lastRow, int lastCol, string label)
    {
        for (var row = 1; row <= lastRow; row++)
        {
            for (var col = 1; col <= Math.Min(lastCol, 6); col++)
            {
                var text = sheet.Cell(row, col).GetString().Trim();
                if (text.Equals(label, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }
        }

        return 0;
    }

    private static string? FindTitle(IXLWorksheet sheet, int lastRow, int lastCol)
    {
        for (var row = 1; row <= Math.Min(lastRow, 5); row++)
        {
            for (var col = 1; col <= lastCol; col++)
            {
                var text = sheet.Cell(row, col).GetString().Trim();
                if (text.StartsWith("FY", StringComparison.OrdinalIgnoreCase))
                {
                    return text;
                }
            }
        }

        return null;
    }

    private static bool TryNormalizeMonth(string input, out int monthNumber, out string label)
    {
        switch (input.Trim().ToLowerInvariant())
        {
            case "jan":
            case "january":
                monthNumber = 1;
                label = "Jan";
                return true;
            case "feb":
            case "february":
                monthNumber = 2;
                label = "Feb";
                return true;
            case "mar":
            case "march":
                monthNumber = 3;
                label = "Mar";
                return true;
            case "apr":
            case "april":
                monthNumber = 4;
                label = "Apr";
                return true;
            case "may":
                monthNumber = 5;
                label = "May";
                return true;
            case "jun":
            case "june":
                monthNumber = 6;
                label = "Jun";
                return true;
            case "jul":
            case "july":
                monthNumber = 7;
                label = "Jul";
                return true;
            case "aug":
            case "august":
                monthNumber = 8;
                label = "Aug";
                return true;
            case "sep":
            case "sept":
            case "september":
                monthNumber = 9;
                label = "Sep";
                return true;
            case "oct":
            case "october":
                monthNumber = 10;
                label = "Oct";
                return true;
            case "nov":
            case "november":
                monthNumber = 11;
                label = "Nov";
                return true;
            case "dec":
            case "december":
                monthNumber = 12;
                label = "Dec";
                return true;
            default:
                monthNumber = 0;
                label = string.Empty;
                return false;
        }
    }

    private static decimal? GetDecimalOrNull(IXLCell cell)
    {
        if (cell.TryGetValue<decimal>(out var dec))
        {
            return dec;
        }

        if (cell.TryGetValue<double>(out var dbl))
        {
            return (decimal)dbl;
        }

        return null;
    }

    private static FyTotals? ComputeFyTotals(FySummary summary, DateTime selectedDate)
    {
        var months = summary.Months
            .Where(m => m.Return.HasValue || m.Cash.HasValue || m.PnL.HasValue)
            .ToList();

        if (months.Count == 0)
        {
            return null;
        }

        var endIndex = months.FindLastIndex(m => m.MonthNumber == selectedDate.Month);
        if (endIndex < 0)
        {
            endIndex = months.Count - 1;
        }

        var slice = months.Take(endIndex + 1).ToList();
        var cashStart = slice.FirstOrDefault(m => m.Cash.HasValue)?.Cash;
        var cashLatest = slice.LastOrDefault(m => m.Cash.HasValue)?.Cash;
        var pnlYtd = SumNullable(slice.Select(m => m.PnL));
        if (!pnlYtd.HasValue && cashStart.HasValue && cashLatest.HasValue)
        {
            pnlYtd = cashLatest.Value - cashStart.Value;
        }

        decimal? returnYtd = null;
        if (cashStart.HasValue && cashLatest.HasValue && cashStart.Value != 0m)
        {
            returnYtd = (cashLatest.Value - cashStart.Value) / cashStart.Value;
        }
        else
        {
            returnYtd = SumNullable(slice.Select(m => m.Return));
        }

        var totalReturn = ComputeTotalReturn(summary);
        var totalCash = ComputeTotalCash(summary);
        var totalPnL = ComputeTotalPnL(summary);

        return new FyTotals(returnYtd, cashLatest, pnlYtd, totalReturn, totalCash, totalPnL);
    }

    private static decimal? ComputeTotalPnL(FySummary summary)
        => summary.TotalPnL ?? SumNullable(summary.Months.Select(m => m.PnL));

    private static decimal? ComputeTotalCash(FySummary summary)
        => summary.TotalCash ?? summary.Months.LastOrDefault(m => m.Cash.HasValue)?.Cash;

    private static decimal? ComputeTotalReturn(FySummary summary)
    {
        if (summary.TotalReturn.HasValue)
        {
            return summary.TotalReturn;
        }

        var firstCash = summary.Months.FirstOrDefault(m => m.Cash.HasValue)?.Cash;
        var lastCash = summary.Months.LastOrDefault(m => m.Cash.HasValue)?.Cash;
        if (firstCash.HasValue && lastCash.HasValue && firstCash.Value != 0m)
        {
            return (lastCash.Value - firstCash.Value) / firstCash.Value;
        }

        return SumNullable(summary.Months.Select(m => m.Return));
    }

    private static string FormatFyMonthLabel(FySummary summary, FyMonth month)
    {
        var year = TryGetFyYear(summary);
        if (!year.HasValue)
        {
            return month.Label;
        }

        var startMonth = summary.Months.Count > 0 ? summary.Months[0].MonthNumber : 1;
        var monthYear = month.MonthNumber >= startMonth ? year.Value : year.Value + 1;
        return $"{month.Label}-{monthYear % 100:00}";
    }

    private static int? TryGetFyYear(FySummary summary)
    {
        if (string.IsNullOrWhiteSpace(summary.Title))
        {
            return null;
        }

        foreach (var token in summary.Title.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (int.TryParse(token, out var year) && year >= 2000 && year <= 2100)
            {
                return year;
            }
        }

        return null;
    }

    private static decimal? SumNullable(IEnumerable<decimal?> values)
    {
        decimal sum = 0m;
        var hasValue = false;
        foreach (var value in values)
        {
            if (!value.HasValue)
            {
                continue;
            }

            sum += value.Value;
            hasValue = true;
        }

        return hasValue ? sum : null;
    }

    private static Panel BuildFyPanel(FySummary? summary, DateTime selectedDate)
    {
        var shouldExpand = AnsiConsole.Profile.Capabilities.Interactive;

        if (summary == null)
        {
            return new Panel(new Markup("[grey]FY data not found.[/]"))
            {
                Header = new PanelHeader("FY Summary", Justify.Left),
                Border = BoxBorder.Rounded,
                Padding = new Padding(1, 0, 1, 0),
                Expand = shouldExpand
            };
        }

        var table = new Table
        {
            Expand = shouldExpand
        }
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Grey37)
            .AddColumn(new TableColumn("Month"))
            .AddColumn(new TableColumn("Return").RightAligned())
            .AddColumn(new TableColumn("Cash").RightAligned())
            .AddColumn(new TableColumn("PnL").RightAligned());

        foreach (var month in summary.Months)
        {
            if (!month.Return.HasValue && !month.Cash.HasValue && !month.PnL.HasValue)
            {
                continue;
            }

            var displayLabel = FormatFyMonthLabel(summary, month);
            var label = month.MonthNumber == selectedDate.Month ? $"[bold]{displayLabel}[/]" : displayLabel;
            table.AddRow(
                label,
                ColorizePercentOrDash(month.Return),
                FormatMoneyOrDash(month.Cash),
                ColorizeChangeOrDash(month.PnL));
        }

        var totalPnL = ComputeTotalPnL(summary);
        var totalReturn = ComputeTotalReturn(summary);
        var totalCash = ComputeTotalCash(summary);

        if (totalPnL.HasValue || totalReturn.HasValue || totalCash.HasValue)
        {
            table.AddRow(
                "[bold]Total[/]",
                ColorizePercentOrDash(totalReturn),
                FormatMoneyOrDash(totalCash),
                ColorizeChangeOrDash(totalPnL));
        }

        return new Panel(table)
        {
            Header = new PanelHeader(summary.Title, Justify.Left),
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };
    }

    private static Panel BuildFyChartPanel(FySummary? summary, DateTime selectedDate, Snapshot snapshot)
    {
        var shouldExpand = AnsiConsole.Profile.Capabilities.Interactive;

        if (summary == null)
        {
            return new Panel(new Markup("[grey]No chart data.[/]"))
            {
                Header = new PanelHeader("FY PnL Chart", Justify.Left),
                Border = BoxBorder.Rounded,
                Padding = new Padding(1, 0, 1, 0),
                Expand = shouldExpand
            };
        }

        var months = summary.Months
            .Select(m => new ChartMonth(FormatFyMonthLabel(summary, m), m.MonthNumber, m.PnL))
            .ToList();

        var currentIndex = months.FindIndex(m => m.MonthNumber == selectedDate.Month);
        if (currentIndex >= 0)
        {
            var current = months[currentIndex];
            months[currentIndex] = current with { PnL = snapshot.MonthToDate };
        }

        var chartMonths = months.Where(m => m.PnL.HasValue).ToList();

        if (chartMonths.Count == 0)
        {
            return new Panel(new Markup("[grey]No PnL data.[/]"))
            {
                Header = new PanelHeader("FY PnL Chart", Justify.Left),
                Border = BoxBorder.Rounded,
                Padding = new Padding(1, 0, 1, 0),
                Expand = shouldExpand
            };
        }

        var max = chartMonths.Max(m => Math.Abs(m.PnL ?? 0m));
        if (max == 0m)
        {
            max = 1m;
        }

        var table = new Table
        {
            Expand = shouldExpand
        }
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Grey37)
            .AddColumn(new TableColumn("Month"))
            .AddColumn(new TableColumn("PnL").NoWrap());

        foreach (var month in chartMonths)
        {
            var bar = BuildBar(month.PnL ?? 0m, max, 18);
            table.AddRow(month.Label, bar);
        }

        return new Panel(table)
        {
            Header = new PanelHeader("FY PnL Chart", Justify.Left),
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };
    }

    private static string BuildBar(decimal value, decimal max, int width)
    {
        var magnitude = Math.Abs(value);
        var ratio = magnitude / max;
        var length = (int)Math.Round((double)(ratio * width));
        length = Math.Clamp(length, 0, width);
        var bar = new string(BarChar, length).PadRight(width);
        var color = value >= 0m ? "green" : "red";
        var sign = value >= 0m ? "+" : "-";
        var abs = Math.Abs(value);
        return $"[{color}]{bar}[/] {sign}{CurrencyPrefix}{abs:N0}";
    }

    private static decimal GetDecimalOrZero(IXLCell cell)
    {
        if (cell.TryGetValue<decimal>(out var dec))
        {
            return dec;
        }

        if (cell.TryGetValue<double>(out var dbl))
        {
            return (decimal)dbl;
        }

        return 0m;
    }

    private static void RenderSnapshot(string workbookPath, string sheetName, Snapshot snapshot, DateTime selectedDate, string? statusMessage, bool interactiveMode, FySummary? fySummary)
    {
        if (AnsiConsole.Profile.Capabilities.Interactive)
        {
            AnsiConsole.Clear();
        }

        var shouldExpand = AnsiConsole.Profile.Capabilities.Interactive;
        var totalWidth = AnsiConsole.Profile.Width;
        if (totalWidth <= 0)
        {
            totalWidth = 120;
        }

        var innerWidth = Math.Max(40, totalWidth - 4);
        var rightHeader1 = "my-portfolio-cli v1.0";
        var rightHeader2 = "by jordi corbilla";
        var rightWidth = Math.Max(rightHeader1.Length, rightHeader2.Length);
        rightWidth = Math.Min(rightWidth, innerWidth - 10);
        var leftWidth = Math.Max(10, innerWidth - rightWidth);

        var headerGrid = new Grid();
        headerGrid.AddColumn(new GridColumn { Width = leftWidth });
        headerGrid.AddColumn(new GridColumn { Width = rightWidth, Alignment = Justify.Right, NoWrap = true });

        headerGrid.AddRow(
            new Markup($"[bold white]Portfolio Status[/]  [grey]({Markup.Escape(sheetName)})[/]"),
            new Markup($"[bold yellow]{rightHeader1}[/]"));

        headerGrid.AddRow(
            new Markup($"[silver]Selected {selectedDate:dddd, MMM d, yyyy}[/]"),
            new Markup($"[bold yellow]{rightHeader2}[/]"));

        if (snapshot.Date != selectedDate.Date)
        {
            headerGrid.AddRow(
                new Markup($"[yellow]Showing last data from {snapshot.Date:dddd, MMM d, yyyy}[/]"),
                new Markup(""));
        }
        else
        {
            headerGrid.AddRow(
                new Markup($"[silver]As of {snapshot.Date:dddd, MMM d, yyyy}[/]"),
                new Markup(""));
        }

        if (!string.IsNullOrWhiteSpace(statusMessage))
        {
            headerGrid.AddRow(
                new Markup($"[grey]{Markup.Escape(statusMessage)}[/]"),
                new Markup(""));
        }

        var header = new Panel(headerGrid)
        {
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.CadetBlue),
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };

        var accountsTable = new Table
        {
            Expand = shouldExpand
        }
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Grey37)
            .AddColumn(new TableColumn("[bold]Account[/]"))
            .AddColumn(new TableColumn("[bold]Value[/]").RightAligned())
            .AddColumn(new TableColumn("[bold]Day PnL[/]").RightAligned())
            .AddColumn(new TableColumn("[bold]Day %[/]").RightAligned());

        foreach (var account in snapshot.Accounts)
        {
            accountsTable.AddRow(
                Markup.Escape(account.Name),
                FormatMoney(account.Current),
                ColorizeChange(account.Change),
                ColorizePercent(account.ChangePct));
        }

        var accountsPanel = new Panel(accountsTable)
        {
            Header = new PanelHeader("Accounts", Justify.Left),
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };

        var summaryGrid = new Grid();
        summaryGrid.AddColumn(new GridColumn().NoWrap());
        summaryGrid.AddColumn(new GridColumn { Alignment = Justify.Right });

        var summaryRowCount = 0;
        summaryGrid.AddRow("Total", FormatMoney(snapshot.Total));
        summaryRowCount++;
        summaryGrid.AddRow("Day PnL", $"{ColorizeChange(snapshot.TotalChange)}  {ColorizePercent(snapshot.TotalChangePct)}");
        summaryRowCount++;
        summaryGrid.AddRow("MTD", $"{ColorizeChange(snapshot.MonthToDate)}  {ColorizePercent(snapshot.MonthToDatePct)}");
        summaryRowCount++;

        var fyTotals = fySummary != null ? ComputeFyTotals(fySummary, selectedDate) : null;
        if (fyTotals != null)
        {
            summaryGrid.AddRow("FY PnL", ColorizeChangeOrDash(fyTotals.PnLYtd));
            summaryRowCount++;
            summaryGrid.AddRow("FY Return", ColorizePercentOrDash(fyTotals.ReturnYtd));
            summaryRowCount++;
            summaryGrid.AddRow("FY Cash", FormatMoneyOrDash(fyTotals.CashLatest));
            summaryRowCount++;

            if (fyTotals.PnLTotal.HasValue)
            {
                summaryGrid.AddRow("FY Total PnL", ColorizeChangeOrDash(fyTotals.PnLTotal));
                summaryRowCount++;
            }

            if (fyTotals.ReturnTotal.HasValue)
            {
                summaryGrid.AddRow("FY Total Return", ColorizePercentOrDash(fyTotals.ReturnTotal));
                summaryRowCount++;
            }

            if (fyTotals.CashTotal.HasValue)
            {
                summaryGrid.AddRow("FY Total Cash", FormatMoneyOrDash(fyTotals.CashTotal));
                summaryRowCount++;
            }
        }

        var summaryPanel = new Panel(summaryGrid)
        {
            Header = new PanelHeader("Summary", Justify.Left),
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };

        var recentTable = new Table
        {
            Expand = shouldExpand
        }
            .Border(TableBorder.Rounded)
            .BorderColor(Color.Grey37)
            .AddColumn(new TableColumn("Date"))
            .AddColumn(new TableColumn("PnL").RightAligned())
            .AddColumn(new TableColumn("%").RightAligned());

        foreach (var day in snapshot.RecentChanges)
        {
            recentTable.AddRow(
                day.Date.ToString("dd-MM-yyyy"),
                ColorizeChange(day.Change),
                ColorizePercent(day.ChangePct));
        }

        var recentPanel = new Panel(recentTable)
        {
            Header = new PanelHeader($"Recent {RecentDayRows} Days", Justify.Left),
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };

        var fyPanel = BuildFyPanel(fySummary, selectedDate);
        var chartPanel = BuildFyChartPanel(fySummary, selectedDate, snapshot);

        var hintsText = interactiveMode
            ? "[grey]Controls:[/]\n[silver]←/→[/] Change month  [silver]↑/↓[/] Change day\n[silver]A[/] Add entry (creates month if missing)  [silver]Q[/] Quit"
            : "[grey]Commands:[/]\n[silver]view[/]  Show snapshot\n[silver]add[/]   Add daily values";

        var hints = new Panel(new Markup(hintsText))
        {
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        };

        var layout = new Layout("root")
            .SplitRows(
                new Layout("header").Size(6),
                new Layout("body"),
                new Layout("footer").Size(5));

        var columnWidth = Math.Clamp(totalWidth / 3, 38, 60);
        var accountsHeight = Math.Max(8, snapshot.Accounts.Count + 5);
        var chartRows = fySummary?.Months.Count(m => m.PnL.HasValue || m.MonthNumber == selectedDate.Month) ?? 0;
        var chartHeight = Math.Max(10, chartRows + 4);
        var summaryHeight = Math.Max(8, summaryRowCount + 4);
        layout["body"].SplitColumns(
            new Layout("left"),
            new Layout("middle").Size(columnWidth),
            new Layout("right").Size(columnWidth));

        layout["left"].SplitRows(
            new Layout("accounts").Size(accountsHeight),
            new Layout("recent"));

        layout["middle"].SplitRows(
            new Layout("fy"),
            new Layout("chart").Size(chartHeight));

        layout["right"].SplitRows(
            new Layout("summary").Size(summaryHeight),
            new Layout("hints").Size(4));

        layout["header"].Update(header);
        layout["accounts"].Update(accountsPanel);
        layout["recent"].Update(recentPanel);
        layout["summary"].Update(summaryPanel);
        layout["fy"].Update(fyPanel);
        layout["chart"].Update(chartPanel);
        layout["hints"].Update(hints);
        layout["footer"].Update(new Panel(new Markup(
            $"[grey]File:[/] {Markup.Escape(workbookPath)}  [grey]Sheet:[/] {Markup.Escape(sheetName)}"))
        {
            Border = BoxBorder.Rounded,
            Padding = new Padding(1, 0, 1, 0),
            Expand = shouldExpand
        });

        AnsiConsole.Write(layout);
    }

    private static string FormatMoney(decimal value)
    {
        var abs = Math.Abs(value);
        var sign = value < 0m ? "-" : "";
        return $"[white]{sign}{CurrencyPrefix}{abs:N2}[/]";
    }

    private static string ColorizeChange(decimal value)
    {
        var color = value >= 0m ? "green" : "red";
        var sign = value >= 0m ? "+" : "-";
        var abs = Math.Abs(value);
        return $"[{color}]{sign}{CurrencyPrefix}{abs:N2}[/]";
    }

    private static string ColorizePercent(decimal value)
    {
        var color = value >= 0m ? "green" : "red";
        var sign = value >= 0m ? "+" : "";
        return $"[{color}]{sign}{value:P2}[/]";
    }

    private static string FormatMoneyOrDash(decimal? value)
        => value.HasValue ? FormatMoney(value.Value) : "[grey]-[/]";

    private static string ColorizeChangeOrDash(decimal? value)
        => value.HasValue ? ColorizeChange(value.Value) : "[grey]-[/]";

    private static string ColorizePercentOrDash(decimal? value)
        => value.HasValue ? ColorizePercent(value.Value) : "[grey]-[/]";

    private static bool DetermineUnicodeSupport()
    {
        var forceAscii = Environment.GetEnvironmentVariable("PORTFOLIO_ASCII");
        if (!string.IsNullOrWhiteSpace(forceAscii) && forceAscii.Trim() == "1")
        {
            return false;
        }

        var forceUnicode = Environment.GetEnvironmentVariable("PORTFOLIO_UNICODE");
        if (!string.IsNullOrWhiteSpace(forceUnicode) && forceUnicode.Trim() == "1")
        {
            return true;
        }

        if (Console.IsOutputRedirected)
        {
            return false;
        }

        if (OperatingSystem.IsWindows())
        {
            var wt = Environment.GetEnvironmentVariable("WT_SESSION");
            var termProgram = Environment.GetEnvironmentVariable("TERM_PROGRAM");
            if (!string.IsNullOrWhiteSpace(wt))
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(termProgram) &&
                termProgram.Contains("vscode", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        return Console.OutputEncoding.WebName.Contains("utf-8", StringComparison.OrdinalIgnoreCase);
    }
}

internal sealed record DateColumn(DateTime Date, int Column);
internal sealed record AccountRow(string Name, int Row);
internal sealed record SheetInfo(string Name, IReadOnlyList<DateColumn> DateColumns, IReadOnlyList<AccountRow> AccountRows, int TotalRow, IXLWorksheet Worksheet);
internal sealed record AccountSnapshot(string Name, decimal Current, decimal Previous, decimal Change, decimal ChangePct);
internal sealed record DailyChange(DateTime Date, decimal Change, decimal ChangePct);
internal sealed record Snapshot(
    DateTime Date,
    IReadOnlyList<AccountSnapshot> Accounts,
    decimal Total,
    decimal TotalChange,
    decimal TotalChangePct,
    decimal MonthToDate,
    decimal MonthToDatePct,
    IReadOnlyList<DailyChange> RecentChanges,
    string SheetName);

internal sealed record SheetSelection(IXLWorksheet? Sheet, DateTime DisplayDate, string? StatusMessage, bool MonthMatched);

internal sealed record MonthColumn(int Column, int MonthNumber, string Label);

internal sealed record FyMonth(string Label, int MonthNumber, decimal? Return, decimal? Cash, decimal? PnL);

internal sealed record FySummary(string Title, IReadOnlyList<FyMonth> Months, decimal? TotalReturn, decimal? TotalCash, decimal? TotalPnL);

internal sealed record FyTotals(decimal? ReturnYtd, decimal? CashLatest, decimal? PnLYtd, decimal? ReturnTotal, decimal? CashTotal, decimal? PnLTotal);

internal sealed class UiState
{
    public UiState(DateTime selectedDate)
    {
        SelectedDate = selectedDate.Date;
    }

    public DateTime SelectedDate { get; set; }
    public string? StatusMessage { get; set; }
    public bool IsMonthMatched { get; set; }
}
