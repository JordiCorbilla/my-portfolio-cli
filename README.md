# my-portfolio-cli

An interactive, console-first portfolio tracker built in C# with Spectre.Console and ClosedXML. It reads an Excel workbook, renders a rich TUI dashboard, and lets you add daily snapshots directly from the terminal.

This project was vibe-coded with Chat GPT-5-2-Codex over a weekend.

## Highlights
- Interactive dashboard with keyboard navigation.
- Daily PnL, MTD, and FY summaries.
- FY table and PnL bar chart.
- Adds daily entries and creates new month sheets on demand.
- Uses a demo workbook for safe sharing.

## Requirements
- .NET 9 SDK
- Windows terminal that supports ANSI rendering (Windows Terminal or VS Code terminal recommended)

## Quick Start
```powershell
dotnet run --project .\
```

By default it loads `my_portfolio.xlsx` from the repo root. Use a different file with:
```powershell
dotnet run --project .\ -- --file .\demo_portfolio.xlsx
```

## Controls (Interactive Mode)
- `?/?` Change month
- `?/?` Change day
- `A` Add entry (creates the month if missing)
- `Q` Quit

## Commands (Optional)
```powershell
dotnet run --project .\ -- view --file .\my_portfolio.xlsx
dotnet run --project .\ -- add --date 2026-03-01 --file .\my_portfolio.xlsx
dotnet run --project .\ -- interactive --file .\my_portfolio.xlsx
```

## Demo Workbook
Use the included `my_portfolio.xlsx` for screenshots, demos, and sharing.

## Month Creation
If you navigate to a month that doesn’t exist (e.g., March) and press `A`, the CLI will:
1. Create a new sheet by copying the previous month.
2. Carry forward the latest filled day as the new baseline.
3. Prompt you for the selected day’s values.

## Unicode vs ASCII
Some Windows terminals don’t render Unicode bars or the £ symbol. The app auto-falls back, but you can force behavior:
```powershell
$env:PORTFOLIO_UNICODE = "1"  # force £ and ¦
$env:PORTFOLIO_ASCII = "1"    # force GBP and #
```

## Tech Stack
- C# / .NET 9
- Spectre.Console
- ClosedXML

## Screenshots
<img width="1913" height="1015" alt="image" src="https://github.com/user-attachments/assets/6eaf0e04-a008-48ca-bf57-9d8a38eb509f" />


