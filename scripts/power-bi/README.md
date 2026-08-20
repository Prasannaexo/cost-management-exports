# Power BI FinOps Dashboard

A parameter-driven Power BI template for the cost data in
[Get-AzureAccessReview.ps1](../access-review/Get-AzureAccessReview.ps1)'s
sibling automation, [New-CostManagementExports.ps1](../New-CostManagementExports.ps1).
Built to the same standard a dedicated FinOps team would expect: day-wise
trends, month-over-month comparisons, and subscription/service breakdowns,
all driven from one parameter instead of hardcoded connection details.

**Why this is a query/measure library instead of a `.pbix` file:** a Power BI
file is Power BI Desktop's own compiled format (a compressed tabular model
plus an undocumented, versioned visual-layout schema) -- there's no reliable
way to hand-author one outside Power BI Desktop itself without risking a
corrupted file that won't open. What's here instead is everything Desktop
needs to build the real thing in about 15 minutes: paste-ready M queries, a
full DAX measure library, and an exact page-by-page layout to recreate.
Once built, **File > Save As > Power BI template (.pbit)** turns it into a
real reusable file others on the team can open directly.

## What you get

- One Power BI **Parameter** for the storage account name -- change it once
  (Home > Transform Data > Edit Parameters) to point the whole report at a
  different environment's storage account, no query edits required.
- Day-wise, month-wise, subscription-wise, and service-wise cost views.
- A professional measure library: month-to-date, day-over-day change,
  month-over-month change, 7-day rolling average, and a spike/anomaly flag.
- A three-page layout matching how FinOps teams typically split cost
  visibility: Executive Summary, Day-over-Day Trend, Service & Subscription
  Breakdown.

## Prerequisites

- `Storage Blob Data Reader` on `stldmcostexports` (or whichever storage
  account you point the parameter at) for your own Entra account -- shared
  keys are disabled, so this is the only way in.
- Power BI Desktop, signed in with your organizational account.

## Step 1 -- Create the parameter

Home > **Manage Parameters** > New Parameter:

- Name: `StorageAccountName`
- Type: Text
- Current Value: `stldmcostexports`

Every query below references `StorageAccountName` instead of a hardcoded
string -- this is what makes the whole report portable across environments.

## Step 2 -- Cost data query

Real schema, confirmed against a live export (CSP/reseller "actual cost"
format):

| Need | Column |
| --- | --- |
| Day-wise | `date` (text, `MM/DD/YYYY` -- must be parsed with an explicit `en-US` locale, see below) |
| Subscription-wise | `subscriptionName` |
| Service-wise (high level) | `serviceFamily`, e.g. `Networking`, `Compute`, `Storage` |
| Service-wise (granular) | `meterCategory`, e.g. `Load Balancer`, `Virtual Machines` |
| Cost | `costInUsd` (normalized -- use this for cross-subscription comparison) or `costInBillingCurrency` (actual invoiced **INR**, via the Meridian Solutions reseller -- don't mix the two in one visual) |

Each export writes a **full month-to-date snapshot on every run**, not
incremental data (confirmed: three separate runs existed for one August
period, each a complete restatement) -- so this query dedupes to the latest
run per subscription per billing period before parsing, to avoid triple
counting costs.

New Blank Query, Advanced Editor, name it `CostData`:

```m
let
    Source = AzureStorage.Blobs(StorageAccountName),
    Container = Source{[Name="cost-management-exports"]}[Data],
    FilteredCsv = Table.SelectRows(Container, each Text.EndsWith([Name], "000001.csv")),
    AddSubscriptionLabel = Table.AddColumn(FilteredCsv, "SubscriptionLabel", each Text.Split([Name], "/"){1}, type text),
    AddPeriod = Table.AddColumn(AddSubscriptionLabel, "Period", each Text.Split([Name], "/"){3}, type text),
    AddRunTimestamp = Table.AddColumn(AddPeriod, "RunTimestamp", each Text.Split([Name], "/"){4}, type text),
    GroupedLatest = Table.Group(AddRunTimestamp, {"SubscriptionLabel", "Period"}, {{"Latest", each Table.Max(_, "RunTimestamp")}}),
    ExpandLatest = Table.ExpandRecordColumn(GroupedLatest, "Latest", {"Content"}, {"Content"}),
    AddCostTable = Table.AddColumn(ExpandLatest, "CostTable", (currentRow) =>
        let
            Parsed = Csv.Document(currentRow[Content], [Delimiter=",", Encoding=65001]),
            Promoted = Table.PromoteHeaders(Parsed, [PromoteAllScalars=true]),
            Tagged = Table.AddColumn(Promoted, "SubscriptionLabel", each currentRow[SubscriptionLabel])
        in
            Tagged
    ),
    Combined = Table.Combine(AddCostTable[CostTable]),
    FixedDate = Table.TransformColumnTypes(Combined, {{"date", type date}}, "en-US"),
    AddYearMonth = Table.AddColumn(FixedDate, "YearMonth", each Date.ToText([date], "yyyy-MM"), type text),
    TypedCosts = Table.TransformColumnTypes(AddYearMonth, {{"costInUsd", type number}, {"costInBillingCurrency", type number}})
in
    TypedCosts
```

**Why the explicit `"en-US"` locale on the date conversion**: `date` comes
through as `MM/DD/YYYY` text. If Power BI's regional settings on your
machine default to a different locale (common if Desktop is set to en-IN),
letting it auto-detect the type can silently swap day and month. The
locale argument above forces the correct interpretation regardless of your
machine's settings.

## Step 3 -- Date table

New Blank Query, name it `DateTable`:

```m
let
    MinDate = Date.From(List.Min(CostData[date])),
    MaxDate = Date.From(Date.EndOfMonth(List.Max(CostData[date]))),
    DateList = List.Dates(MinDate, Duration.Days(MaxDate - MinDate) + 1, #duration(1,0,0,0)),
    ToTable = Table.FromList(DateList, Splitter.SplitByNothing(), {"Date"}),
    AddYear = Table.AddColumn(ToTable, "Year", each Date.Year([Date]), Int64.Type),
    AddMonthNum = Table.AddColumn(AddYear, "MonthNumber", each Date.Month([Date]), Int64.Type),
    AddMonthName = Table.AddColumn(AddMonthNum, "MonthName", each Date.ToText([Date], "MMM yyyy"), type text),
    AddYearMonth = Table.AddColumn(AddMonthName, "YearMonth", each Date.ToText([Date], "yyyy-MM"), type text)
in
    AddYearMonth
```

After loading: right-click `DateTable` in the Fields pane > **Mark as Date
Table** > pick `Date`. Then in Model view, drag a relationship from
`DateTable[Date]` to `CostData[date]`.

## Step 4 -- DAX measure library

Create these against `CostData` (New Measure, or a dedicated measures table
if you prefer keeping them off the fact table):

```dax
Total Cost = SUM(CostData[costInUsd])
Total Cost (Billing Currency) = SUM(CostData[costInBillingCurrency])

Cost - Prior Day = CALCULATE([Total Cost], DATEADD(DateTable[Date], -1, DAY))
DoD Change = [Total Cost] - [Cost - Prior Day]
DoD % Change = DIVIDE([DoD Change], [Cost - Prior Day])

Cost - Prior Month = CALCULATE([Total Cost], PREVIOUSMONTH(DateTable[Date]))
MoM Change = [Total Cost] - [Cost - Prior Month]
MoM % Change = DIVIDE([MoM Change], [Cost - Prior Month])

MTD Cost = TOTALMTD([Total Cost], DateTable[Date])
Prior MTD Cost = CALCULATE([MTD Cost], DATEADD(DateTable[Date], -1, MONTH))
MTD vs Prior MTD % = DIVIDE([MTD Cost] - [Prior MTD Cost], [Prior MTD Cost])

7-Day Avg Daily Cost =
AVERAGEX(
    DATESINPERIOD(DateTable[Date], MAX(DateTable[Date]), -7, DAY),
    [Total Cost]
)

Anomaly Flag =
IF(
    [Total Cost] > 1.5 * [7-Day Avg Daily Cost],
    "Spike",
    "Normal"
)

Cost Share % of Total =
DIVIDE([Total Cost], CALCULATE([Total Cost], ALL(CostData[serviceFamily])))
```

## Step 5 -- Page layout

**Page 1: Executive Summary**
- Three KPI cards: `MTD Cost`, `MoM % Change`, `DoD % Change`
- Line chart: X = `DateTable[Date]`, Y = `Total Cost`, Legend =
  `subscriptionName` -- enable the built-in **Forecast** in the Analytics
  pane for a professional trend projection with no extra DAX
- Slicers across the top: `subscriptionName`, `serviceFamily`, date range
  on `DateTable[Date]`

**Page 2: Day-over-Day Trend**
- Column chart: X = `DateTable[Date]`, Y = `Total Cost`, with data labels
  colored by `Anomaly Flag` (conditional formatting: red/orange for
  "Spike") -- this is the "what changed and when" view
- Table: `date`, `subscriptionName`, `Total Cost`, `DoD % Change`, sorted
  by date descending, with `DoD % Change` conditionally formatted (red for
  increases past a threshold, e.g. background color rules bound to the
  measure)

**Page 3: Service & Subscription Breakdown**
- Bar chart: Y = `serviceFamily`, X = `Total Cost`, sorted descending
- Drill-down bar: add `meterCategory` beneath `serviceFamily` in the same
  axis well for one-click drill to the granular meter level
- Matrix: Rows = `subscriptionName` > `serviceFamily`, Columns =
  `DateTable[YearMonth]`, Values = `Total Cost` -- the subscription x
  service x month pivot FinOps reviews are usually built around
- Table with drill-through target: `date`, `subscriptionName`,
  `resourceGroupName`, `meterName`, `Total Cost`, for right-click
  drill-through from any visual on Pages 1-2 down to line-item detail

## Polish

- Format all cost measures as currency (`$#,##0.00` for `Total Cost`, since
  it's USD-normalized).
- Assign a consistent color per `subscriptionName` in the report theme
  (Format pane > Colors) so the same subscription is always the same color
  across every page.
- Add a **tooltip page** showing the top 5 `meterName` values for whatever
  bar/date point is hovered -- Format pane > Tooltip on the relevant visual.

## Making it reusable

Once built, **File > Save As**, choose **Power BI template (.pbit)**
instead of `.pbix`. Opening the template prompts for `StorageAccountName`
before loading anything -- hand the `.pbit` to anyone else on the team and
they get the exact same report against whatever storage account they enter,
without touching a single query.

## Related

- [../../README.md](../../README.md) -- how the underlying cost exports and
  storage account firewall are provisioned.
- [../access-review/README.md](../access-review/README.md) -- the sibling
  monthly access-review workbook in the same storage account; its Power
  Query connection pattern (blob storage + `Excel.Workbook`) is different
  from the CSV approach here since it's a different data shape.
