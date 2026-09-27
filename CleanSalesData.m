// CleanSalesData.m - Power Query (M) script
// Imports a raw CSV export of sales orders, cleans it and returns a table
// ready for the dashboard.
//
// How to use in Excel:
//   Data > Get Data > From Other Sources > Blank Query > Advanced Editor
//   Paste this script, change SourcePath below, then Close & Load.
//
// Expected CSV columns (any order, extra columns are ignored):
//   Order ID, Order Date, Region, Salesperson, Product, Units

let
    SourcePath = "C:\Data\sales_export.csv",

    Source = Csv.Document(
        File.Contents(SourcePath),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    // Keep only the columns the dashboard needs
    SelectedColumns = Table.SelectColumns(
        PromotedHeaders,
        {"Order ID", "Order Date", "Region", "Salesperson", "Product", "Units"},
        MissingField.UseNull
    ),

    // Trim spaces and fix capitalisation (e.g. " west " -> "West")
    Trimmed = Table.TransformColumns(
        SelectedColumns,
        {
            {"Order ID", each Text.Upper(Text.Trim(_)), type text},
            {"Region", each Text.Proper(Text.Trim(_)), type text},
            {"Salesperson", each Text.Proper(Text.Trim(_)), type text},
            {"Product", each Text.Trim(_), type text}
        }
    ),

    // Set data types (dates are read using the English (UK) format dd/mm/yyyy)
    Typed = Table.TransformColumnTypes(
        Trimmed,
        {{"Order Date", type date}, {"Units", Int64.Type}},
        "en-GB"
    ),

    // Remove rows with errors, missing values or zero/negative units
    NoErrors = Table.RemoveRowsWithErrors(Typed, {"Order Date", "Units"}),
    Valid = Table.SelectRows(
        NoErrors,
        each [Order ID] <> null and [Order ID] <> ""
            and [Order Date] <> null
            and [Units] <> null and [Units] > 0
    ),

    // Remove duplicate orders
    Deduplicated = Table.Distinct(Valid, {"Order ID"}),

    // Add a Month number column for the dashboard
    WithMonth = Table.AddColumn(Deduplicated, "Month", each Date.Month([Order Date]), Int64.Type),

    Sorted = Table.Sort(WithMonth, {{"Order Date", Order.Ascending}, {"Order ID", Order.Ascending}})
in
    Sorted
