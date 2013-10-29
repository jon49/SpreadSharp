namespace SpreadSharp

open System
open NetOffice.ExcelApi
open ExcelExtensions
open Types

module XlRange =

    /// <summary>Returns a worksheet range.</summary>
    /// <param name="worksheet">The worksheet containing the range.</param>
    /// <param name="rangeString">The string representing the range.</param>
    /// <returns>The range object.</returns>
    let get (worksheet : Worksheet) (rangeString : string) =
        worksheet.Range rangeString

    /// <summary>Selects a range of cells.</summary>
    /// <param name="range">The range to select.</param>
    let select (range : Range) = range.Select () |> ignore

    /// <summary>Copies a range to the clipboard.</summary>
    /// <param name="range">The range to copy.</param>
    let copy (range : Range) = range.Copy () |> ignore

    /// <summary>Copies a range to the spedified destination.</summary>
    /// <param name="range">The range to copy.</param>
    /// <param name="destination">The destination range.</param>
    let copyToRange (range : Range) (destination : Range) = range.Copy destination |> ignore

    /// <summary>Cuts a range to the clipboard.</summary>
    /// <param name="range">The range to cut.</param>
    let cut (range : Range) = range.Cut () |> ignore

    /// <summary>Cuts a range and pastes it into the specified destination.</summary>
    /// <param name="range">The range to cut.</param>
    /// <param name="destination">The paste destination range.</param>
    let cutPaste (range : Range) (destination : Range) = range.Cut(destination) |> ignore
    
    /// <summary>Deletes a range.</summary>
    /// <param name="range">The range to delete.</param>
    /// <param name="shift">The shift direction.</param>
    let delete (range : Range) (shift : ShiftDirection) = range.Delete shift |> ignore

    /// <summary>Inserts a cell or a range of cells using the
    /// shift direction and copy origin parameters.</summary>
    /// <param name="range">The range representing the column.</param>
    /// <param name="shift">The shift direction.</param>
    /// <param name="copyOrigin">The column index, count starts at 1.</param>
    let insert (range : Range) (shift : ShiftDirection) = range.Insert shift |> ignore

    /// <summary>Offsets the range by row/column.</summary>
    /// <param name="range">The range to offset.</param>
    /// <param name="byRow">The offset of the row position.</param>
    /// <param name="byRow">The offset of the column position.</param>
    let offset (byRow : int) (byColumn : int) (range : Range) = range.get_Offset(byRow, byColumn)

    /// <summary>Resizes the range by row/column.</summary>
    /// <param name="range">The range to resize.</param>
    /// <param name="byRow">The new size of the row.</param>
    /// <param name="byRow">The new size of the column.</param>
    let resize (byRow : int) (byColumn : int) (range : Range) = range.get_Resize(byRow, byColumn)

    /// <summary>Performs an autofill from a source range to a destination one. The two ranges must overlap.</summary>
    /// <param name="range">The range from which to start.</param>
    /// <param name="destination">The destination range.</param>
    /// <param name="autoFillType">The auto fill type.</param>
    let autoFill (range : Range) destination autoFillType = range.AutoFill(destination, autoFillType) |> ignore

    /// <summary>Format number style in range.</summary>
    /// <param name="range">The range to format.</param>
    /// <param name="format">Style of number.</param>
    let numberFormat (format : string) (range : Range) = range.NumberFormat <- format

    /// <summary>Sets the value of a range.</summary>
    /// <param name="value">The value to use.</param>
    /// <param name="range">The range object.</param>
    let setValue (value : Object) (range : Range) = range.ToExcel value

    /// <summary>Sets the value of a range.</summary>
    /// <param name="value">The value to use.</param>
    /// <param name="range">The range object.</param>
    let setFormula (value : Object) (range : Range) = range.ToExcelFormula value

    /// <summary>Gets the value of a range as 2D object array.</summary>
    /// <param name="range">The range object.</param>
    let getValue (range : Range) = range.To2dArray()

    /// <summary>Gets the current region of range.</summary>
    let currentRegion (range : Range) = range.CurrentRegion

    module Column =

        /// <summary>Returns the range representing the column with the specified index.</summary>
        /// <param name="worksheet">The worksheet containing the column.</param>
        /// <param name="idx">The column index, count starts at 1.</param>
        let byIndex (worksheet : Worksheet) (idx : int) =
            worksheet.Columns.[idx]
            |> fun x -> x.EntireColumn

        /// <summary>Returns the range representing the column with the specified header.</summary>
        /// <param name="worksheet">The worksheet containing the column.</param>
        /// <param name="header">The column header.</param>
        let byHeader (worksheet : Worksheet) (header : string) =
            worksheet.Columns.[header]
            |> fun x -> x.EntireColumn

        /// <summary>Inserts a column using shift direction and copy origin parameters.</summary>
        /// <param name="range">The range representing the column.</param>
        /// <param name="shift">The shift direction.</param>
        let insert (range : Range) (shift : ShiftDirection) = range.EntireColumn.Insert shift |> ignore

        /// <summary>Hides a column.</summary>
        /// <param name="range">The range representing the column to hide.</param>
        let hide (range : Range) = range.EntireColumn.Hidden <- true

        /// <summary>Displays a hidden column.</summary>
        /// <param name="range">The range representing the hidden column.</param>
        let unhide (range : Range) = range.EntireColumn.Hidden <- false

    module Row =

        /// <summary>Returns the range representing the row with the specified index.</summary>
        /// <param name="worksheet">The worksheet containing the row.</param>
        /// <param name="idx">The row index, count starts at 1.</param>
        let byIndex (worksheet : Worksheet) (idx : int) =
            worksheet.Rows.[idx]
            |> fun x -> x.EntireRow

        /// <summary>Insert a row using the shift direction and copy origin parameters.</summary>
        /// <param name="range">The range representing the row.</param>
        /// <param name="shift">The shift direction.</param>
        /// <param name="copyOrigin">The copy origin parameter.</param>
        let insert (range : Range) (shift : ShiftDirection) = range.EntireRow.Insert shift |> ignore

        /// <summary>Hides a row.</summary>
        /// <param name="range">The range representing the row to hide.</param>
        let hide (range : Range) = range.EntireRow.Hidden <- true

        /// <summary>Displays a hidden row.</summary>
        /// <param name="range">The range representing the hidden row.</param>
        let unhide (range : Range) = range.EntireRow.Hidden <- false