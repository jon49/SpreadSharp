namespace SpreadSharp

open NetOffice.ExcelApi

module XlWorkbook =

    /// <summary>Adds a workbook to an Excel app.</summary>
    /// <returns>The new workbook.</returns>
    let add (xlApp : Application) =
        xlApp.Workbooks.Add()

    /// <summary>Closes a workbook. Use the save and saveAs function to save
    /// a workbook before closing it.</summary>
    /// <param name="workbook">The workbook to close.</param>
    let close (workbook : Workbook) = workbook.Close()

    /// <summary>Opens an existing workbook.</summary>
    /// <param name="appClass">The Excel Application.</param>
    /// <param name="fileName">The name of the workbook file.</param>
    /// <returns>The opened workbook.</returns>
    let openWorkbook (xlApp : Application) fileName = xlApp.Workbooks.Open fileName

    /// <summary>Saves a workbook in the MyDocuments folder.</summary>
    /// <param name="workbook">The workbook to save.</param>
    let save (workbook : Workbook) = workbook.Save()

    /// <summary>Saves a workbook using the specified file name.</summary>
    /// <param name="workbook">The workbook to save.</param>
    /// <param name="fileName">The name of the workbook file.</param>
    let saveAs (workbook : Workbook) (fileName : string) = 
        let missing = System.Type.Missing
        workbook.SaveAs(fileName, missing, missing, missing, missing, missing
                        , Enums.XlSaveAsAccessMode.xlNoChange, missing, missing, missing, missing, missing)