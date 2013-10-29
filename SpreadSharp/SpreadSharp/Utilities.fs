namespace SpreadSharp

open Microsoft.FSharp.Reflection
open System
open NetOffice.ExcelApi

module private Utilities =
    
    let boxOrMissing<'T> = function Some (x : 'T) -> box x | None -> Type.Missing

    let setWorksheetName nameOption (worksheet : Worksheet) =
        match nameOption with
            | None      -> worksheet
            | Some name ->
                worksheet.Name <- name
                worksheet

    let recordFieldsNames recordType =
        FSharpType.GetRecordFields recordType
        |> Array.map (fun x -> box x.Name)

    let recordSeqFieldsNames (records: 'T seq) =
        FSharpType.GetRecordFields (typeof<'T>)
        |> Array.map (fun x -> box (x.Name.Replace("_", " ")))

    let fieldsArray records =
        records
        |> Seq.map (fun record ->
            FSharpValue.GetRecordFields record)