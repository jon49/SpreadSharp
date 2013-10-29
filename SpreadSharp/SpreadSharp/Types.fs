namespace SpreadSharp

open System
open NetOffice.ExcelApi.Enums

module Types =

    type ShiftDirection =
        | Down
        | ExcelDecide
        | Right

        member x.Box () =
            match x with
                | Down        -> box XlInsertShiftDirection.xlShiftDown
                | ExcelDecide -> Type.Missing
                | Right       -> box XlInsertShiftDirection.xlShiftToRight

    