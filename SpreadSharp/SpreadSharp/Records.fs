namespace SpreadSharp

module Records =

    let private recordsRange headers fields worksheet =
        let columnsCount = Array.length headers
        let length = Seq.length fields + 1 |> string
        let rangeString = String.concat "" ["A1:"; string (char (columnsCount + 64)) + length]
        XlRange.get worksheet rangeString

    let private displayFields records range =
        let fields = Utilities.fieldsArray records
        let array = array2D fields
        XlRange.setValue array range 

    let private displayFieldNames recordType range =
         let headers = Utilities.recordFieldsNames recordType
         let array = Array2D.init 1 headers.Length (fun i j -> headers.[j])
         XlRange.setValue array range 

    let private displayRecords records recordType range =
        let headers = Utilities.recordFieldsNames recordType
        let fields = Utilities.fieldsArray records
        let array = Array2D.ofSeqs headers fields
        XlRange.setValue array range 

    let private displayRecords' records recordType worksheet =
        let headers = Utilities.recordFieldsNames recordType
        let fields = Utilities.fieldsArray records
        let array = Array2D.ofSeqs headers fields
        let range = recordsRange headers fields worksheet
        XlRange.setValue array range 

    /// <summary>Sends the values of a collection of F# records to an Excel range.</summary>
    /// <param name="rcords">The F# records array.</param>
   /// <param name="range">The range object.</param>         
    let toRange (records : 'T seq) range = displayRecords records (typeof<'T>) range

    ///Returns records as 2D object array.
    let fieldsArray records = Utilities.fieldsArray records |> array2D

//    let fieldsToRange records range = displayFields records range
//    let fieldNamesToRange (records : 'T seq) range = displayFieldNames (typeof<'T>) range

    /// <summary>Sends the values of a collection of F# records to an Excel worksheet.</summary>
    /// <param name="rcords">The F# records collection.</param>
    /// <param name="worksheet">The worksheet object.</param>
    let toWorksheet (records : 'T seq) worksheet = displayRecords' records (typeof<'T>) worksheet

    /// <summary>Saves a collection of records in a workbook using the specified file name.</summary>
    /// <param name="rcords">The records collection.</param>
    /// <param name="filename">The destination file name.</param>
    let saveAs records filename =
        let app = XlApp.start ()
        let wb = XlWorkbook.add app
        let ws = XlWorksheet.byIndex wb 1
        toWorksheet records ws |> ignore
        XlWorkbook.saveAs wb filename
        XlWorkbook.close wb
        XlApp.quit app

    let fieldsNames (records: 'T seq) =
        Utilities.recordSeqFieldsNames records
   
