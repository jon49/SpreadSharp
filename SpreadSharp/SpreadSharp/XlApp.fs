namespace SpreadSharp

open NetOffice.ExcelApi

module XlApp =
    

    
    /// <summary>Starts Excel in visible mode.</summary>
    /// <returns>The Excel Application instance.</returns>
    let start () = new Application()
        

    /// <summary>Starts Excel in hidden mode.</summary>
    /// <returns>The created Excel Application instance.</returns>
    let startHidden () =
        let xlApp = new Application()
        xlApp.Visible <- false
        
    /// <summary>Returns a reference to an already running Excel instance.</summary>
    /// <returns>The running Excel Application instance.</returns>
    let getActiveApp () = NetOffice.ExcelApi.GlobalHelperModules.GlobalModule.Application
    
    /// <summary>Closes Excel and releases its related COM objects.</summary>
    /// <param name="appClass">The Excel Application.</param>
    let quit (xlApp : Application) = 
        xlApp.Quit()
        xlApp.Dispose()
        
    /// <summary>Sets the visible property of Excel to false.</summary>
    /// <param name="appClass">The Excel application class instance.</param>
    let hide (xlApp : Application) =
        xlApp.Visible <- false

    /// <summary>Sets the visible property of Excel to true.</summary>
    /// <param name="appClass">The Excel application class instance.</param>
    let unhide (xlApp : Application) =
        xlApp.Visible <- true

    /// <summary>Restores the control of Excel to the user.</summary>
    /// <param name="appClass">The Excel application class instance.</param>
    let restoreUserControl xlApp =
        unhide xlApp
        xlApp.UserControl <- true
        xlApp.Dispose()