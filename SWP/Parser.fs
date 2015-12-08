open FSharp.Data
open System
open System.Globalization
open System.IO
open System.Collections
open FSharp.Data.Runtime.IO
open System.Windows.Forms



//Parses a csv file and returns the site if a site entry is present
//There should only be one entry in the entire file containing the string "Site:"
//If there are multiple, only the first instance will be returned
let GetSiteInSwipeLog(csv : seq<CsvRow>) =
    let mutable site : string = "No site entry found"
    for row in csv do
        let x : string = string(row.Item(0))
        let cont : bool = x.Contains("Site:") 
        match cont with
        | true  -> site <- x.Replace("Site:", "").Trim()
        | false -> ()
    //return
    site       



//Parses a section and returns the user if a user is present
//There should only be one entry in the entire section containing the string "User Name:"
//If there are multiple, only the first instance will be returned
let GetUserInSection(section : seq<CsvRow>) =
    let mutable user : string = "No user entry found"
    for row in section do
        let x : string = string(row.Item(0))
        let cont : bool = x.Contains("User Name:") 
        match cont with
        | true  -> user <- x.Replace("User Name:", "").Trim()
        | false -> ()
    //return
    user



//Remove a specific substring from a string
//Taken from Stackoverflow question 20308875
let StripString (stripChars:string) (text:string) =
    text.Split(stripChars.ToCharArray(), StringSplitOptions.RemoveEmptyEntries) |> String.Concat



//Extracts a date from an input string and returns a DateTime object representing the date
let ExtractDate(dateString : string) =
    let mutable Excluded : seq<string> = seq [ "(Mon)"; "(Tue)"; "(Wed)"; "(Thu)"; "(Fri)"; "(Sat)"; "(Sun)"; ]
    let mutable s = ""
    let mutable temp = dateString
    //Clean input
    for m in Excluded do
        match temp with
        | temp when temp.Contains(m) -> s <- temp.Replace(m, "")
        | _ -> ()
    
    //Extract Date
    let mutable output = DateTime.MinValue
    match DateTime.TryParseExact(s.Trim(), "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None) with 
    | true,d    -> output <- d 
    | false,_   -> ()
    //Return
    output



//Takes a timestamp with expected h/mm/ss formating and optional
// h/mm/ssTT formating and returns a TimeSpan representation
let ExtractTimeStamp ( timeStamp : string, date : DateTime ) =
    let mutable ttFormated  : bool              = true
    let mutable ttAM        : bool              = false
    let mutable ttPM        : bool              = false
    let mutable output      : Option<DateTime>  = Some DateTime.MinValue
    let mutable temp        : string            = timeStamp.Trim()
    
    //Check for tt formating
    match timeStamp with
    | timeStamp when timeStamp.Contains "AM"    -> ttAM         <- true
    | timeStamp when timeStamp.Contains "PM"    -> ttPM         <- true  
    | _                                         -> ttFormated   <- false

    //The data arrives with timestamps in "h:mm:sstt" format
    //The built in parsing function is unable to parse this format
    //Adding a single space between the second and the AM/PM characters
    //i.e transforming to "h:mm:ss tt" format allows the parsing to work
    match ttFormated with
    | ttFormated when ttFormated = true && ttAM = true  -> temp <- temp.Replace("AM", " AM")
    | ttFormated when ttFormated = true && ttPM = true  -> temp <- temp.Replace("PM", " PM")
    | _                                                 -> ()

    match DateTime.TryParseExact(temp, "h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None) with
    | true,d    -> output <- Some d
    | false,_   -> output <- None
    
    //Return
    output



//Reconstructs a CSV row string to contain the site and user
let CreateSwipeLogEntryRow (row : CsvRow, site : string, user : string) =
    let mutable RowString   : string    = ""
    let mutable date        : DateTime  = DateTime.MinValue
    let mutable inTime      : DateTime  = DateTime.MinValue
    let mutable outTime     : DateTime  = DateTime.MinValue
    let mutable IscsvInTimeNA  : bool      = false
    let mutable IscsvOutTimeNA : bool      = false

    //Store raw row values
    let csvEventDate           : string    = string(row.Item(0))
    let csvInTime              : string    = string(row.Item(1))
    let csvInStatus            : string    = string(row.Item(2))
    let csvInDoor              : string    = string(row.Item(3))
    let csvOutTime             : string    = string(row.Item(4))
    let csvOutStatus           : string    = string(row.Item(5))
    let csvOutDoor             : string    = string(row.Item(6))
    let csvLunch               : string    = string(row.Item(7))
    let csvBreak               : string    = string(row.Item(8))
    let csvSupper              : string    = string(row.Item(9))
    let csvDaily               : string    = string(row.Item(10))
    let csvRunning             : string    = string(row.Item(11))
    //Only include date entries with actual data, remove 'reporting/structural' rows
    match csvEventDate with
    | null                                          -> () //Ignore empty rows
    | csvEventDate when csvEventDate.Contains(":")        -> () //Ignore rows where the csvEventDate column contains a ":"
    | csvEventDate when csvEventDate.Contains("Event")    -> () //Ignore rows where the csvEventDate column contains a "Event"
    | _                                             ->
        //Extract date
        date        <- ExtractDate csvEventDate
        //Determine if timestamps are "N/A" (missing due to not swipe or system fault)
        match csvInTime with
        | "N/A"     -> IscsvInTimeNA <- true
        | _         -> IscsvInTimeNA <- false
        match csvOutTime with 
        | "N/A"     -> IscsvOutTimeNA <- true
        | _         -> IscsvOutTimeNA <- false
        //Clean timestamps
        match ExtractTimeStamp(csvInTime, date) with
        | Some(x)   -> inTime <- inTime.Add(x.TimeOfDay)
        | None      -> ()
        match ExtractTimeStamp(csvOutTime, date) with
        | Some(x)   -> outTime <- outTime.Add(x.TimeOfDay)
        | None      -> ()



        //Handle N/A entries in timestamps
        let mutable inTimeString        = ""
        let mutable outTimeString       = ""
        match IscsvInTimeNA with
        | true      -> inTimeString     <- String.Empty                     // Leave blank
        | false     -> inTimeString     <- inTime.ToOADate().ToString()     // Pass an Excel representation of the DateTime value
        match IscsvOutTimeNA with
        | true      -> outTimeString    <- String.Empty                     // Leave blank
        | false     -> outTimeString    <- outTime.ToOADate().ToString()    // Pass an Excel representation of the DateTime value

        let strings     = [ site; 
                            user; 
                            date.ToOADate().ToString(); 
                            inTimeString; 
                            csvInStatus; 
                            csvInDoor; 
                            outTimeString; 
                            csvOutStatus; 
                            csvOutDoor; 
                            csvLunch; 
                            csvBreak; 
                            csvSupper; 
                            csvDaily; 
                            csvRunning ]
        RowString   <- String.concat ", " strings
    //Return
    RowString



//Takes a section held in a sequence and return it as a CSV string
let CreateSwipeLogSection(section : seq<CsvRow>, site : string) =
    let user : string = GetUserInSection section
    let entries : seq<string> = seq { for s in section do 
                                        yield CreateSwipeLogEntryRow(s, site, user) }
    let entriesWithoutBlanks = entries |> Seq.filter(fun x -> x.Length > 0)
    let SectionString : string = String.concat Environment.NewLine entriesWithoutBlanks
    //Return
    SectionString



//Coordinator function for Swipe Log parsing
let ParseSwipeLog (csv : seq<CsvRow>) =
    let site : string = GetSiteInSwipeLog csv

    //Splitting the file into sections by user
    let mutable tempUDS : seq<seq<CsvRow>> = Seq.empty
    let mutable sect : seq<CsvRow> = Seq.empty
    let mutable cond : bool = false
    for row in csv do
        cond <- row.Item(0).Contains "Regular Time:"
        match cond with
        | false -> sect <- Seq.append sect [row]
        | true -> 
            tempUDS <- Seq.append tempUDS [sect]
            sect <- Seq.empty
    let UserDataSections : seq<seq<CsvRow>> = tempUDS

    //Combines all sections into one CSV file
    let SectionStrings : seq<string> = seq { for section in UserDataSections do 
                                                yield CreateSwipeLogSection(section, site) }
    //Return
    SectionStrings



[<EntryPoint>]
[<STAThreadAttribute>]
let main argv = 
    
    //Ask the user to pick a file
    let ofd = new OpenFileDialog(Filter = "Comma Separated Values files (*.csv)|*.csv", Multiselect = false)
    let ofdResult = ofd.ShowDialog()

    //Load the content of the user-picked file
    let csvData = CsvFile.Load(ofd.FileName).Cache()

    //Ask the user to specify a location and name to save the output file under
    let sfd = new SaveFileDialog(Filter = "Comma Separated Values files (*.csv)|*.csv")
    let sfdResult = sfd.ShowDialog()

    //Create the output file and save it in the user-picked location
    File.WriteAllLines(sfd.FileName, 
        seq { yield "Site, Name, Date, InTime, InStatus, InDoor, OutTime, OutStatus, OutDoor, Lunch, Break, Supper, Daily, Running"; 
                yield! ParseSwipeLog(csvData.Rows)})

    0 // return an integer exit code

