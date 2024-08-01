// 01 - FileFullPath_PasteHere
"https://XXXX.sharepoint.com/sites/XXX/Shared%20Documents/General/XXXXX/File_Final.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

// 02 - RootFullPath
let
    Source = #"01 - FileFullPath_PasteHere",
    #"Converted to Table" = #table(1, {{Source}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Converted to Table", "Column1", Splitter.SplitTextByEachDelimiter({"/"}, QuoteStyle.Csv, true), {"Column1.1", "Column1.2"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Column1.1", type text}, {"Column1.2", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Column1.2"}),
    #"Column1 1" = #"Removed Columns"{0}[Column1.1]
in
    #"Column1 1"

// 03 - RootPath
let
    Source = #"01 - FileFullPath_PasteHere",
    #"Converted to Table" = #table(1, {{Source}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Converted to Table", "Column1", Splitter.SplitTextByEachDelimiter({"sites/"}, QuoteStyle.Csv, true), {"Column1.1", "Column1.2"}),
    #"Added Suffix" = Table.TransformColumns(#"Split Column by Delimiter", {{"Column1.1", each _ & "sites/", type text}}),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Added Suffix", "Column1.2", Splitter.SplitTextByEachDelimiter({"/"}, QuoteStyle.Csv, false), {"Column1.2.1", "Column1.2.2"}),
    #"Changed Type" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"Column1.2.1", type text}, {"Column1.2.2", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Column1.2.2"}),
    #"Merged Columns" = Table.CombineColumns(#"Removed Columns",{"Column1.1", "Column1.2.1"},Combiner.CombineTextByDelimiter("", QuoteStyle.None),"RootPath"),
    RootPath1 = #"Merged Columns"{0}[RootPath]
in
    RootPath1

// 04 - function_root
let
  func = 
    // fnGetSharepointFile
    let
      Source = (FullPath as text) => 
        let
          // Helper function
          fnUriUnescapeString = 
            //Source: https://stackoverflow.com/questions/36242695/how-to-decodeuricomponent-ex-2f3f263d
          (data as text) as text => 
            let
              ToList = List.Buffer(Text.ToList(data)),
              Accumulate = List.Accumulate(
                  ToList, 
                  [Bytes = {}], 
                  (state, current) => 
                    let
                      HexString = state[HexString]?,
                      NextHexString = HexString & current,
                      NextState = 
                        if HexString <> null then 
                          if Text.Length(NextHexString) = 2 then 
                            [
                              Bytes
                                = state[Bytes]
                                  & Binary.ToList(
                                    Binary.FromText(NextHexString, BinaryEncoding.Hex)
                                  )
                            ]
                          else 
                            [HexString = NextHexString, Bytes = state[Bytes]]
                        else if current = "%" then 
                          [HexString = "", Bytes = state[Bytes]]
                        else 
                          [Bytes = state[Bytes] & {Character.ToNumber(current)}]
                    in
                      NextState
                ),
              FromBinary = Text.FromBinary(Binary.FromList(Accumulate[Bytes]))
            in
              FromBinary,
          StaticRoot = SharePoint.Contents(#"03 - RootPath", [ApiVersion = 15]),
          ExtractRoot = fnUriUnescapeString(Text.BeforeDelimiter(FullPath, "/", 4)),
          FileName = fnUriUnescapeString(
              Text.AfterDelimiter(FullPath, "/", {0, RelativePosition.FromEnd})
            ),
          NonRootFolders = fnUriUnescapeString(
              Text.BeforeDelimiter(
                  Text.AfterDelimiter(FullPath, ExtractRoot), 
                  "/", 
                  {0, RelativePosition.FromEnd}
                )
            ),
          SubfoldersList = List.Buffer(List.Select(Text.Split(NonRootFolders, "/"), each _ <> null and _ <> "")),
          NavigateIn = List.Accumulate(
              SubfoldersList, 
              StaticRoot, 
              (state, current) => state{[Name = current]}[Content]
            ),
          #"Filtered Rows" = Table.SelectRows(NavigateIn, each ([Name] = FileName))[Content]{0}
        in
          #"Filtered Rows"
    in
      Source,
  documentation = [
    Documentation.Name = " Sharepoint.GetFile ", 
    Documentation.Description = " Convenient way to get SP file by entering full URL. ", 
    Documentation.LongDescription
      = " Convenient way to get SP file by entering full URL. !! Root path to SP file has to be hardcoded in the function code itself !! ", 
    Documentation.Category = " Accessing Data Functions ", 
    Documentation.Source = "  www.TheBIccountant.com, see:  https://wp.me/p6lgsG-2kR .  ", 
    Documentation.Version = " 1.2: 30-Mar-2021-ImprovedSpeed ", 
    Documentation.Author = " Imke Feldmann ", 
    Documentation.Examples = {[Description = "  ", Code = "  ", Result = "  "]}
  ]
in
  Value.ReplaceType(func, Value.ReplaceMetadata(Value.Type(func), documentation))

// 05 - files_in_folder
let
    Source = #"04 - function_root"(#"02 - RootFullPath")
in
    Source
