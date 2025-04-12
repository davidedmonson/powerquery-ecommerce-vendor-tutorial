= let
    // Retrieve the File Path from the named range "FilePath"
    FilePath = Excel.CurrentWorkbook(){[Name="FilePath"]}[Content]{0}[Column1],
    
    // Combine the FilePath with the file name "TQxNyBwElbls.csv"
    FullPathToFile1 = FilePath & "TQxNyBwElbls.csv",
    
    // Load the CSV file from the full path
    Source = Csv.Document(
        File.Contents(FullPathToFile1), 
        [Delimiter = ",", Encoding = 65001, QuoteStyle=QuoteStyle.Csv]
    )
in
    Source