namespace FXCore

open System.IO.Compression
open System.IO

module internal Utilities =

    let inline (|?) value defaultValue = defaultArg value defaultValue

    let amount (value:string) =
        let divideArray (values:float[]) = values.[0]/values.[1]
        value.Split("/".[0]) |> Array.map float |> divideArray

    let decompressFileAndRead sourceFile = 
        use zippedFile = new GZipStream(File.OpenRead(sourceFile), CompressionMode.Decompress)
        use streamedFile = new StreamReader(zippedFile)
        streamedFile.ReadToEnd()
