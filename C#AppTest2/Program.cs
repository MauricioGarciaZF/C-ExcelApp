// See https://aka.ms/new-console-template for more information



using ClosedXML.Excel;
using System;
using System.IO;

class Program
{
    static void Main(string[] args) // Empezar el codigo principal
    {
        if (args.Length < 2) // Revisar que se recibieron dos argumentos
        {
            //Console.WriteLine("Usage: dotnet run <ExcelFile1> <OutputFile>");
            Console.WriteLine("Usage: dotnet run <RawDataFile> <OutputFile>");
            return;
        }

        string RawDataPath = args[0]; // Raw data file
        //string file2Path = args[1];
        string TargetDataPath = args[1]; //Target file

        if (!File.Exists(TargetDataPath)) //Checar que el documento exista 
        {
            Console.WriteLine("One or both input files do not exist.");
            return;
        }

        try
        {
            // Load the first Excel file
            using var workbook1 = new XLWorkbook(RawDataPath);
            var worksheet1 = workbook1.Worksheet(1);





            // Load the second Excel file
            //using var workbook2 = new XLWorkbook(file2Path);
            //var worksheet2 = workbook2.Worksheet(1);

            // Create a new workbook for the output 
            //using var outputWorkbook = new XLWorkbook(TargetDataPath);
            //var outputWorksheet = outputWorkbook.Worksheets.Add("JoinedData");

            using (var outputworkbook = new XLWorkbook(TargetDataPath))
            {
                //select worksheet
                var worksheet = outputworkbook.Worksheet("Export");

                var firstEmptyRow = worksheet.LastRowUsed().RowNumber() + 1; 


                //Add data
                //worksheet.Cell(firstEmptyRow, 5).Value = "New Data 1";
                //worksheet.Cell(firstEmptyRow, 2).Value = "New Data 2";
                //worksheet.Cell(firstEmptyRow, 3).Value = DateTime.Now;


                CopyRows()





                //Save data
                outputworkbook.Save();
                Console.WriteLine($"Data added succesfully");

            }



                // Copy rows from the first worksheet
                // var lastRow = CopyRows(worksheet1, outputWorksheet, 2); // Copiar las filas del worksheet1 en la primera fila

                // Copy rows from the second worksheet
                //CopyRows(worksheet2, outputWorksheet, lastRow + 2); // Copiar las filas del worksheet2 en la siguiente fila.

                // Save the output file
                //outputWorkbook.SaveAs(TargetDataPath);

                //Console.WriteLine($"Rows from both files have been joined and saved to {TargetDataPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static int CopyRows(IXLWorksheet RawData, IXLWorksheet TargetData, int startRow) //Funcion para copiar filas
    {
        int currentRow = startRow;

        foreach (var row in RawData.RowsUsed())  // For de filas usadas, el resto no (seleccionar fila) (Filas del RawData file)
        { 
            foreach (var cell in row.CellsUsed()) // For de celdas usadas, el resto no (seleccionar columna) (ir en cada celda de cada fila del RawData file) 
            {
                TargetData.Cell(currentRow, cell.Address.ColumnNumber).Value = cell.Value; // Darle el valor necesario
            }
            currentRow++;
        }

        return currentRow - 1; //Regresar la ultima fila para poder seguir copiando en la siguente fila
       
    }


}

