// See https://aka.ms/new-console-template for more information



using ClosedXML.Excel;
using System;
using System.IO;

class Program
{
    static void Main(string[] args) // Empezar el codigo principal
    {
        if (args.Length < 3) // Revisar que se recibieron dos argumentos
        {
            //Console.WriteLine("Usage: dotnet run <ExcelFile1> <OutputFile>");
            Console.WriteLine("Usage: dotnet run <FilterValue> <RawDataFile> <OutputFile>");
            return;
        }
        string FilterValue = args[0];
        string RawDataPath = args[1]; // Raw data file
        //string file2Path = args[1];
        string TargetDataPath = args[2]; //Target file

        if (!File.Exists(TargetDataPath)){ //Checar que el documento exista 
            Console.WriteLine("One or both input files do not exist.");
            return;
        }

        try
        {
            // Load the first Excel file
            using var RawDataWorkbook = new XLWorkbook(RawDataPath);
            var RawDataWorksheet = RawDataWorkbook.Worksheet("Part1");

        
            //Get the filter column number
            
            int FilterColumnNumber = GetColumnNumber(RawDataWorksheet,"Process_Area");
            

            using (var outputworkbook = new XLWorkbook(TargetDataPath)){
                //select worksheet
                var TargetWorksheet = outputworkbook.Worksheet("Export");

                var firstEmptyRow = TargetWorksheet.LastRowUsed().RowNumber() + 1; //Variable to select the last used row and sum 1

                //Erase first part of data in the TargetWorkSheet
                EraseAllRowsExceptFirstInRange(outputworkbook); 




                // foreach (var row in RawDataWorksheet.RowsUsed()){ // Usar cada fila del source
                    
                
                //     var cellValue = row.Cell(FilterColumnNumber).GetValue<string>(); // get the column filter value
                    
                //     if (cellValue == FilterValue){ // Check that it meets filter
                            
                             
                //             CopyRow(row, TargetWorksheet.Row(firstEmptyRow)); // Copy row from source to target next row
                //             firstEmptyRow++;// Increase 
                //         }
                // }



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
        int currentRow = startRow; // Row to start writing data in target file

        foreach (var row in RawData.RowsUsed().Skip(1))  // Skip the first row of the raw data files 
        { 
            foreach (var cell in row.CellsUsed()) // For de celdas usadas, el resto no (seleccionar columna) (ir en cada celda de cada fila del RawData file) 
            {
                TargetData.Cell(currentRow, cell.Address.ColumnNumber).Value = cell.Value; // Darle el valor necesario
            }
            currentRow++;
        }

        return currentRow - 1; //Regresar la ultima fila para poder seguir copiando en la siguente fila
       
    }


    static int GetColumnNumber(IXLWorksheet worksheet, string columnName)
    {
        // Find the column number based on the header row
        var headerRow = worksheet.Row(1);
        foreach (var cell in headerRow.CellsUsed())
        {
            if (cell.GetValue<string>().Equals(columnName, StringComparison.OrdinalIgnoreCase))
            {
                return cell.Address.ColumnNumber;
            }
        }
        return -1; // Column not found
    }

    static void CopyRow(IXLRow sourceRow, IXLRow targetRow) //Toma dos filas
    {
        foreach (var cell in sourceRow.CellsUsed()) // Para cada celda usada de la fuente 
        {
            targetRow.Cell(cell.Address.ColumnNumber).Value = cell.Value; // En cada posicion de la fuente asignarle el valor de la fuente al target
        }
    }

    public static void EraseAllRowsExceptFirst(string filePath)
    {
        // Check if the file exists
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("The specified file does not exist.", filePath);
        }

        // Open the Excel file
        using (var workbook = new XLWorkbook(filePath))
        {
            // Loop through all worksheets in the workbook
            foreach (var worksheet in workbook.Worksheets){
                // Get the total number of rows
                var totalRows = worksheet.RowsUsed().Count();

                // Delete all rows except the first one
                if (totalRows > 1){
                    worksheet.Rows(2, totalRows).Delete(); // Inital row and final row to delete
                }
            }

            // Save the changes back to the file
            workbook.Save();
        }
    }

    public static void ExtendFormulas(string filePath, int[] columnNumbers, int targetRowNumber)
    {
        // Open the Excel file
        using (var workbook = new XLWorkbook(filePath))
        {
            // Loop through all worksheets in the workbook
            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var columnNumber in columnNumbers){
                    // Get the first row with a formula in the specified column
                    //Get the first cell in the column specified with a formula
                    var firstFormulaCell = worksheet.Column(columnNumber)
                                                    .CellsUsed(c => c.HasFormula)
                                                    .FirstOrDefault();

                    if (firstFormulaCell != null)
                    {
                        // Get the formula from the first formula cell
                        var formula = firstFormulaCell.FormulaA1; // Get the formula from the cell 

                        // Extend the formula to the target row number
                        // For loop till specified row (Select the next row in first instance)
                        for (int row = firstFormulaCell.Address.RowNumber + 1; row <= targetRowNumber; row++){ 
                            
                            // Seleccionar the next row formula and give it the same formula
                            worksheet.Cell(row, columnNumber).FormulaA1 = formula; 
                        }
                    }
                }
            }

            // Save the changes back to the file
            workbook.Save();

            /*
            knlknlkn
            */

        }
    }

    public static void EraseAllRowsExceptFirstInRange(XLWorkbook TargetWorkbook) //Funcion para borrar todas las filas excepto la primera
    {
        // Open the Excel file
        // Select the export worksheet
        var worksheet =  workbook.Worksheet("Export");

        // Define the range from column A to column ED
        // Select a range (rectangle) in the excel 
        var range = worksheet.Range("A2:ED" + 2);

        // Clear the content of the range
        range.Clear(XLClearOptions.Contents);
        

        // Save the changes back to the file
        TargetWorkbook.Save();
        
    }


}

