// See https://aka.ms/new-console-template for more information



using ClosedXML.Excel;
using System;
using System.IO;
using System.Transactions;

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

                //Erase first part of data in the TargetWorkSheet
                EraseAllRowsExceptFirstInRange(outputworkbook); //Erase both ranges from data and formulas

                //var firstEmptyRow = TargetWorksheet.LastRowUsed().RowNumber() +1; //Variable to select the last used row and sum 1
                
                var firstEmptyRow = 2;
                CopyDataRows(outputworkbook,RawDataWorkbook,firstEmptyRow,FilterValue,FilterColumnNumber);
                
                //ExtendFormulasInRange(outputworkbook);

                ExtendOneFormulaInRange(outputworkbook);

                //int columntest = TargetWorksheet.Column("EH").ColumnNumber();

                //TargetWorksheet.Cell(3, columntest).FormulaA1 = "=\"EA"+"3"+"\"" ;

                //TargetWorksheet.Cell(3, columntest).FormulaA1 = "=IF(EA"+3+"<>\"\",\"Satisfied\",\"No Satisfy Link\")" ;  

            
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
        catch (Exception ex){
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    public static void ExtendOneFormulaInRange(XLWorkbook TargetWorkbook){

        var worksheet =  TargetWorkbook.Worksheet("Export");

        // Get the last used row in the worksheet
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;


        // Statement and then increase the current row 
        //"=IF("+ CurrentRow +"\"<>\"\",\"Satisfied\",\"No Satisfy Link\")";

        /*
        =IF(EA2<>"","Satisfied","No Satisfy Link") E
        =IF(ISERROR(VLOOKUP(K2,Exclusion!A:A,1,FALSE)),H"VALID","INVALID") EI
        =IF(ISERROR(VLOOKUP(K2,Export_S002_210624!S:S,1,FALSE)),"YES",IF(AND(VLOOKUP(K2,Export_S002_210624!S:BN,29,FALSE)="",VLOOKUP(K2,Export_S002_210624!S:BN,6,FALSE)<>"Pass"),"NO","YES")) EJ
        =IF(AR2<>"",VLOOKUP(AR2,SwBehav2Team!A:C,2,FALSE),"No SW Behavior assigned") EK
        =IF(AR2<>"",VLOOKUP(AR2,SwBehav2Team!A:C,3,FALSE),"No SW Behavior assigned") EL

        =IF(ISERROR(VLOOKUP(K2,Export_S002_210624!S:S,1,FALSE)),"YES","NO") EN
        =IF(VLOOKUP(K2,Export_S002_210624!S:BN,29,FALSE)="","EMPTY","LINKED") EO
        =VLOOKUP(K2,Export_S002_210624!S:BN,6,FALSE) EP

        */
        

        // Apply the formula to all rows from row 4 to the last row
        for (int row = 2; row <= lastRow; row++){
            worksheet.Cell(row, worksheet.Column("EH").ColumnNumber()).FormulaA1 = "=IF(EA"+row+"<>\"\",\"Satisfied\",\"No Satisfy Link\")";
            worksheet.Cell(row, worksheet.Column("EI").ColumnNumber()).FormulaA1 = "=IF(ISERROR(VLOOKUP(K"+row+",Exclusion!A:A,1,FALSE)),\"VALID\",\"INVALID\")" ; 
            worksheet.Cell(row, worksheet.Column("EJ").ColumnNumber()).FormulaA1 = "=IF(ISERROR(VLOOKUP(K"+row+",Export_S002_210624!S:S,1,FALSE)),\"YES\",IF(AND(VLOOKUP(K"+row+",Export_S002_210624!S:BN,29,FALSE)=\"\",VLOOKUP(K"+row+",Export_S002_210624!S:BN,6,FALSE)<>\"Pass\"),\"NO\",\"YES\"))";
            worksheet.Cell(row, worksheet.Column("EK").ColumnNumber()).FormulaA1 = "=IF(AR"+row+"<>\"\",VLOOKUP(AR"+row+",SwBehav2Team!A:C,2,FALSE),\"No SW Behavior assigned\")";
            worksheet.Cell(row, worksheet.Column("EL").ColumnNumber()).FormulaA1 = "=IF(AR"+row+"<>\"\",VLOOKUP(AR"+row+",SwBehav2Team!A:C,3,FALSE),\"No SW Behavior assigned\")";

            worksheet.Cell(row, worksheet.Column("EN").ColumnNumber()).FormulaA1 = "=IF(ISERROR(VLOOKUP(K"+row+",Export_S002_210624!S:S,1,FALSE)),\"YES\",\"NO\")";
            worksheet.Cell(row, worksheet.Column("EO").ColumnNumber()).FormulaA1 = "=IF(VLOOKUP(K"+row+",Export_S002_210624!S:BN,29,FALSE)=\"\",\"EMPTY\",\"LINKED\")" ;
            worksheet.Cell(row, worksheet.Column("EP").ColumnNumber()).FormulaA1 = "=VLOOKUP(K"+row+",Export_S002_210624!S:BN,6,FALSE)";
        }

        
    
        // Save the changes back to the file
        TargetWorkbook.Save();
    
    }

    public static void ExtendFormulasInRange(XLWorkbook TargetWorkbook){

        var worksheet =  TargetWorkbook.Worksheet("Export");

        // Get the last used row in the worksheet
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        // If there are rows beyond row 3, extend the formulas
        if (lastRow > 2){
            // Loop through each column in the range EE to EP
            for (int col = worksheet.Column("EE").ColumnNumber(); col <= worksheet.Column("EP").ColumnNumber(); col++){
                // Get the formula in the cell EE3 to EP3
                var formulaCell = worksheet.Cell(2, col);
                if (formulaCell.HasFormula){
                    var formula = formulaCell.FormulaA1;

                    // Apply the formula to all rows from row 4 to the last row
                    for (int row = 3; row <= lastRow; row++){
                        worksheet.Cell(row, col).FormulaA1 = formula;
                    }
                }
            }
        }
    
        // Save the changes back to the file
        TargetWorkbook.Save();
    
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

    public static void EraseAllRowsExceptFirstInRange(XLWorkbook TargetWorkbook) //Funcion para borrar todas las filas excepto la primera
    {
        // Open the Excel file
        // Select the export worksheet
        var worksheet =  TargetWorkbook.Worksheet("Export");

        // Define the range from column A to column ED
        // Select a range (rectangle) in the excel 
        var rangeA = worksheet.Range("A2:ED" + worksheet.LastRowUsed().RowNumber()); // Define range A2 untill ED{last row}

        var rangeB = worksheet.Range("EE2:EP" + worksheet.LastRowUsed().RowNumber()); // Select untill the last row used in the doc

        // Clear the content of the range
        rangeA.Clear(XLClearOptions.Contents);
        rangeB.Clear(XLClearOptions.Contents);
        
        // Save the changes back to the file
        TargetWorkbook.Save();
    }

    public static void CopyDataRows(XLWorkbook TargetWorkbook, XLWorkbook RawDataWorkbook, int firstEmptyRow, string FilterValue, int FilterColumnNumber ) 
    {
        //Funcion para borrar todas las filas excepto la primera
        // Open the Excel file
        // Select the export worksheet
        var TargetWorksheet =  TargetWorkbook.Worksheet("Export");

        var RawDataWorksheet = RawDataWorkbook.Worksheet("Part1");

        
        foreach (var row in RawDataWorksheet.RowsUsed()){ // Usar cada fila del source

            var cellValue = row.Cell(FilterColumnNumber).GetValue<string>(); // get the column filter value
            
            if (cellValue == FilterValue){ // Check that it meets filter
                    
                    CopyRow(row, TargetWorksheet.Row(firstEmptyRow)); // Copy row from source to target next row
                    firstEmptyRow++;// Increase 
                }
        }

        // Save the changes back to the file
        TargetWorkbook.Save();
    }


}

