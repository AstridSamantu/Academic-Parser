using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using System.IO;
using System.Text;

class Disciplina
{
    public string Cod { get; set; }
    public string Nume { get; set; }
    public int Credite { get; set; }
    public string FormaEvaluare { get; set; }
    public int OreCurs { get; set; }
    public int OreSeminar { get; set; }
    public int OreLaborator { get; set; }
    public int OreProiect { get; set; }
    public int OrePractica { get; set; }
    public string Categorie { get; set; }
    public int OrePregatireIndividuala { get; set; }

    public int An { get; set; }
    public int Semestru { get; set; }
}

class Program
{
    static void Main()
    {
        // Lista in care salvam disciplinele
        List<Disciplina> discipline = new List<Disciplina>();

        // Lista coduri parcurse
        List<String> coduri = new List<String>();

        // Calea catre fisierul Excel
        string filePath = @"2023-2027_AC_PI_C-RO.xlsx";

        // Deschidem fisierul Excel
        using (var workbook = new XLWorkbook(filePath))
        {
            // Selectam foaia de lucru
            var worksheet = workbook.Worksheet("PLANURI");

            int lastRow = worksheet.LastRowUsed().RowNumber();
            int lastColumn = worksheet.LastColumnUsed().ColumnNumber();
            int semestru = 0;
            int an = 0;
            // Iteram coloanele
            for (int col = 1; col <= lastColumn; col++)
            {
                for (int row = 1; row <= lastRow; row++)
                {
                    var cell = worksheet.Cell(row, col);
                    if (cell.IsEmpty())
                        continue;
                    string text = cell.GetValue<string>();
                    // Aflam anul si semestrul
                    if (text.Contains("SEMESTRUL 1"))
                    {
                        semestru = 1;
                        an = 1;
                    }
                    if (text.Contains("SEMESTRUL 2"))
                    {
                        semestru = 2;
                        an = 1;
                    }
                    if (text.Contains("SEMESTRUL 3"))
                    {
                        semestru = 3;
                        an = 2;
                    }
                    if (text.Contains("SEMESTRUL 4"))
                    {
                        semestru = 4;
                        an = 2;
                    }
                    if (text.Contains("SEMESTRUL 5"))
                    {
                        semestru = 5;
                        an = 3;
                    }
                    if (text.Contains("SEMESTRUL 6"))
                    {
                        semestru = 6;
                        an = 3;
                    }
                    if (text.Contains("SEMESTRUL 7"))
                    {
                        semestru = 7;
                        an = 4;
                    }

                    if (text.Contains("SEMESTRUL 8"))
                    {
                        semestru = 8;
                        an = 4;
                    }

          
                    // Gasim cod si Extragem datele
                    if (text.StartsWith("L00") && !text.EndsWith("-ij"))
                    {
                        // Daca gaseste acelasi cod, sare peste
                        if (coduri.Contains(text)) continue;
                        coduri.Add(text);
                    
                        Disciplina d = new Disciplina();
                        d.Cod = text;
                        d.Semestru = semestru;
                        d.An = an;
                        d.Nume = worksheet.Cell(row - 2, col).GetValue<string>();
                        if (worksheet.Cell(row, col + 3).TryGetValue<int>(out int credite)) d.Credite = credite;
                        d.FormaEvaluare = worksheet.Cell(row, col + 4).GetValue<string>();
                        if (worksheet.Cell(row, col + 5).TryGetValue<int>(out int curs)) d.OreCurs = curs;
                        if (worksheet.Cell(row, col + 6).TryGetValue<int>(out int sem)) d.OreSeminar = sem;
                        if (worksheet.Cell(row, col + 7).TryGetValue<int>(out int lab)) d.OreLaborator = lab;
                        if (worksheet.Cell(row, col + 8).TryGetValue<int>(out int pr)) d.OreProiect = pr;
                        if (worksheet.Cell(row, col + 9).TryGetValue<int>(out int prac)) d.OrePractica = prac;
                        d.Categorie = worksheet.Cell(row, col + 10).GetValue<string>();
                        if (worksheet.Cell(row, col + 11).TryGetValue<int>(out int preg)) d.OrePregatireIndividuala = preg;
                        discipline.Add(d);
                    }
                }
            }
            // Afisare
            Console.WriteLine($"Am extras {discipline.Count()} discipline.");
            foreach (var d in discipline)
            {
                Console.WriteLine($"COD: {d.Cod} | NUME: {d.Nume} | CREDITE: {d.Credite} | FORMA DE EVALUARE: {d.FormaEvaluare} | ORE CURS: {d.OreCurs} | ORE SEMINAR: {d.OreSeminar} | ORE LABORATOR: {d.OreLaborator} | ORE PROIECT: {d.OreProiect} | ORE PRACTICA: {d.OrePractica} | CATEGORIE: {d.Categorie} | ORE PREGATIRE INDIVIDUALA: {d.OrePregatireIndividuala} | SEMESTRU: {d.Semestru} | AN: {d.An}");
                Console.WriteLine();
            }

            // Creare Excel
            // Cream un nou Workbook pentru export
            using (var exportWorkbook = new XLWorkbook())
            {
                var sheet = exportWorkbook.Worksheets.Add("Discipline Extrase");

                // Scriem Header-ul (Capul de tabel)
                sheet.Cell(1, 1).Value = "COD";
                sheet.Cell(1, 2).Value = "NUME";
                sheet.Cell(1, 3).Value = "CREDITE";
                sheet.Cell(1, 4).Value = "FORMA EVALUARE";
                sheet.Cell(1, 5).Value = "ORE CURS";
                sheet.Cell(1, 6).Value = "ORE SEMINAR";
                sheet.Cell(1, 7).Value = "ORE LABORATOR";
                sheet.Cell(1, 8).Value = "ORE PROIECT";
                sheet.Cell(1, 9).Value = "ORE PRACTICA";
                sheet.Cell(1, 10).Value = "CATEGORIE";
                sheet.Cell(1, 11).Value = "PREG. INDIV.";
                sheet.Cell(1, 12).Value = "AN";
                sheet.Cell(1, 13).Value = "SEMESTRU";

                // Formatam header-ul (Bold și fundal gri)
                var headerRow = sheet.Row(1);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                // Populam tabelul cu datele din lista 'discipline'
                for (int i = 0; i < discipline.Count; i++)
                {
                    var d = discipline[i];
                    int currentRow = i + 2; // Incepem de la randul 2 (1 e header-ul)

                    sheet.Cell(currentRow, 1).Value = d.Cod;
                    sheet.Cell(currentRow, 2).Value = d.Nume;
                    sheet.Cell(currentRow, 3).Value = d.Credite;
                    sheet.Cell(currentRow, 4).Value = d.FormaEvaluare;
                    sheet.Cell(currentRow, 5).Value = d.OreCurs;
                    sheet.Cell(currentRow, 6).Value = d.OreSeminar;
                    sheet.Cell(currentRow, 7).Value = d.OreLaborator;
                    sheet.Cell(currentRow, 8).Value = d.OreProiect;
                    sheet.Cell(currentRow, 9).Value = d.OrePractica;
                    sheet.Cell(currentRow, 10).Value = d.Categorie;
                    sheet.Cell(currentRow, 11).Value = d.OrePregatireIndividuala;
                    sheet.Cell(currentRow, 12).Value = d.An;
                    sheet.Cell(currentRow, 13).Value = d.Semestru;
                }

                // Ajustam automat latimea coloanelor
                sheet.Columns().AdjustToContents();

                // Salvam fisierul
                string exportPath = "Discipline_Extrase_C-RO.xlsx";
                exportWorkbook.SaveAs(exportPath);

                Console.WriteLine($"Succes! Datele au fost salvate in: {exportPath}");
            }

            // Cream CSV
            string csvPath = "Discipline_Extrase_C-RO.csv";

            // Folosim StreamWriter pentru a scrie text în fisier
            // Encoding.UTF8 asigura ca diacriticele sunt salvate corect
            using (var writer = new System.IO.StreamWriter(csvPath, false, System.Text.Encoding.UTF8))
            {
                // Scriem Header-ul
                writer.WriteLine("Cod;Nume;Credite;FormaEvaluare;OreCurs;OreSeminar;OreLaborator;OreProiect;OrePractica;Categorie;OrePregatireIndividuala;An;Semestru");

                // Parcurgem lista și scriem fiecare rand
                foreach (var d in discipline)
                {
                    // Folosim separatorul ';'
                    string line = $"{d.Cod};{d.Nume};{d.Credite};{d.FormaEvaluare};{d.OreCurs};{d.OreLaborator};{d.OreProiect};{d.OrePractica};{d.Categorie};{d.OrePregatireIndividuala};{d.An};{d.Semestru}";
                    writer.WriteLine(line);
                }
            }
            Console.WriteLine($"Succes! Datele au fost salvate si in CSV: {csvPath}");
        }
    }
}

