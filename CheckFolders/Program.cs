using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace CheckFolders
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine('\n');
            Console.WriteLine(@"Wybierz operację do wykonania:");
            Console.WriteLine('\n');
            Console.WriteLine(@"1 - Weryfikacja plików XML i WKT");
            Console.WriteLine(@"2 - Statystyka plików WKT dla skanów");

            ConsoleKeyInfo keyPressed = Console.ReadKey(true);

            string startPath = args.Length == 0 ? AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\') : args[0].TrimEnd('\\');

            Console.WriteLine("\nPobieranie listy folerów...");

            string[] directories = Directory.GetDirectories(startPath, "*", SearchOption.AllDirectories);

            Console.WriteLine($@"Pobrano {directories.Length} folderów.");

            Console.WriteLine(@"Analizowanie folderów...");

            Array.Sort(directories, new NaturalStringComparer()); 

            List<string> operatDirectory = (from dir in directories
                let directory = new DirectoryInfo(dir)
                let subdirs = directory.GetDirectories()
                where subdirs.Length == 0
                select dir).ToList();

            Console.WriteLine($@"Koniec analizy folderów. Pozostało {operatDirectory.Count} folderów.");

            switch (keyPressed.KeyChar)
            {
                case '1':

                    //  -----------------------------------------------------------------------------------
                    //  foldery bez skanów

                    Console.WriteLine("\nWeryfikacja czy wszystkie foldery mają pliki PDF...");

                    List<Operat> operatSkanyTest = new List<Operat>();

                    foreach (string dir in from dir in operatDirectory
                        let fileNames = Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly)
                        where fileNames.Length == 0
                        select dir)
                    {
                        Operat operat = new Operat
                        {
                            FullPath = dir,
                            OperatName = dir.Split(Path.DirectorySeparatorChar).Last(),
                            Info = @"Brak plików PDF w katalogu"
                        };

                        operatSkanyTest.Add(operat);
                    }

                    Console.WriteLine($@"Liczba błędów: {operatSkanyTest.Count}");

                    //  -----------------------------------------------------------------------------------
                    //  Czy w każdym katalogu jest plik XML dla operatu i czy jego nazwa jest poprawna

                    Console.WriteLine("\nWeryfikacja czy wszystkie operaty posiadają plik opisowy XML...");

                    List<Operat> operatXmlTest = new List<Operat>();

                    foreach (string dir in from dir in operatDirectory
                        let operatName = dir.Split(Path.DirectorySeparatorChar).Last()
                        let fileNames = Directory.GetFiles(dir, operatName + ".xml", SearchOption.TopDirectoryOnly)
                        where fileNames.Length == 0
                        select dir)
                    {
                        Operat operat = new Operat
                        {
                            FullPath = dir,
                            OperatName = dir.Split(Path.DirectorySeparatorChar).Last(),
                            Info = @"Brak pliku XML dla operatu"
                        };

                        operatXmlTest.Add(operat);
                    }

                    Console.WriteLine($@"Liczba błędów: {operatXmlTest.Count}");

                    //  -----------------------------------------------------------------------------------
                    //  Czy w każdym katalogu jest plik WKT dla operatu i czy jego nazwa jest poprawna

                    Console.WriteLine("\nWeryfikacja czy wszystkie operaty posiadają plik z zakresem WKT...");

                    List<Operat> operatWktTest = new List<Operat>();

                    foreach (string dir in from dir in operatDirectory
                        let operatName = dir.Split(Path.DirectorySeparatorChar).Last()
                        let fileNames = Directory.GetFiles(dir, operatName + ".wkt", SearchOption.TopDirectoryOnly)
                        where fileNames.Length == 0
                        select dir)
                    {
                        Operat operat = new Operat
                        {
                            FullPath = dir,
                            OperatName = dir.Split(Path.DirectorySeparatorChar).Last(),
                            Info = @"Brak pliku WKT dla operatu"
                        };

                        operatWktTest.Add(operat);
                    }

                    Console.WriteLine($@"Liczba błędów: {operatWktTest.Count}");

                    //  -----------------------------------------------------------------------------------
                    //  Weryfikacja poprawności nazw plików WKT dla szkiców

                    Console.WriteLine("\nWeryfikacja poprawności nazw plików WKT dla szkiców...");

                    List<Operat> operatWktNazwaTest = new List<Operat>();

                    foreach (string dir in operatDirectory)
                    {
                        List<string> wktFiles = Directory.GetFiles(dir, "*.wkt", SearchOption.TopDirectoryOnly).ToList();
                        List<string> pdfFiles = Directory.GetFiles(dir, "*.pdf", SearchOption.TopDirectoryOnly).ToList();
                        List<string> xmlFiles = Directory.GetFiles(dir, "*.xml", SearchOption.TopDirectoryOnly).ToList();

                        for (int i = 0; i < wktFiles.Count; i++)
                        {
                            wktFiles[i] = Path.GetFileNameWithoutExtension(wktFiles[i]);
                        }

                        for (int i = 0; i < pdfFiles.Count; i++)
                        {
                            pdfFiles[i] = Path.GetFileNameWithoutExtension(pdfFiles[i]);
                        }

                        for (int i = 0; i < xmlFiles.Count; i++)
                        {
                            xmlFiles[i] = Path.GetFileNameWithoutExtension(xmlFiles[i]);
                        }

                        List<string> searchIn = new List<string>();
                        searchIn.AddRange(pdfFiles);
                        searchIn.AddRange(xmlFiles);

                        List<string> errors = wktFiles.Except(searchIn).ToList();

                        operatWktNazwaTest.AddRange(errors.Select(err => new Operat
                        {
                            FullPath = dir, 
                            OperatName = dir.Split(Path.DirectorySeparatorChar).Last(), 
                            Info = err + ".wkt"
                        }));
                    }

                    Console.WriteLine($@"Liczba błędów: {operatWktNazwaTest.Count}");

                    //  -----------------------------------------------------------------------------------
                    //  Czy pliki są zgodne z nazwą katalogu

                    Console.WriteLine("\nWeryfikacja zgodności nazwy plików z nazwą operatu...");

                    List<Operat> operatNazwaPlikiTest = new List<Operat>();

                    foreach (string dir in operatDirectory)
                    {
                        string operatName = dir.Split(Path.DirectorySeparatorChar).Last();

                        List<string> files = Directory.GetFiles(dir, "*.*", SearchOption.TopDirectoryOnly).ToList();

                        for (int i = 0; i < files.Count; i++)
                        {
                            files[i] = Path.GetFileName(files[i]);
                        }

                        foreach (string file in files)
                        {
                            if (!file.ToUpper().Equals(operatName + ".WKT") && 
                                !file.ToUpper().Equals(operatName + ".XML") &&
                                (!file.ToUpper().Contains(operatName + "-") || !file.ToUpper().EndsWith(".PDF")) &&
                                (!file.ToUpper().Contains(operatName + "-") || !file.ToUpper().EndsWith(".WKT")))
                            {
                                Operat operat = new Operat
                                {
                                    FullPath = dir,
                                    OperatName = dir.Split(Path.DirectorySeparatorChar).Last(),
                                    Info = file
                                };

                                operatNazwaPlikiTest.Add(operat);
                            }
                        }
                    }

                    Console.WriteLine($@"Liczba błędów: {operatNazwaPlikiTest.Count}");

                    //  -----------------------------------------------------------------------------------
                    //  Zapisywanie do XLS

                    string outputFile = Path.Combine(startPath, "wynik_weryfikacji.xlsx");

                    Console.WriteLine($"\nZapisywanie wyników do pliku {outputFile}...");

                    File.Delete(outputFile);

                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        //  -------------------------------------------------------------------------------
                        ExcelWorksheet arkusz = excelPackage.Workbook.Worksheets.Add("Foldery bez PDF");

                        arkusz.Cells[1, 1].Value = "Katalog";
                        arkusz.Cells[1, 2].Value = "Nazwa operatu";
                        arkusz.Cells[1, 3].Value = "Uwagi";

                        for (int i = 0; i < operatSkanyTest.Count; i++)
                        {
                            arkusz.Cells[i + 2, 1].Value = operatSkanyTest[i].FullPath;
                            arkusz.Cells[i + 2, 2].Value = operatSkanyTest[i].OperatName;
                            arkusz.Cells[i + 2, 3].Value = operatSkanyTest[i].Info;
                        }

                        arkusz.Cells["A:C"].AutoFilter = true;
                        arkusz.View.FreezePanes(2, 1);
                        arkusz.Cells.AutoFitColumns();

                        //  -------------------------------------------------------------------------------
                        arkusz = excelPackage.Workbook.Worksheets.Add("Operat bez XML");

                        arkusz.Cells[1, 1].Value = "Katalog";
                        arkusz.Cells[1, 2].Value = "Nazwa operatu";
                        arkusz.Cells[1, 3].Value = "Uwagi";

                        for (int i = 0; i < operatXmlTest.Count; i++)
                        {
                            arkusz.Cells[i + 2, 1].Value = operatXmlTest[i].FullPath;
                            arkusz.Cells[i + 2, 2].Value = operatXmlTest[i].OperatName;
                            arkusz.Cells[i + 2, 3].Value = operatXmlTest[i].Info;
                        }

                        arkusz.Cells["A:C"].AutoFilter = true;
                        arkusz.View.FreezePanes(2, 1);
                        arkusz.Cells.AutoFitColumns();

                        //  -------------------------------------------------------------------------------
                        arkusz = excelPackage.Workbook.Worksheets.Add("Operat bez WKT");

                        arkusz.Cells[1, 1].Value = "Katalog";
                        arkusz.Cells[1, 2].Value = "Nazwa operatu";
                        arkusz.Cells[1, 3].Value = "Uwagi";

                        for (int i = 0; i < operatWktTest.Count; i++)
                        {
                            arkusz.Cells[i + 2, 1].Value = operatWktTest[i].FullPath;
                            arkusz.Cells[i + 2, 2].Value = operatWktTest[i].OperatName;
                            arkusz.Cells[i + 2, 3].Value = operatWktTest[i].Info;
                        }

                        arkusz.Cells["A:C"].AutoFilter = true;
                        arkusz.View.FreezePanes(2, 1);
                        arkusz.Cells.AutoFitColumns();

                        //  -------------------------------------------------------------------------------
                        arkusz = excelPackage.Workbook.Worksheets.Add("WKT bez PDF");

                        arkusz.Cells[1, 1].Value = "Katalog";
                        arkusz.Cells[1, 2].Value = "Nazwa operatu";
                        arkusz.Cells[1, 3].Value = "Uwagi";

                        for (int i = 0; i < operatWktNazwaTest.Count; i++)
                        {
                            arkusz.Cells[i + 2, 1].Value = operatWktNazwaTest[i].FullPath;
                            arkusz.Cells[i + 2, 2].Value = operatWktNazwaTest[i].OperatName;
                            arkusz.Cells[i + 2, 3].Value = operatWktNazwaTest[i].Info;
                        }

                        arkusz.Cells["A:C"].AutoFilter = true;
                        arkusz.View.FreezePanes(2, 1);
                        arkusz.Cells.AutoFitColumns();

                        //  -------------------------------------------------------------------------------
                        arkusz = excelPackage.Workbook.Worksheets.Add("Złe nazwy plików");

                        arkusz.Cells[1, 1].Value = "Katalog";
                        arkusz.Cells[1, 2].Value = "Nazwa operatu";
                        arkusz.Cells[1, 3].Value = "Uwagi";

                        for (int i = 0; i < operatNazwaPlikiTest.Count; i++)
                        {
                            arkusz.Cells[i + 2, 1].Value = operatNazwaPlikiTest[i].FullPath;
                            arkusz.Cells[i + 2, 2].Value = operatNazwaPlikiTest[i].OperatName;
                            arkusz.Cells[i + 2, 3].Value = operatNazwaPlikiTest[i].Info;
                        }

                        arkusz.Cells["A:C"].AutoFilter = true;
                        arkusz.View.FreezePanes(2, 1);
                        arkusz.Cells.AutoFitColumns();

                        excelPackage.SaveAs(new FileInfo(outputFile));
                    }

                    break;

                case '2':

                    Console.WriteLine("\nZliczanie plików WKT...");

                    List<Operat> operatWktCountList = new List<Operat>();

                    foreach (string dir in operatDirectory)
                    {
                        string operatName = dir.Split(Path.DirectorySeparatorChar).Last();

                        List<string> filesWkt = Directory.GetFiles(dir, "*.wkt", SearchOption.TopDirectoryOnly).ToList();
                        List<string> fileWktOperat = Directory.GetFiles(dir, operatName + ".wkt", SearchOption.TopDirectoryOnly).ToList();

                        filesWkt = filesWkt.Except(fileWktOperat).ToList();

                        Operat operat = new Operat
                        {
                            FullPath = dir,
                            OperatName = dir.Split(Path.DirectorySeparatorChar).Last(),
                            Info = filesWkt.Count.ToString()
                        };

                        operatWktCountList.Add(operat);
                    }

                    //  -----------------------------------------------------------------------------------
                    //  Zapisywanie do XLS

                    outputFile = Path.Combine(startPath, "statystyka_wkt.xlsx");

                    Console.WriteLine($"\nZapisywanie wyników do pliku {outputFile}...");

                    File.Delete(outputFile);

                    using (ExcelPackage excelPackage = new ExcelPackage())
                    {
                        //  -------------------------------------------------------------------------------
                        ExcelWorksheet arkusz = excelPackage.Workbook.Worksheets.Add("WKT");

                        arkusz.Cells[1, 1].Value = "Katalog";
                        arkusz.Cells[1, 2].Value = "Nazwa operatu";
                        arkusz.Cells[1, 3].Value = "Liczba WKT";

                        for (int i = 0; i < operatWktCountList.Count; i++)
                        {
                            arkusz.Cells[i + 2, 1].Value = operatWktCountList[i].FullPath;
                            arkusz.Cells[i + 2, 2].Value = operatWktCountList[i].OperatName;
                            arkusz.Cells[i + 2, 3].Value = Convert.ToInt32(operatWktCountList[i].Info);
                        }

                        arkusz.Cells["A:C"].AutoFilter = true;
                        arkusz.View.FreezePanes(2, 1);
                        arkusz.Cells.AutoFitColumns();

                        excelPackage.SaveAs(new FileInfo(outputFile));
                    }

                    break;
            }



            

            Console.WriteLine("\nWciśnij dowolny klawisz");
            Console.ReadKey();

        }
    }
}
