using System;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection;

class Program
{
    const string defaultUserName = "User";
    static void Main()
    {
        string answer;
        string homepage;
        Console.WriteLine("Welcome To Nottingham Books and Library Information System");
        Console.Write("\nEnter your name: ");
        string name = Console.ReadLine();

        if (name == "")
        {
            Console.WriteLine("No name entered. Assigning default name " + defaultUserName + " instead");
            name = defaultUserName;
        }

        do
        {
            Console.WriteLine("\nHello " + name + " " + "\n");
            string[] menue = { "1. View Libraries And Bookstores", "2. Where To Find?", "3. Goodreads", "4. Personnel Only" };
            for (int i = 0; i < menue.Length; ++i)
            {
                Console.WriteLine("\t" + menue[i]);
            }
            Console.Write("\nPlease choose one of the above numbers [1-4]: ");

            int Choice = int.Parse(Console.ReadLine());
            if (Choice == 1)
            {
                //code for the do-while loop was taken from:
                //Elton sampaio, Youtube, 2021
                //Simple menu in c# console application
                //https://youtube.com/watch?v=byqyLO8sQpI&feature=share
                //[Access Date: 02.01.2023] 
                bool finished = false;
                do
                {
                    do
                    {
                        Console.WriteLine("\n\nView Libraries And Bookstores");
                        Console.WriteLine("\nChoose one of the following:");
                        string[] librarylist = { "1. Hyson Green Library", "2. Radford Lenton Library", "3. Aspley Library", "4. Bilborough Library", "5. Wollaton Library", "6. Bromley House Library ", "7. Strelley Road Library", "8. St. Anns Library", "9. Boots Library ", "10. The Dales Centre Library ", "11. George Green Library ", "12. Mapperley Library ", "13. Hallward Library ", "14. Meadows Library ", "15. Arnold Library ", "16. Ilkeston Library ", "17. Waterstones", "18. Bookwise", "19. Five Leaves Bookshop" };
                        for (int i = 0; i < librarylist.Length; ++i)
                        {
                            Console.WriteLine("\t" + librarylist[i]);
                        }
                        Console.Write("\nPlease choose one of the following numbers for details: ");
                        int Choice1 = int.Parse(Console.ReadLine());

                        if (Choice1 == 1)
                        {
                            Console.WriteLine("\nHyson Green Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nHyson Green Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 2)
                        {
                            Console.WriteLine("\nRadford Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nRadford Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 915 2849");
                            Console.WriteLine("Email: radford_lenton.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Lenton Boulevard, Nottingham, NG7 2BY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/radford-lenton-library/");
                            finished = true;
                        }

                        else if (Choice1 == 3)
                        {
                            Console.WriteLine("\nAspley Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nAspley Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 915 2802");
                            Console.WriteLine("Email: aspley.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Nuthall Road, Nottingham, NG8 5DD");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/aspley-library/");
                            finished = true;
                        }

                        else if (Choice1 == 4)
                        {
                            Console.WriteLine("\nBilborough Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nBilborough Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 915 2820");
                            Console.WriteLine("Email: bilborough.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Bracebridge Drive, Nottingham NG8 4PN");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/bilborough-library/");
                            finished = true;
                        }

                        else if (Choice1 == 5)
                        {
                            Console.WriteLine("\nWollaton Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nWollaton Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 915 2809");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Wollaton Library, Bramcote Lane, Wollaton, Nottingham, NG8 2NA");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/wollaton-library/");
                            finished = true;
                        }

                        else if (Choice1 == 6)
                        {
                            Console.WriteLine("\nBromley House Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nBromley House Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 7)
                        {
                            Console.WriteLine("\nStrelley Road Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nStrelley Road Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 8)
                        {
                            Console.WriteLine("\nSt. Anns Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nSt. Anns Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 9)
                        {
                            Console.WriteLine("\nBoots Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nBoots Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 10)
                        {
                            Console.WriteLine("\nThe Dales Centre Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nThe Dales Centre Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 11)
                        {
                            Console.WriteLine("\nGeorge Green Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nGeorge Green Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 12)
                        {
                            Console.WriteLine("\nMapperley Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nMapperley Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 13)
                        {
                            Console.WriteLine("\nHallward Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nHallward Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 14)
                        {
                            Console.WriteLine("\nMeadows Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nMeadows Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 15)
                        {
                            Console.WriteLine("\nArnold Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nArnold Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 16)
                        {
                            Console.WriteLine("\nIlkeston Library");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nIlkeston Library is a Nottingham City Public Library.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 17)
                        {
                            Console.WriteLine("\nWaterstones");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nWaterstones is one of the biggest bookstore in Nottingham, including reading areas and a cafe.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 18)
                        {
                            Console.WriteLine("\nBookwise");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nBookwise is a charity bookstore. All the profits from this shop are donated to Music For Everyone.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else if (Choice1 == 19)
                        {
                            Console.WriteLine("\nFive Leaves Bookshop");
                            Console.WriteLine("\nAbout:");
                            Console.WriteLine("\nFive Leaves Bookshop.");
                            Console.WriteLine("Tel: 0115 883 8332");
                            Console.WriteLine("Email: hyson_green.library@nottinghamcity.gov.uk");
                            Console.WriteLine("Address: Hyson Green Library, The Mary Potter Centre, 76 Gregory Boulevard, Nottingham, NG7 5HY");
                            Console.WriteLine("For more information, refer to our official website" + "\n https://www.nottinghamcitylibraries.co.uk/library/hyson-green-library-mary-potter-centre/");
                            finished = true;
                        }

                        else
                        {

                            Console.Write("\nOption not available.");

                            finished = false;
                        }

                    } while (!finished);

                    do
                    {
                        Console.Write("\n\nWould you like to search another library? [y/n]): ");
                        answer = Console.ReadLine().ToLower();


                    }
                    while (answer != "y" && answer != "n" && answer != "yes" && answer != "no" && answer != "yep");

                } while (answer == "y" || answer == "yes");

            }

            else if (Choice == 2)
            {
                Console.WriteLine("\nWhere to Find?");
                Console.WriteLine("");
                readsheet();
            }

            else if (Choice == 3)
            {
                Console.WriteLine("\nGoodreads");
                Console.WriteLine("");

                string[] txtfiles = { "1. New Releases Of The Month", "2. Book Of The Year" };
                for (int i = 0; i < txtfiles.Length; ++i)
                {
                    Console.WriteLine("\t" + txtfiles[i]);
                }
                Console.Write("\nPlease choose one of the above numbers [1-2]: ");

                int Goodreads = int.Parse(Console.ReadLine());


                while (Goodreads > 2)
                {
                    Console.Write("Invalid Option.");
                    break;
                }
                if (Goodreads == 1)
                {
                    //gr = Goodreads
                    string[] gr = { "1. February 2023", "2. January 2023", "3. November 2022" };
                    for (int i = 0; i < gr.Length; ++i)
                    {
                        Console.WriteLine("\t" + gr[i]);
                    }
                    Console.Write("\nPlease choose one of the above numbers [1-3]: ");
                    //BOTM = Book Of The Month
                    int BOTM = int.Parse(Console.ReadLine());
                    while (BOTM > 2)
                    {
                        Console.Write("Invalid Option.");
                        break;
                    }
                    if (BOTM == 1)
                    {
                        Console.Clear();

                        //Code to read .txt file taken from:
                        //Microsoft Learn, 2021
                        //How to read from a text file (C# Programming Guide)
                        //https://learn.microsoft.com/en-us/dotnet/csharp/programming-guide/file-system/how-to-read-from-a-text-file
                        ////[Accessed Date: 09.01.2023]
                        ///
                        string[] document = System.IO.File.ReadAllLines(@"E:\Visual Codes\Library Management\FEBRUARY 2023.txt");
                        System.Console.WriteLine("Contents of WriteLines2.txt = ");
                        foreach (string details in document)
                        {
                            Console.WriteLine("\t" + details);
                        }
                    }
                    else if (BOTM == 2)
                    {
                        Console.Clear();
                        string[] document = System.IO.File.ReadAllLines(@"E:\Visual Codes\Library Management\JANUARY 2023.txt");
                        System.Console.WriteLine("Contents of WriteLines2.txt = ");
                        foreach (string details in document)
                        {
                            Console.WriteLine("\t" + details);
                        }
                    }
                    else if (BOTM == 3)
                    {
                        Console.Clear();
                        string[] document = System.IO.File.ReadAllLines(@"E:\Visual Codes\Library Management\DECEMBER 2022.txt");
                        System.Console.WriteLine("Contents of WriteLines2.txt = ");
                        foreach (string details in document)
                        {
                            Console.WriteLine("\t" + details);
                        }
                    }

                    else { }
                }

                else if (Goodreads == 2)
                {
                    string[] lines = System.IO.File.ReadAllLines(@"E:\Visual Codes\Library Management\2022.txt");
                    System.Console.WriteLine("Contents of WriteLines2.txt = ");
                    foreach (string line in lines)
                    {
                        Console.WriteLine("\t" + line);
                    }
                }

                else { }

            }

            else if (Choice == 4)
            {
                Console.WriteLine("\nPersonnel Only");
                Console.WriteLine("");
                string password = "bibliography";
                Console.Write("Enter your password: ");
                string enteredpassword = Console.ReadLine();

                while (enteredpassword != password)
                {
                    Console.Write("Incorrect Password. Try Again: ");
                    enteredpassword = Console.ReadLine();
                }

                string[] sheets = { "1. Create  A New Sheet", "2. Add Data To Existing Sheet", "3. Delete Data From Existing Sheet" };
                for (int i = 0; i < sheets.Length; ++i)
                {
                    Console.WriteLine("\t" + sheets[i]);
                }
                Console.Write("\nPlease choose one of the above numbers [1-3]: ");

                int ChoseSheet = int.Parse(Console.ReadLine());
                if (ChoseSheet == 1)
                {
                    createsheet();
                }

                else if (ChoseSheet == 2)
                {
                    addtosheet();
                }

                else if (ChoseSheet == 3)
                {
                    deletefromsheet();
                }

                else { }
            }

            else { }

            Console.Write("\n\nWould you like to go back to homepage or quit? [h/q]): ");
            homepage = Console.ReadLine().ToLower();
        }
        while (homepage == "h" || homepage == "home");
    }

    static void readsheet()
    {
        //Code to read excel files taken from this webpage
        //Mike, C#,Windows Form, WPF, LINQ, Entity Framework Examples and Codes, 2018
        //Reading Excel file in C# Console Application
        //https://www.csharp-console-examples.com/general/reading-excel-file-in-c-console-application/
        //[Accessed Date: 08.01.2023]

        Excel.Application bookxls = new Microsoft.Office.Interop.Excel.Application();
        Workbook excelBook = bookxls.Workbooks.Open(@"E:\Visual Codes\Library ManagementBooklist.xls");
        _Worksheet excelSheet = excelBook.Sheets[1];
        Range excelRange = excelSheet.UsedRange;


        int rows = excelRange.Rows.Count;
        int cols = excelRange.Columns.Count;

        for (int i = 1; i <= rows; i++)
        {
            Console.Write("\r\n");
            for (int j = 1; j <= cols; j++)
            {

                if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                    Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
            }
        }
        bookxls.Quit();
    }
    public class NewBooks
    {
        public string Book { get; set; }
        public string Author { get; set; }
        public double Price { get; set; }
        public string Availability { get; set; }

    }

    //Codes to create a new xls file, add to xls file and delete from xls file taken from:
    //Ramakrishna Basagalla, C# Corner, 2018
    //CRUD In Excel File In C#
    //https://www.c-sharpcorner.com/article/create-update-delete-and-reading-the-excel-file-in-c-sharp/
    //[Accessed Date: 08.01.2023]
    static void createsheet()
    {
        string fileName;
        Console.Write("Enter File Name :");
        fileName = Console.ReadLine();

        Excel.Application bookxls = new Microsoft.Office.Interop.Excel.Application();
        Excel.Workbook newxlsbook;
        Excel.Worksheet newxlsheet;
        object misValue = System.Reflection.Missing.Value;

        //Format to add cells and texts
        newxlsbook = bookxls.Workbooks.Add(misValue);
        newxlsheet = (Excel.Worksheet)newxlsbook.Worksheets.get_Item(1);
        newxlsheet.Cells[1, 1] = "Book";
        newxlsheet.Cells[1, 2] = "Author";
        newxlsheet.Cells[1, 3] = "Price (£)";
        newxlsheet.Cells[1, 4] = "Availability";
        /*newxlsheet.Cells[2, 1] = "";
        newxlsheet.Cells[2, 2] = "";*/

        string location = @"E:\Visual Codes\" + fileName + ".xls";//Dont forget, you have to add to exist location
        newxlsbook.SaveAs(location, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        newxlsbook.Close(true, misValue, misValue);
        bookxls.Quit();

    }
    static void addtosheet()
    {
        string location = @"E:\Visual Codes\Library Management\Booklist.xls";

        IList<NewBooks> booklist = new List<NewBooks>()
        {
        new NewBooks(){ Book="Mix Tape", Author="Jane Sanderson", Price=3.99, Availability="waterstones"},
        new NewBooks(){ Book="Dear Emmie Blue", Author="Lia Louis", Price=8.59, Availability="Waterstones"},
        new NewBooks(){ Book="And Then There Were None", Author="Agatha Christie", Price=1.00, Availability="Bookwise" }
        };

        Excel.Application bookxls = new Excel.Application();

        Excel.Workbook newxlsbook = bookxls.Workbooks.Open(location, 0, false, 5, "", "", false,
        Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        Excel.Worksheet newxlsheet = (Excel.Worksheet)newxlsbook.Worksheets.get_Item(1);

        Excel.Range xlRange = newxlsheet.UsedRange;
        int rowNumber = xlRange.Rows.Count + 1;

        foreach (NewBooks @new in booklist)
        {
            newxlsheet.Cells[rowNumber, 1] = @new.Book;
            newxlsheet.Cells[rowNumber, 2] = @new.Author;
            newxlsheet.Cells[rowNumber, 3] = @new.Price;
            newxlsheet.Cells[rowNumber, 4] = @new.Availability;
            rowNumber++;
        }

        bookxls.DisplayAlerts = false;
        newxlsbook.SaveAs(location, Excel.XlFileFormat.xlOpenXMLWorkbook,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
            Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
        newxlsbook.Close();
        bookxls.Quit();
    }
    static void deletefromsheet()
    {
        string location = @"E:\Visual Codes\Library Management\Booklist.xls";

        Excel.Application bookxls = new Excel.Application();

        Excel.Workbook oldxlsbook = bookxls.Workbooks.Open(location, 0, false, 5, "", "", false,
            Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
        Excel.Worksheet oldxlsheet = (Excel.Worksheet)oldxlsbook.Worksheets.get_Item(1);

        Excel.Range range1 = oldxlsheet.get_Range("A2", "B2");

        range1.EntireRow.Delete(Type.Missing);

        Excel.Range range2 = oldxlsheet.get_Range("B3", "B3");
        range2.Cells.Clear();


        bookxls.DisplayAlerts = false;
        oldxlsbook.SaveAs(location, Excel.XlFileFormat.xlOpenXMLWorkbook,
            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange,
            Excel.XlSaveConflictResolution.xlLocalSessionChanges, Missing.Value, Missing.Value,
            Missing.Value, Missing.Value);
        oldxlsbook.Close();
        bookxls.Quit();
    }
}