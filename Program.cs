// See https://aka.ms/new-console-template for more information
using ConsoleToGoogleSheets;

Console.WriteLine("Hello, World!");
SheetsAccess sheetsAccess = new SheetsAccess();

sheetsAccess.ReadEntrys();
sheetsAccess.CreateEntry();


