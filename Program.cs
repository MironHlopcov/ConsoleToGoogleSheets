// See https://aka.ms/new-console-template for more information
using ConsoleToGoogleSheets;

Console.WriteLine("Hello, World!");
SheetsService sheetsAccess = new SheetsService("ConsoleToGoogleSheets", 
    "1HwvBHPkJjDxHJgBFLM82WV51N2AVZDU339Qw-nc6FYk", 
    "Table-1",
    "secure-answer.json"
    );

var items = sheetsAccess.ReadItems();
var response = await sheetsAccess.WriteItemsAsync(items.ToArray());

while (response == null)
{
    Console.Write("/");
}
Console.WriteLine("Finish");
//sheetsAccess.ReadEntrys();


//sheetsAccess.CreateEntry();


