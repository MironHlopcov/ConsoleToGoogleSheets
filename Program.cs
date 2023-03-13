// See https://aka.ms/new-console-template for more information
using ConsoleToGoogleSheets;
using ConsoleToGoogleSheets.Models;
using System.IO;
using System;

Console.WriteLine("Hello, World!");
SheetsService sheetsAccess = new SheetsService("ConsoleToGoogleSheets", 
    "1HwvBHPkJjDxHJgBFLM82WV51N2AVZDU339Qw-nc6FYk", 
    "Table-1",
    "secure-answer.json"
    );

var items = sheetsAccess.ReadItemsAsync();
var copyOfItems = items.GetRange(0, 20);


var newItem = new Item { EEID="E04625", FullName="Adam Dang", JobTitle="Director", Department="Sales", BusinessUnit="Research & Development", Gender="Male",
     Ethnicity="Asian", Age=4, HireDate= DateTime.Parse("07/09/2002"), AnnualSalary=166.331m, Bonus="18 %", Country="China", City="Chongqing" };
var newItem2 = new Item
{
    EEID = "E04626",
    FullName = "Adama Danga",
    JobTitle = "Directorka",
    Department = "Sales",
    BusinessUnit = "Research & Development",
   
    Age = 45,
    HireDate = DateTime.Parse("07/09/2002"),
    AnnualSalary = 166.331m,
    Bonus = "18 %",
    Country = "China",
    City = "Chongqing"
};
//copyOfItems.Clear();
copyOfItems.Add(newItem);
copyOfItems.Add(newItem2);
var response = sheetsAccess.AddItemsAsync(copyOfItems.ToArray());

int i = 0;
while (response.Status != TaskStatus.RanToCompletion)
{
    Console.CursorVisible=false;
    await Task.Delay(100);
    var position = Console.GetCursorPosition();
    switch (i)
    {
            case 0 : Console.Write("|");
            i++;
            break;
            case 1 : Console.Write("/");
            i++;
            break;
            case 2 : Console.Write("-");
            i++;
            break;
            case 3 : Console.Write("\\");
            i = 0;
            break;
            default : i=0;
            break;
    }
    Console.SetCursorPosition(position.Left,position.Top);
    Console.CursorVisible = true;
}
Console.WriteLine("Finish");
//sheetsAccess.ReadEntrys();


//sheetsAccess.CreateEntry();


