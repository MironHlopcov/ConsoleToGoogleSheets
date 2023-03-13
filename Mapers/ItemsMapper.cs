using ConsoleToGoogleSheets.Models;
using Google.Apis.Sheets.v4.Data;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleToGoogleSheets.Mapers
{
    public static class ItemsMapper
    {
        public static List<Item> MapFromRangeData(IList<IList<object>> values)
        {
            var items = new List<Item>();
            foreach (var value in values)
            {
                Item item  = new Item();
                Type typeItem = item.GetType();
                var propertis = typeItem.GetProperties();
                for (int i = 0; i < propertis.Length; i++)
                {
                    var prop = propertis[i];
                    if (prop != null && value.Count - 1 >= i)
                    {
                        var typezedVamue = prop.PropertyType.Name switch
                        {
                            nameof(DateTime) =>value[i].ToString() =="" ? null: DateTime.Parse(value[i].ToString()),
                            nameof(Int32)=> value[i].ToString() == "" ? null : Int32.Parse(value[i].ToString()),
                            nameof(Decimal)=>value[i].ToString() == "" ? null : Decimal.Parse(value[i].ToString()),
                            _ => value[i]
                        };
                        //prop.SetValue(item, value[i], null);
                        prop.SetValue(item, typezedVamue, null);
                    }
                }
                items.Add(item);
            }
            return items;
        }
        public static ValueRange MapToRangeData(Item[] items)
        {
            if (items.Length > 0)
            {
                Type typeItem = items.First().GetType();
                var propertis = typeItem.GetProperties();
                ValueRange rangeData = new();
                rangeData.Values = new List<IList<object>>();
                foreach (Item it in items)
                {
                    var objectList = new List<object>();
                    for (int i = 0; i < propertis.Length; i++)
                    {
                        var prop = propertis[i];
                        var value = prop.GetValue(it);
                        objectList.Add(value);
                    }
                    rangeData.Values.Add(objectList);
                }
                return rangeData;
            }
            else
                return null;
        }
    }
}
