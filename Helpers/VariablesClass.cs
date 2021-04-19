using System;
using System.Data;

namespace Helpers
{
    public class VariablesClass
    {
        public string UserMessage { get; set; }

        public void Refresh()
        {
            foreach (var property in this.GetType().GetProperties())
            {
                switch (property.PropertyType.Name)
                {
                    case "String": property.SetValue(this, ""); break;
                    case "DataTable": property.SetValue(this, new DataTable()); break;
                    case "Int32": property.SetValue(this, 0); break;
                    case "Double": property.SetValue(this, 0); break;
                    case "CellStyleClass": property.SetValue(this, new Helpers.Excel.CellStyleClass()); break;
                    case "Boolean": property.SetValue(this, false); break;
                    default: throw new Exception("Ошибка сброса настроек: " + property.Name + ", тип: " + property.PropertyType.Name);
                }
            }
        }

    }
}
