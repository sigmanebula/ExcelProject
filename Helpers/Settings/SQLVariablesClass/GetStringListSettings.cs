namespace Helpers
{
    public partial class SQLVariablesClass
    {
        public string GetStringListSettings(string prefix)
        {
            string result = prefix + "Settings:";

            foreach (var property in this.GetType().GetProperties())
                result += prefix + property.Name + " " + (property.GetValue(this) ?? "").ToString();

            return result;
        }

        public string GetStringListSettings()
        {
            return GetStringListSettings("\n");
        }
    }
}