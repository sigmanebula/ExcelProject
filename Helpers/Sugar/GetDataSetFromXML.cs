namespace Helpers
{
    public static partial class Sugar
    {
        public static System.Data.DataSet GetDataSetFromXML(string xml, ref string errorText)
        {
            System.Data.DataSet dataSet = new System.Data.DataSet();

            if (errorText == "")
                try
                {
                    var stream = new System.IO.MemoryStream();
                    var writer = new System.IO.StreamWriter(stream);

                    writer.Write(xml);
                    writer.Flush();
                    stream.Position = 0;

                    dataSet.ReadXml(stream);
                }
                catch(System.Exception exception)
                {
                    errorText += "Создание датасета из К2 XML. Ошибка: " + exception.Message + System.Environment.NewLine;
                }

            return dataSet;
        }

        public static System.Data.DataSet GetDataSetFromXML(string xml)
        {
            string errorText = "";

            System.Data.DataSet dataSet = GetDataSetFromXML(xml, ref errorText);

            if (errorText != "")
                throw new System.Exception(errorText);
            
            return dataSet;
        }
    }
}
