namespace Helpers
{
    public partial class ProjectIDNumberListClass
    {
        public string GetStringListData(string prefix)
        {
            string result = "";
            
            for (int i = 0; i < List.Count; i++)
                result +=
                    prefix
                    + "ProjectID/Number["
                    + i.ToString()
                    + "] "
                    + List[i].ProjectID.ToString()
                    + "/"
                    + List[i].ProjectNumber.ToString()
                    ;

            return result;
        }

        public string GetStringListData()
        {
            return GetStringListData("\n");
        }
    }
}
