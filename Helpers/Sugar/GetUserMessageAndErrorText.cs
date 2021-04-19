namespace Helpers
{
    public static partial class Sugar
    {
        public static string GetUserMessageAndErrorText(string userMessage, string errorText, bool isGetErrorMessage)
        {
            if (!string.IsNullOrEmpty(userMessage))
                userMessage = "";

            if (!string.IsNullOrEmpty(errorText))
            {
                if (isGetErrorMessage)
                    userMessage += ((userMessage == "") ? System.Environment.NewLine : "") + errorText;
                else
                    throw new System.Exception(errorText);
            }
            return userMessage;
        }
    }
}
