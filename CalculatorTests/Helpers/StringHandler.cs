namespace Consinco.Helpers
{
    public class StringHandler
    {
        public static List<string> ParseStringToList(string inputString)
        {
            string[] parts = inputString.Split(',');
            List<string> itemList = new List<string>();

            foreach (string part in parts)
            {
                itemList.Add(part.Trim());
            }

            return itemList;
        }
    }
}
