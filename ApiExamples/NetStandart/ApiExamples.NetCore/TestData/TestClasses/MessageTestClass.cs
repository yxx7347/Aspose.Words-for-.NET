namespace ApiExamples.NetCore.TestData.TestClasses
{
    public class MessageTestClass
    {
        public string Name { get; set; }
        public string Message { get; set; }

        public MessageTestClass(string name, string message)
        {
            this.Name = name;
            this.Message = message;
        }
    }
}
