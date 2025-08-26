namespace ReverseGeoCoding
{
    internal class FromBackend
    {
        private string inputFilepath;
        private string outputFilepath;

        public FromBackend(string inputFilepath, string outputFilepath)
        {
            this.inputFilepath = inputFilepath;
            this.outputFilepath = outputFilepath;
        }
    }
}