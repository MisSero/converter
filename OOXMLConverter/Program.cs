namespace OOXMLConverter;

public class Program
{
	static void Main(string[] args)
	{
		Console.WriteLine("Введите путь к HTML файлу для создания word документа");
		string? htmlFilePath = Console.ReadLine();

		if (!string.IsNullOrEmpty(htmlFilePath) && File.Exists(htmlFilePath))
		{
			string wordFilePath = Path.Combine(Path.GetDirectoryName(htmlFilePath),
				Path.GetFileNameWithoutExtension(htmlFilePath)) + ".docx";

			Converter converter = new Converter();
			converter.CreateWordFile(htmlFilePath, wordFilePath);
		}
	}
}
