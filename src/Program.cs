using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JsonToXlsx;

internal class Program
{
    static void Main(string[] args)
    {
        var filePaths = args.ToList();

        if (filePaths.All(f => Path.GetExtension(f).ToLower() == ".json"))
        {
            // Used to store each file's key-value pairs
            var data = new Dictionary<string, List<object>>();
            var fileNames = new List<string>(); // Save the column names

            // Read the content of each JSON file
            for (int i = 0; i < filePaths.Count; i++)
            {
                string filePath = filePaths[i];
                string jsonContent = File.ReadAllText(filePath);

                Console.WriteLine(filePath);

                // Parse JSON into Dictionary<string, object>
                var record = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonContent);

                // Get the short file name (without path and extension)
                string shortFileName = Path.GetFileNameWithoutExtension(filePath);
                fileNames.Add(shortFileName);

                // Add JSON data to 'data', ensuring keys align across multiple files
                foreach (var kvp in record!)
                {
                    if (!data.ContainsKey(kvp.Key))
                    {
                        // Initialize a new row for the key
                        data[kvp.Key] = [];
                    }

                    // Add value and ensure alignment with other files
                    data[kvp.Key].Add(kvp.Value);
                }
            }

            // Fill missing keys with null if a file doesn't contain the key
            foreach (var key in data.Keys)
            {
                while (data[key].Count < filePaths.Count)
                {
                    data[key].Add(null!); // Use null to fill missing values
                }
            }

            // Convert data into the format MiniExcel expects (List<Dictionary<string, object>>)
            var rows = new List<Dictionary<string, object>>();

            // Create the first row (header), with 'key' as the column name
            var headerRow = new Dictionary<string, object> { { "key", "key" } };
            for (int i = 0; i < fileNames.Count; i++)
            {
                headerRow[fileNames[i]] = fileNames[i]; // Use short file names as header columns
            }

            // Create rows for each key and corresponding values
            foreach (var key in data.Keys)
            {
                var row = new Dictionary<string, object>
            {
                { "key", key }
            };

                for (int i = 0; i < data[key].Count; i++)
                {
                    row[fileNames[i]] = data[key][i];
                }

                rows.Add(row);
            }

            // Export to Excel file
            string excelFilePath = "output.xlsx";
            MiniExcel.SaveAs(excelFilePath, rows, printHeader: true, configuration: new OpenXmlConfiguration() { AutoFilter = false }, overwriteFile: true);

            Console.WriteLine("Excel file generated successfully!");
        }
        else if (filePaths.All(f => Path.GetExtension(f).ToLower() == ".xlsx"))
        {
            Console.WriteLine(filePaths[0]);
            FileInfo fileInfo = new(filePaths[0]);
            IEnumerable<dynamic> xlsx = MiniExcel.Query(filePaths[0], sheetName: "Sheet1");
            List<List<string>> lines = [];

            // Iterate over each row (line) in the Excel file
            foreach (dynamic lineItem in xlsx)
            {
                List<string> line = [];  // Create a new list to store values of the current row

                // Loop through each key-value pair (cell) in the current row
                foreach (KeyValuePair<string, object> item in lineItem)
                {
                    // Convert each cell value to a string, if null, replace it with an empty string
                    line.Add(item.Value?.ToString() ?? string.Empty);
                }

                // Add the processed row (list of strings) to the list of lines
                lines.Add(line);
            }

            // Create a list to store file paths and associated JSON data
            var columns = new List<(string, Dictionary<string, object>)>();

            // Process each line (row) in the list
            for (int i = 0; i < lines.Count; i++)
            {
                var line = lines[i];

                // For the first row (header row), we treat it as the file name identifiers
                if (i == 0)
                {
                    // Start from the second column (skip the 'key' column)
                    for (int j = 1; j < line.Count(); j++)
                    {
                        // Combine the directory name with the header value to create the JSON file path
                        columns.Add((Path.Combine(fileInfo.DirectoryName, $"{line[j]}.json"), []));
                    }

                    // Skip further processing for the header row
                    continue;
                }

                string key = null!;  // Initialize a variable to store the key for each row

                // Iterate over each cell in the current row
                for (int j = 0; j < line.Count(); j++)
                {
                    // The first column of each row contains the key
                    if (j == 0)
                    {
                        // Set the key to the value in the first column
                        key = line[j];
                        continue;
                    }

                    // If the key is null, skip the processing
                    if (key == null)
                    {
                        continue;
                    }

                    var value = line[j];
                    columns[j - 1].Item2.Add(key, value);
                }
            }

            // Generate JSON files from the collected data
            foreach (var column in columns)
            {
                var fileName = column.Item1;
                var json = JsonConvert.SerializeObject(column.Item2);
                File.WriteAllText(fileName, JsonBeautifier.Beautify(json));
                Console.WriteLine(fileName);
            }

            Console.WriteLine("Json files generated successfully!");
        }
    }
}

file static class JsonBeautifier
{
    public static string Beautify(string input)
    {
        object? obj = JsonConvert.DeserializeObject(SortJson(input));

        if (obj != null)
        {
            using StringWriter textWriter = new();
            using JsonTextWriter jsonWriter = new(textWriter)
            {
                Formatting = Formatting.Indented,
                Indentation = 4,
                IndentChar = ' ',
            };
            new JsonSerializer().Serialize(jsonWriter, obj);
            return textWriter.ToString();
        }
        else
        {
            return input;
        }
    }

    public static string SortJson(string json)
    {
        SortedDictionary<string, object>? dic = JsonConvert.DeserializeObject<SortedDictionary<string, object>>(json);
        SortedDictionary<string, object> keyValues = new(dic);
        keyValues.OrderBy(m => m.Key);

        SortedDictionary<string, object> tempKeyValues = new(keyValues);
        foreach (KeyValuePair<string, object> kv in tempKeyValues)
        {
            Type t0 = typeof(JObject);
            Type? t1 = kv.Value?.GetType();

            if (t0 == t1)
            {
                string jsonItem = JsonConvert.SerializeObject(kv.Value);
                jsonItem = SortJson(jsonItem);
                keyValues[kv.Key] = JsonConvert.DeserializeObject<JObject>(jsonItem)!;
            }
        }
        return JsonConvert.SerializeObject(keyValues);
    }
}
