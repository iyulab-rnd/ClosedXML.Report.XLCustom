using ClosedXML.Excel.Drawings;

namespace ClosedXML.Report.XLCustom.Functions;

/// <summary>
/// Provides built-in functions that can be registered with XLCustomTemplate
/// </summary>
public static class BuiltInFunctions
{
    /// <summary>
    /// Makes text bold
    /// </summary>
    public static readonly IXLCustomFunction Bold = new DelegateFunction(
        (cell, value, parameters) =>
        {
            if (cell == null) return;

            Debug.WriteLine($"Bold function called for cell {cell.Address}");

            // 스타일 적용
            cell.Style.Font.Bold = true;

            // 값 설정
            cell.SetValue(value);

            Debug.WriteLine($"Cell value after Bold: {cell.Value}, Bold applied: {cell.Style.Font.Bold}");
        }
    );

    /// <summary>
    /// Applies italic formatting to a cell
    /// </summary>
    public static readonly IXLCustomFunction Italic = new DelegateFunction(
        (cell, value, parameters) =>
        {
            cell.Style.Font.Italic = true;
            cell.SetValue(value);
        }
    );

    /// <summary>
    /// Sets the font color of a cell
    /// </summary>
    public static readonly IXLCustomFunction Color = new DelegateFunction(
        (cell, value, parameters) =>
        {
            if (parameters.Length > 0)
            {
                // Try to parse as named color
                try
                {
                    var color = XLColor.FromName(parameters[0]);
                    cell.Style.Font.FontColor = color;
                }
                catch
                {
                    // Try to parse as hex color
                    try
                    {
                        var hexColor = parameters[0].StartsWith("#") ? parameters[0] : "#" + parameters[0];
                        var color = XLColor.FromHtml(hexColor);
                        cell.Style.Font.FontColor = color;
                    }
                    catch
                    {
                        // Fallback to default
                        cell.Style.Font.FontColor = XLColor.Black;
                    }
                }
            }

            cell.SetValue(value);
        }
    );

    /// <summary>
    /// Adds a hyperlink to a cell
    /// </summary>
    public static readonly IXLCustomFunction Link = new DelegateFunction(
        (cell, value, parameters) =>
        {
            if (value != null)
            {
                string url = value.ToString();
                string text = parameters.Length > 0 ? parameters[0] : url;

                try
                {
                    cell.Value = text;

                    // Check if it's an internal or external link
                    if (url.StartsWith("#") || url.Contains("!"))
                    {
                        // For internal links
                        var hyperlink = cell.CreateHyperlink();
                        hyperlink.InternalAddress = url;
                        // Apply hyperlink style
                        cell.Style.Font.Underline = XLFontUnderlineValues.Single;
                        cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                    }
                    else
                    {
                        // For external links
                        if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                            url = "http://" + url;

                        var hyperlink = cell.CreateHyperlink();
                        hyperlink.ExternalAddress = new Uri(url);

                        // Apply hyperlink style
                        cell.Style.Font.Underline = XLFontUnderlineValues.Single;
                        cell.Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Hyperlink);
                    }
                }
                catch (Exception)
                {
                    // If URL is invalid, just display the text
                    cell.Value = text;
                }
            }
        }
    );

    /// <summary>
    /// Adds an image to a cell
    /// </summary>
    public static readonly IXLCustomFunction Image = new DelegateFunction(
        (cell, value, parameters) =>
        {
            if (value == null) return;

            try
            {
                IXLPicture picture;

                if (value is byte[] bytes)
                {
                    using (var ms = new MemoryStream(bytes))
                    {
                        picture = cell.Worksheet.AddPicture(ms);
                    }
                }
                else if (value is Stream stream)
                {
                    picture = cell.Worksheet.AddPicture(stream);
                }
                else
                {
                    string path = value.ToString();
                    if (File.Exists(path))
                    {
                        picture = cell.Worksheet.AddPicture(path);
                    }
                    else if (Uri.TryCreate(path, UriKind.Absolute, out var uri))
                    {
                        // This is a placeholder - you would need to add code to download the image
                        // using HttpClient or similar
                        throw new NotSupportedException("URL images are not supported directly. " +
                                                      "Download the image first and then use the file path.");
                    }
                    else
                    {
                        cell.Value = "Invalid image path";
                        return;
                    }
                }

                // Configure the picture
                picture.MoveTo(cell);

                // Process parameters (width, height)
                foreach (var param in parameters)
                {
                    string[] parts = param.Split('=');
                    if (parts.Length != 2) continue;

                    string name = parts[0].Trim().ToLower();
                    string value_str = parts[1].Trim();

                    if (int.TryParse(value_str, out int size))
                    {
                        switch (name)
                        {
                            case "width":
                                picture.Width = size;
                                break;
                            case "height":
                                picture.Height = size;
                                break;
                            case "scale":
                                picture.Scale(size / 100.0);
                                break;
                        }
                    }
                }

                // Clear the cell value so it doesn't overlap with the image
                cell.Value = string.Empty;
            }
            catch (Exception ex)
            {
                cell.Value = $"Error: {ex.Message}";
            }
        }
    );

    /// <summary>
    /// A function implementation that uses a delegate
    /// </summary>
    private class DelegateFunction : IXLCustomFunction
    {
        private readonly Action<IXLCell, object, string[]> _processFunction;

        public DelegateFunction(Action<IXLCell, object, string[]> processFunction)
        {
            _processFunction = processFunction ?? throw new ArgumentNullException(nameof(processFunction));
        }

        public void Process(IXLCell cell, object value, string[] parameters)
        {
            _processFunction(cell, value, parameters);
        }
    }

    // 디버깅용 헬퍼 메서드
    public static void DebugFunctions()
    {
        Debug.WriteLine("Built-in functions info:");
        Debug.WriteLine($"Bold function exists: {Bold != null}");
    }   
}