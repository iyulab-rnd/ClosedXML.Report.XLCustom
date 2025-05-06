namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Utility class for converting objects to XLCellValue
/// </summary>
public static class XLCellValueConverter
{
    /// <summary>
    /// Converts an object to XLCellValue
    /// </summary>
    public static XLCellValue FromObject(object obj, IFormatProvider provider = null)
    {
        if (obj == null)
            return Blank.Value;

        // If already an XLCellValue, return directly
        if (obj is XLCellValue cellValue)
            return cellValue;

        // Handle special types
        return obj switch
        {
            null => Blank.Value,
            Blank blank => blank,
            bool logical => logical,
            string text => text,
            XLError error => error,
            DateTime dateTime => dateTime,
            TimeSpan timeSpan => timeSpan,
            sbyte number => number,
            byte number => number,
            short number => number,
            ushort number => number,
            int number => number,
            uint number => number,
            long number => number,
            ulong number => number,
            float number => number,
            double number => number,
            decimal number => number,
            // For any other type, safely convert to string
            _ => SafeToString(obj, provider)
        };
    }

    /// <summary>
    /// Safely converts an object to string, handling exceptions
    /// </summary>
    private static string SafeToString(object obj, IFormatProvider provider)
    {
        try
        {
            return Convert.ToString(obj, provider);
        }
        catch (Exception ex)
        {
            // Log the error and return a safe representation
            Debug.WriteLine($"Error converting {obj?.GetType().Name ?? "null"} to string: {ex.Message}");
            return $"{obj?.GetType().Name ?? "null"}";
        }
    }
}