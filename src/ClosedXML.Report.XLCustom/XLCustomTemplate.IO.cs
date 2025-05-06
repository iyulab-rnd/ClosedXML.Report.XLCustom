namespace ClosedXML.Report.XLCustom;

public partial class XLCustomTemplate
{
    /// <summary>
    /// Saves the workbook to a file
    /// </summary>
    public void SaveAs(string file)
    {
        CheckIsDisposed();
        Workbook.SaveAs(file);
    }

    /// <summary>
    /// Saves the workbook to a file with options
    /// </summary>
    public void SaveAs(string file, SaveOptions options)
    {
        CheckIsDisposed();
        Workbook.SaveAs(file, options);
    }

    /// <summary>
    /// Saves the workbook to a file with validation
    /// </summary>
    public void SaveAs(string file, bool validate, bool evaluateFormulae = false)
    {
        CheckIsDisposed();
        Workbook.SaveAs(file, validate, evaluateFormulae);
    }

    /// <summary>
    /// Saves the workbook to a stream
    /// </summary>
    public void SaveAs(Stream stream)
    {
        CheckIsDisposed();
        Workbook.SaveAs(stream);
    }

    /// <summary>
    /// Saves the workbook to a stream with options
    /// </summary>
    public void SaveAs(Stream stream, SaveOptions options)
    {
        CheckIsDisposed();
        Workbook.SaveAs(stream, options);
    }

    /// <summary>
    /// Saves the workbook to a stream with validation
    /// </summary>
    public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false)
    {
        CheckIsDisposed();
        Workbook.SaveAs(stream, validate, evaluateFormulae);
    }
}