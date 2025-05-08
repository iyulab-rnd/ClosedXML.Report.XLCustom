# ClosedXML.Report.XLCustom

ClosedXML.Report.XLCustom extends [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) with enhanced expression handling capabilities while maintaining full compatibility with the original library.

## Overview

ClosedXML.Report.XLCustom builds upon ClosedXML.Report, preserving its familiar syntax while adding powerful dynamic expression evaluation and cell manipulation features through a simple yet expressive syntax.

## Dependencies

- [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) - The base reporting library
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) - Excel spreadsheet manipulation library

## Key Features

- **Full ClosedXML.Report Compatibility**: Maintains all existing functionality and syntax:
  - Standard variables: `{{VariableName}}`
  - Property paths: `{{Object.Property.SubProperty}}`
  - All template tags from the original library

- **Enhanced Expression Syntax**:
  - Format expressions: `{{Value:format}}` for values with standard .NET formatting
  - Function expressions: `{{Value|function(param1,param2)}}` for direct cell manipulation

- **Extensibility**: Register custom functions programmatically
  - Global function registry for application-wide functions
  - Local registry option for template-specific functions

- **Error Handling**: Clear error indicators and messages in cells when expressions fail

## Installation

```
Install-Package ClosedXML.Report.XLCustom
```

Or via .NET CLI:

```
dotnet add package ClosedXML.Report.XLCustom
```

## Quick Start

```csharp
// Create a template from an existing Excel file
var options = new XLCustomTemplateOptions { 
    UseGlobalRegistry = true,
    RegisterBuiltInFunctions = true
};
var template = new XLCustomTemplate("template.xlsx", options);

// Register built-in functions
template.RegisterBuiltIns();

// Add variables as in standard ClosedXML.Report
template.AddVariable("Title", "Sales Report");
template.AddVariable("Products", productList);

// Register custom functions
template.RegisterFunction("highlight", (cell, value, parameters) => {
    // Manipulate cell directly
    cell.SetValue(value);
    cell.Style.Fill.BackgroundColor = XLColor.Yellow;
    cell.Style.Font.Bold = true;
});

// Generate the report
template.Generate();

// Save the result
template.SaveAs("result.xlsx");
```

## Expression Syntax

### 1. Standard Variables (ClosedXML.Report compatible)

```
{{VariableName}}
{{Object.Property}}
```

### 2. Format Expressions

Format expressions use the colon (`:`) syntax with standard .NET format strings:

```
{{Value:F2}}           // Numeric format with 2 decimal places
{{Date:yyyy-MM-dd}}    // Date format
{{Price:C}}            // Currency format
```

### 3. Function Expressions

Function expressions use the pipe (`|`) syntax with positional parameters:

```
{{Value|function}}                 // Function with no parameters
{{Value|style(bold,red)}}          // Apply styling with parameters
{{ImageUrl|image(150)}}            // Display image with width parameter
{{LinkUrl|link(Click here)}}       // Create link with display text
```

Note: When a function has no parameters, the parentheses can be omitted.

Register custom functions:

```csharp
// Register a cell processor
template.RegisterFunction("image", (cell, value, parameters) => {
    string imagePath = value?.ToString();
    if (!string.IsNullOrEmpty(imagePath))
    {
        var picture = cell.Worksheet.AddPicture(imagePath);
        
        // Apply optional size parameters
        if (parameters?.Length > 0 && int.TryParse(parameters[0], out int width))
            picture.Width = width;
            
        picture.MoveTo(cell);
    }
});

// In your template: {{LogoUrl|image(150)}}
```

## Built-in Functions

ClosedXML.Report.XLCustom comes with several built-in functions that you can register:

```csharp
// Register all built-in functions
template.RegisterBuiltIns();

// Or register them individually
template.RegisterBuiltInFunctions();
```

### Built-in Functions

| Name | Description | Example | Parameters |
|------|-------------|---------|------------|
| `bold` | Makes text bold | `{{Text|bold}}` | None |
| `italic` | Makes text italic | `{{Text|italic}}` | None |
| `color` | Sets text color | `{{Text|color(Red)}}` | 1: Color name |
| `link` | Creates a hyperlink | `{{Url|link(Click here)}}` | 1: Display text (optional) |
| `image` | Displays an image | `{{ImagePath|image(100,100)}}` | 1: Width (optional), 2: Height (optional) |

## Global vs Local Function Registry

ClosedXML.Report.XLCustom supports two registry modes:

```csharp
// Global registry (default) - functions available to all templates
var template1 = new XLCustomTemplate("template1.xlsx"); // useGlobalRegistry = true by default

// Local registry - functions available only to this template
var options = new XLCustomTemplateOptions { UseGlobalRegistry = false };
var template2 = new XLCustomTemplate("template2.xlsx", options);

// Reset the global registry if needed
XLCustomRegistry.ResetFunctionRegistry();
```

## Error Handling

When an expression fails to evaluate or a function encounters an error, the cell will show an error message in red text. This helps identify issues in your templates:

- Variable evaluation errors: "Error: [error message]"
- Unknown functions: "Unknown function: [name]" 
- Function execution errors: "Function error: [error message]"

## Implementation Details

XLCustomTemplate automatically performs a preprocessing step on your Excel template to convert the enhanced expression syntax (like `{{Value:format}}` and `{{Value|function}}`) into compatible ClosedXML.Report tags:

- Format expressions (`{{Value:format}}`) ¡æ `<<format name="Value" format="format">>`
- Function expressions (`{{Value|function(params)}}`) ¡æ `<<customfunction name="Value" function="function" parameters="params">>`

This preprocessing is done automatically when you call methods like `Generate()`, `AddVariable()`, or `SaveAs()`. You can also force preprocessing with the `Preprocess()` method.

## Debugging

For debugging purposes, you can test expression processing without applying it to a workbook:

```csharp
// Test expression processing
string result = template.DebugExpression("{{Value|bold}}");
Console.WriteLine(result); // Outputs: <<customfunction name="Value" function="bold">>
```

## Parameter Handling

When using custom functions with multiple parameters, ClosedXML.Report.XLCustom properly handles parameter escaping:

```csharp
// Parameter with commas
{{Text|function('parameter with, comma')}}

// Parameter with parentheses
{{Text|function('parameter with (parens)')}}
```

The library correctly parses these parameters and passes them to your custom functions.

## Global Variables

ClosedXML.Report.XLCustom supports global variables that can be used across all templates:

```csharp
// Register global variables
template.RegisterGlobalVariable("CompanyName", "ACME Inc.");
template.RegisterGlobalVariable("ReportDate", DateTime.Today);

// Register global variable with dynamic value
template.RegisterGlobalVariable("RandomNumber", () => new Random().Next(1, 100));
```

### Built-in Global Variables

The following built-in global variables are automatically available:

| Name | Description | Example |
|------|-------------|---------|
| `Today` | Current date | `{{Today:d}}` |
| `Now` | Current date and time | `{{Now:f}}` |
| `Year` | Current year | `{{Year}}` |
| `Month` | Current month | `{{Month}}` |
| `Day` | Current day | `{{Day}}` |
| `MachineName` | Computer name | `{{MachineName}}` |
| `UserName` | User name | `{{UserName}}` |