# ClosedXML.Report.XLCustom

ClosedXML.Report.XLCustom extends [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) with enhanced expression handling and template processing capabilities while maintaining full compatibility with the original library.

## Overview

ClosedXML.Report.XLCustom builds upon ClosedXML.Report, preserving its familiar syntax while adding powerful dynamic expression evaluation and formatting capabilities.

## Dependencies

- [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) - The base reporting library
- [ClosedXML](https://github.com/ClosedXML/ClosedXML) - Excel spreadsheet manipulation library
- .NET 8.0 or later

## Key Features

- **Full ClosedXML.Report Compatibility**: Maintains all existing functionality and syntax:
  - Standard variables: `{{VariableName}}`
  - Property paths: `{{Object.Property.SubProperty}}`
  - Collection processing: `<<Range(Collection)>>` with `{{item.Property}}`
  - All template tags: `<<Range>>`, `<<Row>>`, `<<Group>>`, etc.

- **Enhanced Expression Syntax**:
  - Format expressions: `{{Value:format}}` for values with standard or custom formatting
  - Function expressions: `{{Value|function(params)}}` for direct cell manipulation

- **Extensibility**: Register custom formats and functions programmatically

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
var template = new XLCustomTemplate("template.xlsx");

// Register built-in formatters and functions
template.RegisterBuiltIns();

// Add variables as in standard ClosedXML.Report
template.AddVariable("Title", "Sales Report");
template.AddVariable("Products", productList);

// Register custom formats and functions
template.RegisterFormat("upper", (value, parameters) => value?.ToString()?.ToUpper());
template.RegisterFunction("style", (cell, value, parameters) => {
    // Manipulate cell directly
    cell.Style.Font.Bold = true;
    cell.Value = value;
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

Format expressions use the colon (`:`) syntax:

```
{{Value:F2}}           // Numeric format with 2 decimal places
{{Date:yyyy-MM-dd}}    // Date format
{{Price:C}}            // Currency format
{{Text:upper}}         // Custom "upper" formatter (converts to uppercase)
```

Register custom formatters:

```csharp
// Register a custom formatter
template.RegisterFormat("upper", (value, parameters) => 
    value?.ToString()?.ToUpper());

// In your template: {{Text:upper}}
```

### 3. Function Expressions

Function expressions use the pipe (`|`) syntax:

```
{{Value|style(bold, color=red)}}   // Apply bold red styling with parameters
{{ImageUrl|image(width=150)}}      // Display image with width parameter
{{LinkUrl|link}}                   // Link function with no parameters
```

Note: When a function has no parameters, the parentheses can be omitted (e.g., `{{Value|bold}}` instead of `{{Value|bold()}}`).


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

## Collection Processing

ClosedXML.Report.XLCustom fully supports ClosedXML.Report's collection processing:

```csharp
// Add collection data
template.AddVariable("Products", productList);
```

In Excel template:
```
<<Range(Products)>>
  A{{Row}}: {{item.Name}}
  B{{Row}}: {{item.Price}}
  C{{Row}}: {{item.Price:C}}           // Format expression
  D{{Row}}: {{item.Name|style(bold)}}  // Function expression
<<EndRange>>
```

## Built-in Formatters and Functions

ClosedXML.Report.XLCustom comes with several built-in formatters and functions that you can register:

```csharp
// Register all built-in formatters and functions
template.RegisterBuiltIns();

// Or register them separately
template.RegisterBuiltInFormatters();
template.RegisterBuiltInFunctions();
```

### Built-in Formatters

| Name | Description | Example |
|------|-------------|---------|
| `upper` | Converts text to uppercase | `{{Text:upper}}` |
| `lower` | Converts text to lowercase | `{{Text:lower}}` |
| `titlecase` | Converts text to title case | `{{Text:titlecase}}` |
| `mask` | Applies a mask to a value | `{{Phone:mask(###-###-####)}}` |
| `truncate` | Truncates text to specified length | `{{Description:truncate(50,...)}}` |

### Built-in Functions

| Name | Description | Example |
|------|-------------|---------|
| `bold` | Makes text bold | `{{Text|bold}}` |
| `italic` | Makes text italic | `{{Text|italic}}` |
| `color` | Sets text color | `{{Text|color(Red)}}` |
| `link` | Creates a hyperlink | `{{Url|link(Click here)}}` |
| `image` | Displays an image | `{{ImagePath|image(width=100,height=100)}}` |

## Relationship with ClosedXML.Report

This library extends the excellent [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report) project. Key points:
- XLCustom wraps and extends the XLTemplate class from ClosedXML.Report
- All existing ClosedXML.Report templates are fully compatible
- ClosedXML.Report tags (like `<<Group>>`, `<<Sort>>`, etc.) continue to work as expected
- XLCustom adds new expression processing capabilities that enhance the original functionality

## Advanced Examples

### Custom Formatter Example

```csharp
// Create a currency formatter with custom symbol
template.RegisterFormat("currency", (value, parameters) => 
{
    if (value == null) return null;
    
    if (decimal.TryParse(value.ToString(), out decimal amount))
    {
        string currencySymbol = parameters.Length > 0 ? parameters[0] : "$";
        return $"{currencySymbol}{amount:N2}";
    }
    
    return value;
});

// In template: {{Price:currency(¢æ)}}
```

### Custom Function Example

```csharp
// Create a conditional highlighting function
template.RegisterFunction("highlight", (cell, value, parameters) => 
{
    if (value == null) return;
    
    cell.Value = value;
    
    // Check if value meets the condition
    bool highlight = false;
    
    if (parameters.Length >= 2)
    {
        string condition = parameters[0];
        string threshold = parameters[1];
        
        if (decimal.TryParse(value.ToString(), out decimal numValue) && 
            decimal.TryParse(threshold, out decimal numThreshold))
        {
            switch (condition.ToLower())
            {
                case "gt":
                case ">":
                    highlight = numValue > numThreshold;
                    break;
                case "lt":
                case "<":
                    highlight = numValue < numThreshold;
                    break;
                case "eq":
                case "=":
                case "==":
                    highlight = numValue == numThreshold;
                    break;
            }
        }
    }
    
    if (highlight)
    {
        string color = parameters.Length >= 3 ? parameters[2] : "Yellow";
        cell.Style.Fill.BackgroundColor = XLColor.FromName(color);
        cell.Style.Font.Bold = true;
    }
});

// In template: {{Amount|highlight(>,1000,LightGreen)}}
```
