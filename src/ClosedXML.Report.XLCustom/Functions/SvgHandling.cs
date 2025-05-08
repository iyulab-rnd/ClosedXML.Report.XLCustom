using SkiaSharp;
using Svg.Skia;
using System;
using System.IO;

namespace ClosedXML.Report.XLCustom.Functions
{
    /// <summary>
    /// Provides SVG image processing functionality
    /// </summary>
    public static class SvgHandling
    {
        /// <summary>
        /// Default width used when SVG dimensions cannot be determined
        /// </summary>
        private const int DEFAULT_SVG_WIDTH = 300;

        /// <summary>
        /// Default height used when SVG dimensions cannot be determined
        /// </summary>
        private const int DEFAULT_SVG_HEIGHT = 300;

        /// <summary>
        /// Converts SVG file to PNG format
        /// </summary>
        public static bool ConvertSvgToPng(string svgFilePath, string pngFilePath)
        {
            if (string.IsNullOrEmpty(svgFilePath) || string.IsNullOrEmpty(pngFilePath))
            {
                Log.Debug("Invalid file paths for SVG conversion");
                return false;
            }

            if (!File.Exists(svgFilePath))
            {
                Log.Debug($"SVG file not found: {svgFilePath}");
                return false;
            }

            try
            {
                using (var svg = new SKSvg())
                {
                    // Load SVG
                    svg.Load(svgFilePath);

                    // Validate SVG loaded successfully
                    if (svg.Picture == null)
                    {
                        Log.Debug("Failed to load SVG - Picture is null");
                        return false;
                    }

                    // Extract dimensions from SVG
                    var dimensions = GetSvgDimensions(svg);

                    // Create PNG
                    return CreatePngFromSvg(svg, pngFilePath, dimensions.width, dimensions.height);
                }
            }
            catch (Exception ex)
            {
                Log.Debug($"Error in SVG to PNG conversion: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets dimensions from SVG image with fallback to default values
        /// </summary>
        private static (int width, int height) GetSvgDimensions(SKSvg svg)
        {
            if (svg?.Picture == null)
            {
                return (DEFAULT_SVG_WIDTH, DEFAULT_SVG_HEIGHT);
            }

            SKRect bounds = svg.Picture.CullRect;
            int width = (int)bounds.Width;
            int height = (int)bounds.Height;

            // Use default dimensions if values are invalid
            if (width <= 0 || height <= 0)
            {
                width = DEFAULT_SVG_WIDTH;
                height = DEFAULT_SVG_HEIGHT;
            }

            return (width, height);
        }

        /// <summary>
        /// Creates PNG file from SVG image with specified dimensions
        /// </summary>
        private static bool CreatePngFromSvg(SKSvg svg, string pngFilePath, int width, int height)
        {
            if (svg?.Picture == null)
            {
                return false;
            }

            try
            {
                // Scale dimensions to ensure they are reasonable
                // Extremely large dimensions can cause memory issues
                if (width > 4000 || height > 4000)
                {
                    double scale = Math.Min(4000.0 / width, 4000.0 / height);
                    width = (int)(width * scale);
                    height = (int)(height * scale);
                    Log.Debug($"SVG dimensions scaled down to {width}x{height}");
                }

                // Ensure dimensions are positive
                width = Math.Max(1, width);
                height = Math.Max(1, height);

                using (var surface = SKSurface.Create(new SKImageInfo(width, height)))
                {
                    if (surface == null)
                    {
                        Log.Debug("Failed to create surface for SVG rendering");
                        return false;
                    }

                    var canvas = surface.Canvas;
                    canvas.Clear(SKColors.Transparent);

                    // Scale SVG to fit if needed
                    float scaleX = width / svg.Picture.CullRect.Width;
                    float scaleY = height / svg.Picture.CullRect.Height;

                    if (scaleX != 1.0f || scaleY != 1.0f)
                    {
                        canvas.Scale(scaleX, scaleY);
                    }

                    canvas.DrawPicture(svg.Picture);
                    canvas.Flush();

                    using (var image = surface.Snapshot())
                    {
                        // Verify image was created successfully
                        if (image == null)
                        {
                            Log.Debug("Failed to get snapshot from surface");
                            return false;
                        }

                        using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
                        {
                            if (data == null)
                            {
                                Log.Debug("Failed to encode image");
                                return false;
                            }

                            try
                            {
                                using (var stream = File.OpenWrite(pngFilePath))
                                {
                                    data.SaveTo(stream);
                                }
                            }
                            catch (Exception ex)
                            {
                                Log.Debug($"Failed to write PNG file: {ex.Message}");
                                return false;
                            }
                        }
                    }

                    // Verify file was created successfully
                    if (!File.Exists(pngFilePath))
                    {
                        Log.Debug($"PNG file was not created: {pngFilePath}");
                        return false;
                    }

                    var fileInfo = new FileInfo(pngFilePath);
                    if (fileInfo.Length == 0)
                    {
                        Log.Debug("Created PNG file is empty");
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                Log.Debug($"Error creating PNG from SVG: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Validates if file is a valid SVG
        /// </summary>
        public static bool IsValidSvg(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            try
            {
                string content = File.ReadAllText(filePath);
                return !string.IsNullOrEmpty(content) && (content.Contains("<svg") || content.Contains("<?xml"));
            }
            catch (Exception ex)
            {
                Log.Debug($"Error validating SVG file: {ex.Message}");
                return false;
            }
        }
    }
}