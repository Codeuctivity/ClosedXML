using SkiaSharp;
using System;

namespace ClosedXML.Utils
{
    internal static class GraphicsUtils
    {
        internal static SKRect MeasureString(string text, SKFont font)
        {
            using var paint = new SKPaint();
            paint.Typeface = font.Typeface;

            paint.TextSize = font.Size;

            var skBounds = SKRect.Empty;
            var textWidth = paint.MeasureText(text.AsSpan(), ref skBounds);
            return skBounds;
        }
    }
}