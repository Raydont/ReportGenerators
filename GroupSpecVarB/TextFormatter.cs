using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace GroupSpecVarB
{
    public class TextFormatter
    {
        Font _font;

        int _displayResolutionX;

        public TextFormatter(Font font)
        {
            _font = font;
            _displayResolutionX = GetDisplayResolutionX();
        }

        [DllImport("User32.dll")]
        static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        public enum DeviceCap
        {
            #region
            /// <summary>
            /// Device driver version
            /// </summary>
            DRIVERVERSION = 0,
            /// <summary>
            /// Device classification
            /// </summary>
            TECHNOLOGY = 2,
            /// <summary>
            /// Horizontal size in millimeters
            /// </summary>
            HORZSIZE = 4,
            /// <summary>
            /// Vertical size in millimeters
            /// </summary>
            VERTSIZE = 6,
            /// <summary>
            /// Horizontal width in pixels
            /// </summary>
            HORZRES = 8,
            /// <summary>
            /// Vertical height in pixels
            /// </summary>
            VERTRES = 10,
            /// <summary>
            /// Number of bits per pixel
            /// </summary>
            BITSPIXEL = 12,
            /// <summary>
            /// Number of planes
            /// </summary>
            PLANES = 14,
            /// <summary>
            /// Number of brushes the device has
            /// </summary>
            NUMBRUSHES = 16,
            /// <summary>
            /// Number of pens the device has
            /// </summary>
            NUMPENS = 18,
            /// <summary>
            /// Number of markers the device has
            /// </summary>
            NUMMARKERS = 20,
            /// <summary>
            /// Number of fonts the device has
            /// </summary>
            NUMFONTS = 22,
            /// <summary>
            /// Number of colors the device supports
            /// </summary>
            NUMCOLORS = 24,
            /// <summary>
            /// Size required for device descriptor
            /// </summary>
            PDEVICESIZE = 26,
            /// <summary>
            /// Curve capabilities
            /// </summary>
            CURVECAPS = 28,
            /// <summary>
            /// Line capabilities
            /// </summary>
            LINECAPS = 30,
            /// <summary>
            /// Polygonal capabilities
            /// </summary>
            POLYGONALCAPS = 32,
            /// <summary>
            /// Text capabilities
            /// </summary>
            TEXTCAPS = 34,
            /// <summary>
            /// Clipping capabilities
            /// </summary>
            CLIPCAPS = 36,
            /// <summary>
            /// Bitblt capabilities
            /// </summary>
            RASTERCAPS = 38,
            /// <summary>
            /// Length of the X leg
            /// </summary>
            ASPECTX = 40,
            /// <summary>
            /// Length of the Y leg
            /// </summary>
            ASPECTY = 42,
            /// <summary>
            /// Length of the hypotenuse
            /// </summary>
            ASPECTXY = 44,
            /// <summary>
            /// Shading and Blending caps
            /// </summary>
            SHADEBLENDCAPS = 45,

            /// <summary>
            /// Logical pixels inch in X
            /// </summary>
            LOGPIXELSX = 88,
            /// <summary>
            /// Logical pixels inch in Y
            /// </summary>
            LOGPIXELSY = 90,

            /// <summary>
            /// Number of entries in physical palette
            /// </summary>
            SIZEPALETTE = 104,
            /// <summary>
            /// Number of reserved entries in palette
            /// </summary>
            NUMRESERVED = 106,
            /// <summary>
            /// Actual color resolution
            /// </summary>
            COLORRES = 108,

            // Printing related DeviceCaps. These replace the appropriate Escapes
            /// <summary>
            /// Physical Width in device units
            /// </summary>
            PHYSICALWIDTH = 110,
            /// <summary>
            /// Physical Height in device units
            /// </summary>
            PHYSICALHEIGHT = 111,
            /// <summary>
            /// Physical Printable Area x margin
            /// </summary>
            PHYSICALOFFSETX = 112,
            /// <summary>
            /// Physical Printable Area y margin
            /// </summary>
            PHYSICALOFFSETY = 113,
            /// <summary>
            /// Scaling factor x
            /// </summary>
            SCALINGFACTORX = 114,
            /// <summary>
            /// Scaling factor y
            /// </summary>
            SCALINGFACTORY = 115,

            /// <summary>
            /// Current vertical refresh rate of the display device (for displays only) in Hz
            /// </summary>
            VREFRESH = 116,
            /// <summary>
            /// Horizontal width of entire desktop in pixels
            /// </summary>
            DESKTOPVERTRES = 117,
            /// <summary>
            /// Vertical height of entire desktop in pixels
            /// </summary>
            DESKTOPHORZRES = 118,
            /// <summary>
            /// Preferred blt alignment
            /// </summary>
            BLTALIGNMENT = 119
            #endregion
        }

        int GetDisplayResolutionX()
        {
            IntPtr p = GetDC(new IntPtr(0));
            int res = GetDeviceCaps(p, (int)DeviceCap.LOGPIXELSX);
            return res;
        }

        /// Коэффициент, учитывающий различие между шириной строки,
        /// определяемой в этой программе и шириной строки, получаемой в CAD
        static readonly float widthFactor = 1.5f;

        public double GetTextWidth(string text, Font font)
        {
            Size size = TextRenderer.MeasureText(text, font);
            double width = 25.4f * widthFactor * (double)size.Width / (double)_displayResolutionX;
            return width;
        }

        public string[] Wrap(string text, double maxWidth)
        {
            return Wrap(text, _font, maxWidth);
        }

        public string[] WrapName(string text, double maxWidth)
        {
            string patternWrapWord = "(База данных|Спецификация|Описание применения|Текст программы|Лист утверждения|Программа|Библиотека|Операционная система|Модуль)";
            Regex regex = new Regex(patternWrapWord);
            MatchCollection match;
            match = regex.Matches(text);

            int indx;
            List<string> rowList = new List<string>();
            List<string> wrappedLines = new List<string>();
            string beginPart;


            for (int i = 0; i < match.Count; i++)
            {
                if (match[i].Value.Trim() != string.Empty)
                    rowList.Add(match[i].Value);
            }

            if (rowList.Count > 0)
            {
                foreach (string row in rowList)
                {
                    indx = text.IndexOf(row);
                    beginPart = text.Remove(indx);
                    if (beginPart.Trim() != string.Empty)
                    {
                        text = text.Remove(0, beginPart.Length);
                        text = text.Trim();
                        wrappedLines.AddRange(Wrap(beginPart, maxWidth));
                    }
                    wrappedLines.Add(row);
                    text = text.Remove(0, row.Length);
                }
                return wrappedLines.ToArray();
            }
            else return Wrap(text, _font, maxWidth);

        }

        /// Разбивка на строки по символам перевода строки \n
        string[] Wrap(string text, Font font, double maxWidth)
        {
            string[] lines = text.Split(new string[] { @"\n" }, StringSplitOptions.None);

            List<string> wrappedLines = new List<string>();
            foreach (string line in lines)
            {
                string[] wrappedBySpaces = WrapBySyllables(line, font, maxWidth);
                wrappedLines.AddRange(wrappedBySpaces);
            }
            return wrappedLines.ToArray();
        }

        /// Выделение слога - последовательности символов с завершающим пробелом
        /// или символом мягкого переноса. Символ мягкого переноса (если присутствует)
        /// выносится в отдельную группу
        static readonly Regex _syllableRegex = new Regex(
            @"(?<syllable>[^\x20\u00AD]+)(?<softHyphen>\u00AD+)?",
            RegexOptions.Compiled);

        /// Разбивка на строки по слогам (разделители - пробелы и знаки мягкого переноса)
        string[] WrapBySyllables(string text, Font font, double maxWidth)
        {
            if (text == "")
                return new string[] { "" };

            List<string> lines = new List<string>();
            MatchCollection mc = _syllableRegex.Matches(text);
            string currentLine = "";
            bool currentLineEndsWithSoftHyphen = false;
            foreach (Match match in mc)
            {
                string syllable = match.Groups["syllable"].Value;
                bool endsBySoftHyphen = match.Groups["softHyphen"].Success;

                string candidateLine;
                if (currentLine.Length > 0)
                {
                    if (currentLineEndsWithSoftHyphen)
                        candidateLine = currentLine + syllable;
                    else
                        candidateLine = currentLine + " " + syllable;
                }
                else
                    candidateLine = syllable;
                double candidateLineWidth = GetTextWidth(candidateLine, font);

                if (candidateLineWidth > maxWidth)
                {
                    // Перенос очередной последовательности символов на новую строку

                    if (currentLine.Length > 0)
                    {
                        if (currentLineEndsWithSoftHyphen)
                            lines.Add(currentLine + "-");
                        else
                            lines.Add(currentLine);
                    }
                    currentLine = "";
                    currentLineEndsWithSoftHyphen = false;

                    candidateLine = syllable;
                    candidateLineWidth = GetTextWidth(candidateLine, font);
                }

                if (candidateLineWidth > maxWidth)
                {
                    // Разбивка строки-кандидата по разделителям-не буквам

                    if (currentLine.Length > 0)
                    {
                        if (currentLineEndsWithSoftHyphen)
                            lines.Add(currentLine + "-");
                        else
                            lines.Add(currentLine);
                    }
                    currentLine = "";
                    currentLineEndsWithSoftHyphen = false;

                    string[] candidateLines = WrapByAlphaAndDelimiter(candidateLine, font, maxWidth);
                    foreach (string s in candidateLines)
                        lines.Add(s);
                }
                else
                {
                    // строка-кандидат не превышает максимально допустимую ширину
                    currentLine = candidateLine;
                    currentLineEndsWithSoftHyphen = endsBySoftHyphen;
                }
            }
            if (currentLine.Length > 0)
                lines.Add(currentLine);
            return lines.ToArray();
        }

        static readonly Regex _alphaAndDelimiterRegex = new Regex(
            @"[\p{Ll}\p{Lu}]+[^\p{Ll}\p{Lu}]*", RegexOptions.Compiled);

        /// Разбивка на строки по разделителям-не буквам
        string[] WrapByAlphaAndDelimiter(string text, Font font, double maxWidth)
        {
            if (text == "")
                return new string[] { "" };

            List<string> lines = new List<string>();
            MatchCollection mc = _alphaAndDelimiterRegex.Matches(text);
            string currentLine = "";
            foreach (Match match in mc)
            {
                string candidateLine = currentLine + match.Value;
                double candidateLineWidth = GetTextWidth(candidateLine, font);

                if (candidateLineWidth > maxWidth)
                {
                    // Перенос очередной последовательности символов на новую строку

                    if (currentLine.Length > 0)
                        lines.Add(currentLine);
                    currentLine = "";

                    candidateLine = match.Value;
                    candidateLineWidth = GetTextWidth(candidateLine, font);
                }

                if (candidateLineWidth > maxWidth)
                {
                    // Жесткая разбивка новой последовательности символов

                    if (currentLine.Length > 0)
                        lines.Add(currentLine);
                    currentLine = "";

                    string[] candidateLines = WrapByCharacters(candidateLine, font, maxWidth);
                    foreach (string s in candidateLines)
                        lines.Add(s);
                }
                else
                    // строка-кандидат не превышает максимально допустимую ширину
                    currentLine = candidateLine;
            }
            if (currentLine.Length > 0)
                lines.Add(currentLine);
            return lines.ToArray();
        }

        /// Разбивка на строки по символам
        string[] WrapByCharacters(string text, Font font, double maxWidth)
        {
            if (text == "")
                return new string[] { "" };

            List<string> lines = new List<string>();
            string currentLine = "";
            for (int i = 0; i < text.Length; ++i)
            {
                string candidateLine = currentLine + text.Substring(i, 1);
                double candidateLineWidth = GetTextWidth(candidateLine, font);
                if (candidateLineWidth > maxWidth)
                {
                    if (currentLine.Length > 0)
                        lines.Add(currentLine);
                    currentLine = text.Substring(i, 1);
                }
                else
                    currentLine = candidateLine;
            }
            if (currentLine.Length > 0)
                lines.Add(currentLine);
            return lines.ToArray();
        }
    }
}
