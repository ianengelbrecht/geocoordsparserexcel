using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using GeoCoordinateParser;

namespace GeoCoordinatesParserExcel
{
    public static class Functions
    {
        [ExcelFunction(Description = "Convert verbatim coordinates to decimal coordinates")]
        public static object TODECIMAL(string coordsString)
        {
            try
            {
                Coordinates converted = CoordinatesConverter.convert(coordsString);
                return converted.decimalCoordinates;
            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }

        }

        [ExcelFunction(Description = "Get decimal latitude from verbatim coordinates")]
        public static object TODECIMALLAT(string coordsString)
        {
            try
            {
                Coordinates converted = CoordinatesConverter.convert(coordsString);
                return converted.decimalLatitude;
            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }
        }

        [ExcelFunction(Description = "Get decimal longitude from verbatim coordinates")]
        public static object TODECIMALLNG(string coordsString)
        {
            try
            {
                Coordinates converted = CoordinatesConverter.convert(coordsString);
                return converted.decimalLongitude;
            }
            catch
            {
                return ExcelError.ExcelErrorValue;
            }
        }
    }
}
