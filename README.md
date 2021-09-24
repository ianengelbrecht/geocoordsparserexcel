## Geographic coordinates parser for MS Excel

This is based on the [original coordinates parser development for javascript](https://www.npmjs.com/package/geo-coordinates-parser). Reads a variety of coordinate formats and converts them to decimal degrees. Provides three functions:

TODECIMAL(coordsString) returns a decimal coordinates string, e.g. '-23.45464, 28.343545'
TODECIMALLAT(coordsString) only returns the latitude
TODECIMALLAT(coordsString) only returns the longitude

##### NB In all cases coordsString must include latitude **AND** longitude. The converter checks that the format is consistent between them, and throws an error otherwise.  
###
To get the .xll add-in files for Excel, please visit [http://bit.ly/convertcoords](http://bit.ly/convertcoords). Note that you must install the bit version of the file that matches the bitness of your Excel version (not your computer). The simplest way to install the add in is to download it to a external flash disk, then double click the file on the flash disk. Excel then ensures the file is installed at the appropriate location. The download folder also includes a demo file with examples of the formats you can convert. 

This add-in makes use of a the coordinates converter, available as a dll, at [https://github.com/ianengelbrecht/GeoCoordinatesParserDotNet](https://github.com/ianengelbrecht/GeoCoor.dinatesParserDotNet)