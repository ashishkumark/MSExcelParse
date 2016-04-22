# MSExcelParse
Sample file using Apache Tika to parse MS Excel file & OpenOffice Spreadsheet

-------------------
Create new test1.xls file to avoid below error during execution:
Exception in thread "main" org.apache.tika.exception.TikaException: Error creating OOXML extractor
	at org.apache.tika.parser.microsoft.ooxml.OOXMLExtractorFactory.parse(OOXMLExtractorFactory.java:123)
	at org.apache.tika.parser.microsoft.ooxml.OOXMLParser.parse(OOXMLParser.java:87)
	at com.ashish.parse.excel.ExcelParse.main(ExcelParse.java:30)
Caused by: org.apache.poi.openxml4j.exceptions.InvalidFormatException: Package should contain a content type part [M1.13]
	at org.apache.poi.openxml4j.opc.ZipPackage.getPartsImpl(ZipPackage.java:201)
	at org.apache.poi.openxml4j.opc.OPCPackage.getParts(OPCPackage.java:684)
	at org.apache.poi.openxml4j.opc.OPCPackage.open(OPCPackage.java:275)
	at org.apache.tika.parser.microsoft.ooxml.OOXMLExtractorFactory.parse(OOXMLExtractorFactory.java:73)
	... 2 more
--------------------------
