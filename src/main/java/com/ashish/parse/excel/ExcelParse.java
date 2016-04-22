package com.ashish.parse.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.microsoft.ooxml.OOXMLParser;
import org.apache.tika.parser.odf.OpenDocumentParser;
import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.SAXException;

public class ExcelParse {

	public static void main(String[] args) throws IOException, SAXException, TikaException {
	      //detecting the file type
	      BodyContentHandler handler = new BodyContentHandler();
	      Metadata metadata = new Metadata();
	      ParseContext pcontext = new ParseContext();
	      
//	      System.out.println(System.getProperty("user.dir"));
//	      System.out.println((new File(".")).getAbsoluteFile());
//
	      FileInputStream MSEinputstream = new FileInputStream(new File(".\\test1.xls"));
	      
	      //OOXml parser - MS Office Parser
	      OOXMLParser  msofficeparser = new OOXMLParser (); 
	      msofficeparser.parse(MSEinputstream, handler, metadata,pcontext);

	      System.out.println("Contents of the document:" + handler.toString());
	      System.out.println("Metadata of the document:");
	      String[] metadataNamesMS = metadata.names();
	      
	      for(String name : metadataNamesMS) {
	         System.out.println(name + ": " + metadata.get(name));
	      }

	      handler = new BodyContentHandler();
	      metadata = new Metadata();
	      
//	      FileInputStream OOinputstream = new FileInputStream(new File("C:\\Dev\\Gitlab\\repo\\MSExcelParse\\test1.ods"));
	      FileInputStream OOinputstream = new FileInputStream(new File(".\\test1.ods"));

	      // Open Office Parser
	      OpenDocumentParser openofficeparser = new OpenDocumentParser();
	      openofficeparser.parse(OOinputstream, handler, metadata, pcontext);
	      
	      System.out.println("Contents of the document:" + handler.toString());
	      System.out.println("Metadata of the document:");
	      String[] metadataNames = metadata.names();
	      
	      for(String name : metadataNames) {
	         System.out.println(name + ": " + metadata.get(name));
	      }
	}

}
