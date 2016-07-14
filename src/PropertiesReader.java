import java.io.File;
import java.io.IOException;
import java.security.cert.CertPathChecker;
import java.util.HashMap;
import java.util.Map;
import java.util.Iterator;
import java.util.Stack;

import org.apache.jena.ext.com.google.common.base.CaseFormat;
import org.apache.jena.ontology.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.property.PropertyTable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

// the ClassReader use for reading the classes definition file 
// using Apache POI to handling the .xlsx file
public class PropertiesReader {
	static final int ObjectProperty = 1, DatatypeProperty = 2;
	
	public static void read( OntModel m, String filePath ){
		// open file
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(new File(filePath));
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println( filePath + ":file not found");
			e.printStackTrace();
		}
		
		// column index of each header, which is predefined
		int PropertiesIdx = -1, URIIdx = -1, LabelZhIdx = -1, LabelEnIdx = -1, DomainsIdx = -1, RangesIdx = -1;
		// read the header
		Sheet sheet = wb.getSheetAt(0);
		Row header = sheet.getRow(0);
		for( Cell cell : header){
			String s = cell.getStringCellValue();
			int col = cell.getColumnIndex();
			if( s.equals("Properties") )
				PropertiesIdx = col;		
			else if( s.equals("URI") )
				URIIdx = col;
			else if( s.equals("label xml:lang=\"zh\"") )
				LabelZhIdx = col;
			else if( s.equals("label xml:lang=\"en\"") )
				LabelEnIdx = col;			
			else if( s.equals("Domains") )
				DomainsIdx = col;
			else if( s.equals("Ranges") )
				RangesIdx = col;
		}
		// URI is required to build the ontology
		try{
			assert URIIdx != -1;		
		}
		catch (Exception e) {
			System.out.println("URI is required.\n Note that the URI header should be equals to"
					+ " \"URI\" exactly without extra blank space"); 
		}
		// for properties, our ontology model require non-empty domain
		try{
			assert DomainsIdx != -1;
		}
		catch (Exception e) {
			System.out.println("domains of the properties is required for our ontology model.\n Note that the domains header should be equals to"
					+ " \"Domains\" exactly without extra blank space");
		}
			
		// the last element of a class URI should be consistent with its class name, otherwise
		// it's inconvenient to construct the map form class name to its URI.
		Map<String, String> URIMap = new HashMap<String, String>();
		for( Iterator<OntClass> i = m.listClasses(); i.hasNext() ;){
			OntClass ontc = i.next();
			String classURI = ontc.getURI();
			String splitURI[] = classURI.split("[\\,:,#,/]",-1); // to be debugged
			String className = splitURI[splitURI.length-1];
			URIMap.put(className, classURI);
		}
		
		// create properties for the ontology model, one property per row
		// the properties are organized in hierarchical structure.
		Stack<Integer> propertyLevel = new Stack<Integer>();
		Stack<OntProperty> superProperty = new Stack<OntProperty>();
		int rowEnd = sheet.getLastRowNum();
		for( int r = 1; r < rowEnd; r++ ){
			Row row = sheet.getRow(r);
			if( row == null ){
				// ignore empty row
				continue;
			}
			
			// the first non-empty cell will be treated as the name of property.
			int colEnd = row.getLastCellNum();
			String propertyName = null;
			for( int c = 0; c <= colEnd; c++){
				Cell cell = row.getCell(c);
				if( cell != null ){
					propertyName = cell.getStringCellValue();
					while( !propertyLevel.empty() && c <= propertyLevel.peek().intValue() ){
						propertyLevel.pop();
						superProperty.pop();						
					}
					propertyLevel.push(c);
					break;
				}
			}
			if( propertyName == null ) continue;
			
			OntProperty ontp = null;
			try{
				ontp = m.createDatatypeProperty( row.getCell(URIIdx).getStringCellValue() );	
			}
			catch (Exception e) {
				System.out.println("Please check URI of property \""+propertyName+"\" on row "+(r+1));
			}
			//String curURI = ontp.getURI();
			//System.out.println("URI = "+curURI);
			if( !superProperty.empty() )
				ontp.addSuperProperty(superProperty.peek());
			superProperty.push(ontp);
			
			
			// add domain 
			try{
				String domainsStr = row.getCell(DomainsIdx).getStringCellValue().replaceAll("\\sor\\s", " ");
				String domains[] = domainsStr.split("\\s+"); // Splitter as one or more space
				for( String s : domains){
					OntClass domn = m.getOntClass(URIMap.get(s));
					ontp.addDomain(domn);					
				}				
			}catch (Exception e) {
				System.out.println("Please check the Domians of property \""+propertyName+"\" on row "+(r+1));
				System.out.println("If not error found, check the URI of the correspoding class in classes fifinition files instead");
			}
			
			if( RangesIdx!=-1 && row.getCell(RangesIdx)!=null ){
				try{
					String rangesStr = row.getCell(RangesIdx).getStringCellValue().replaceAll("\\sor\\s", " ");
					String ranges[] = rangesStr.split("\\s+",-1);
					for( String s : ranges){
						// domain with name leaded by capital letter is treated as a Class.
						// infer that the property is object property.
						char sArray[] = s.toCharArray();
						if( sArray[0] >= 'A' && sArray[0] <= 'Z' ){
							ontp.convertToObjectProperty();
							OntClass rage = m.getOntClass(URIMap.get(s));
							ontp.addRange(rage);
						}					
					}
				}catch (Exception e) {
					System.out.println("Please check the Ranges of property \""+propertyName+"\" on row "+(r+1));
					System.out.println("If not error found, check the URI of the correspoding class in classes fifinition files instead");
				}
			}
			
			if( LabelZhIdx!=-1 && row.getCell(LabelZhIdx)!=null ){
				ontp.addLabel(row.getCell(LabelZhIdx).getStringCellValue(), "zh");
			}
			if( LabelEnIdx!=-1 && row.getCell(LabelEnIdx)!=null ){
				ontp.addLabel(row.getCell(LabelEnIdx).getStringCellValue(), "en");
			}
			
		}
	}
}
