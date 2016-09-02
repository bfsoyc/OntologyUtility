import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Iterator;
import java.util.Stack;


import org.apache.jena.ontology.*;
import org.apache.jena.rdf.model.*;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
			
		/*
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
		Obsolete---
		*/
		Map<String,String> prefixMap = m.getNsPrefixMap();
		
		// create properties for the ontology model, one property per row
		// the properties are organized in hierarchical structure.
		Stack<Integer> propertyLevel = new Stack<Integer>();
		Stack<OntProperty> superProperty = new Stack<OntProperty>();
		int rowEnd = sheet.getLastRowNum();
		for( int r = 1; r <= rowEnd; r++ ){
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
				ontp = m.createOntProperty( row.getCell(URIIdx).getStringCellValue());
				//ontp.asDatatypeProperty();
			}
			catch (Exception e) {
				System.out.println("Please check URI of property \""+propertyName+"\" on row "+(r+1));
			}
			//String curURI = ontp.getURI();
			//System.out.println("URI = "+curURI);
			if( !superProperty.empty() )
				ontp.addSuperProperty(superProperty.peek());
			superProperty.push(ontp);
			//ontp.asDatatypeProperty();
			
			// add domain 
			try{
				String domains[] = row.getCell(DomainsIdx).getStringCellValue().split("\\s+");
				OntClass doms = null;
				
				try{
					assert domains.length%2 == 1;
				}catch( Exception e ){
					System.out.println("Invalid Domain expression");
				}
				
				for( int i = 0 ; i < domains.length ; i+=2 ){
					String relation = "or";
					if( i > 0 ){
						relation = domains[i-1];
					}
					String partedURI[] = domains[i].split(":");
					// the part before notation ":" should be the prefix, the last part should be the local name
					String fullURI = null;
					if( partedURI.length > 1 ){
						try{
							assert prefixMap.get(partedURI[0])!=null;
						}
						catch (Exception e) {
							System.out.println("Unknown prefix:"+partedURI[0]);
						}						
						fullURI = prefixMap.get(partedURI[0]) + partedURI[1];	
					}					
					else{
						fullURI = prefixMap.get("defaultNs") + partedURI[0];
					}
					OntClass dom = m.getOntClass(fullURI);
					if( i == 0){
						doms = dom;
					}
					
					if( relation.equals("or") ){
						UnionClass unionc = m.createUnionClass( null, m.createList( new RDFNode[] { doms, dom } ));;
						doms = unionc.asClass();
					}
					else if(relation.equals("and") ){
						IntersectionClass interc = m.createIntersectionClass( null, m.createList( new RDFNode[] { doms, dom } ));;
						doms = interc.asClass();
					}
					else{
						System.out.println("Invalid Domain expression");
					}
					
				}
				ontp.addDomain(doms);
				/*
				String domainsStr = row.getCell(DomainsIdx).getStringCellValue().replaceAll("\\sor\\s", " ");
				String domains[] = domainsStr.split("\\s+"); // Splitter as one or more space
				for( String s : domains){
					String partedURI[] = s.split(":");
					// the part before notation ":" should be the prefix, the last part should be the local name
					String fullURI = null;
					String shortURI;
					if( partedURI.length > 1 ){
						try{
							assert prefixMap.get(partedURI[0])!=null;
						}
						catch (Exception e) {
							System.out.println("Unknown prefix:"+partedURI[0]);
						}						
						fullURI = prefixMap.get(partedURI[0]) + partedURI[1];	
						shortURI = s;
					}					
					else{
						fullURI = prefixMap.get("defaultNs") + partedURI[0];
						shortURI ="relico:" +  s;
					}
					
					OntClass domn = m.getOntClass(fullURI);
					ontp.addDomain(domn);	
					// !!!!!!!!!!!!!!!!!!!!!
					// the domain of the property will be the intersection of all classes you add to its domain, which is not what we expect
					// so those code are wrong in our case.
					//ontp.convertToDatatypeProperty();
				 
				}
				*/				
			}catch (Exception e) {
				System.out.println("Please check the Domians of property \""+propertyName+"\" on row "+(r+1));
				System.out.println("If not error found, check the URI of the correspoding class in classes definition files instead");
			}
			
			int propertyType = DatatypeProperty;
			if( RangesIdx!=-1 && row.getCell(RangesIdx)!=null ){
				try{

					String ranges[] = row.getCell(RangesIdx).getStringCellValue().split("\\s+");
					
					try{
						assert ranges.length == 1;
					}catch( Exception e ){
						System.out.println("Invalid Range expression: only single range is supported now");
					}					
					
					for( int i = 0 ; i < ranges.length; i+= 2){
						String partedURI[] = ranges[i].split(":");
						String fullURI = null;
						// range with local name leaded by capital letter is treated as a Class.
						// infer that the property is object property.
						char localName[] = null;
						if( partedURI.length > 1 ){
							localName = partedURI[1].toCharArray();
							try{
								assert prefixMap.get(partedURI[0])!=null;
							}
							catch (Exception e) {
								System.out.println("Unknown prefix:"+partedURI[0]);
							}						
							fullURI = prefixMap.get(partedURI[0]) + partedURI[1];
						}
						else {
							localName = partedURI[0].toCharArray();
							fullURI = prefixMap.get("defaultNs")+partedURI[0];
						}
						
						if( localName[0] >= 'A' && localName[0] <= 'Z' ){
							propertyType = ObjectProperty;
								
							OntClass rage = m.getOntClass(fullURI);
							ontp.addRange(rage);							
						}
						else{
							ontp.addRange(ResourceFactory.createResource(fullURI));					
						}
					}
					
				}catch (Exception e) {
					System.out.println("Please check the Ranges of property \""+propertyName+"\" on row "+(r+1));
					System.out.println("If not error found, check the URI of the correspoding class in classes definition files instead");
				}
			}
			if( propertyType == DatatypeProperty)
				ontp = ontp.convertToDatatypeProperty();
				
						
			if( LabelZhIdx!=-1 && row.getCell(LabelZhIdx)!=null ){
				ontp.addLabel(row.getCell(LabelZhIdx).getStringCellValue(), "zh");
			}
			if( LabelEnIdx!=-1 && row.getCell(LabelEnIdx)!=null ){
				ontp.addLabel(row.getCell(LabelEnIdx).getStringCellValue(), "en");
			}
			//ontp.remove();
		}
	}
}
