import java.io.File;
import java.io.IOException;
import java.util.Stack;
import java.util.Map;
import java.util.HashMap;

import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;



// the ClassReader use for reading the classes definition file 
// using Apache POI to handling the .xlsx file
public class ClassesReader {
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
		int ClassesIdx = -1, URIIdx = -1, LabelZhIdx = -1, LabelEnIdx = -1, CommentZhIdx = -1, IsDefinedByIdx = -1;
		// read the header
		Sheet sheet = wb.getSheetAt(0);
		Row header = sheet.getRow(0);
		for( Cell cell : header){
			String s = cell.getStringCellValue();
			int col = cell.getColumnIndex();
			if( s.equals("Classes") )
				ClassesIdx = col;		
			else if( s.equals("URI") )
				URIIdx = col;
			else if( s.equals("label xml:lang=\"zh\"") )
				LabelZhIdx = col;
			else if( s.equals("label xml:lang=\"en\"") )
				LabelEnIdx = col;		
			else if( s.equals("comment xml:lang=\"zh\""))
				CommentZhIdx = col;
			else if( s.equals("isDefinedBy"))
				IsDefinedByIdx = col;
		}
		// URI is required to build the ontology
		try{
			assert URIIdx != -1;
		}
		catch (Exception e) {
			System.out.println("URI is required.\n Note that the header should be equals to"
					+ " \"URI\" exactly without extra blank space"); 
		}
		
		
		// create classes for the ontology model, one class per row
		// the classes are organized in hierarchical structure.
		Stack<Integer> classLevel = new Stack<Integer>();
		Stack<OntClass> superClass = new Stack<OntClass>();
		Map<String, String> prefixMap = m.getNsPrefixMap();
		int rowEnd = sheet.getLastRowNum();
		for( int r = 1; r <= rowEnd; r++ ){
			Row row = sheet.getRow(r);
			if( row == null ){
				// ignore empty row
				continue;
			}
			
			// the first non-empty cell will be treated as the name of class.
			
			
			int colEnd = row.getLastCellNum();
			String className = null;
			for( int c = 0; c <= colEnd; c++){
				Cell cell = row.getCell(c);
				if( cell != null ){
					className = cell.getStringCellValue();
					while( !classLevel.empty() && c <= classLevel.peek().intValue() ){
						classLevel.pop();
						superClass.pop();						
					}
					classLevel.push(c);
					break;
				}
			}
			if( className == null ) continue;
			OntClass ontc = null;
			try{
				String s = row.getCell(URIIdx).getStringCellValue();
				String partedURI[] = s.split(":");
				String shortURI = null;
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
					shortURI =  s;
				}
				else{
					fullURI = prefixMap.get("defaultNs") + partedURI[1];
					shortURI ="relico:" + s;
				}
				
				ontc = m.createClass( fullURI );
			}
			catch (Exception e) {
				System.out.println("Please check URI of class \""+className+"\" on row "+(r+1));
			}
			if( !superClass.empty() )
				ontc.addSuperClass(superClass.peek());
			superClass.push(ontc);
			
			if( LabelZhIdx!=-1 && row.getCell(LabelZhIdx)!=null ){
				ontc.addLabel(row.getCell(LabelZhIdx).getStringCellValue(), "zh");
			}
			if( LabelEnIdx!=-1 && row.getCell(LabelEnIdx)!=null ){
				ontc.addLabel(row.getCell(LabelEnIdx).getStringCellValue(), "en");
			}
			if( CommentZhIdx!=-1 && row.getCell(CommentZhIdx)!=null ){
				ontc.setComment(row.getCell(CommentZhIdx).getStringCellValue(), "zh");
			}
			if( IsDefinedByIdx!=-1 && row.getCell(IsDefinedByIdx)!=null ){
				// isDefinedBy -- A specialisation of seeAlso that is intended to supply a definition of this resource
				// this reference resource is typically identified by an URI. 
				String partedURI[] = row.getCell(IsDefinedByIdx).getStringCellValue().split(":");
				String fullURI;
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
					fullURI = prefixMap.get(partedURI[0]);
				}
				
				ontc.addIsDefinedBy(m.createResource( fullURI ) );
			}
			
		}
		
	}
}
