import java.io.File;
import java.io.IOException;
import java.util.Map;
import java.util.Stack;

import org.apache.jena.ontology.Individual;
import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.ontology.OntProperty;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class IndividualsReader {
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
		int IndividualsIdx = -1, URIIdx = -1, LabelZhIdx = -1, LabelEnIdx = -1, CommentZhIdx = -1, ClassesIdx = -1;
		// read the header
		Sheet sheet = wb.getSheetAt(0);
		Row header = sheet.getRow(0);
		for( Cell cell : header){
			String s = cell.getStringCellValue();
			int col = cell.getColumnIndex();
			if( s.equals("Individuals") )
				IndividualsIdx = col;		
			else if( s.equals("URI") )
				URIIdx = col;
			else if( s.equals("label xml:lang=\"zh\"") )
				LabelZhIdx = col;
			else if( s.equals("label xml:lang=\"en\"") )
				LabelEnIdx = col;			
			else if( s.equals("comment xml:lang=\"zh\""))
				CommentZhIdx = col;
			else if( s.equals("Classes"))
				ClassesIdx = col;
		}
		
		Map<String,String> prefixMap = m.getNsPrefixMap();
		int rowEnd = sheet.getLastRowNum();
		for( int r = 1; r <= rowEnd; r++ ){
			Row row = sheet.getRow(r);
			if( row == null ){
				// ignore empty row
				continue;
			}
			
			String ClassName = row.getCell(ClassesIdx).getStringCellValue();
			String ClassURI = prefixMap.get("defaultNs") + ClassName;
			OntClass ontc = m.getOntClass( ClassURI );
			
			String indURI = null; // what if the cell is null?
			if( row.getCell(URIIdx)!=null ){
				indURI = row.getCell(URIIdx).getStringCellValue();
			}
			else{
				// if indURI is not provided in advanced, use the default way to construct one.
				// noted that it's not promise that the URI is unique to the one of classes or properties.
				indURI = prefixMap.get("defaultNs") + row.getCell(IndividualsIdx).getStringCellValue();
			}			
			Individual ind = ontc.createIndividual(indURI);
			
			if( LabelZhIdx!=-1 && row.getCell(LabelZhIdx)!=null ){
				ind.addLabel(row.getCell(LabelZhIdx).getStringCellValue(), "zh");
			}
			if( LabelEnIdx!=-1 && row.getCell(LabelEnIdx)!=null ){
				ind.addLabel(row.getCell(LabelEnIdx).getStringCellValue(), "en");
			}
			if( CommentZhIdx!=-1 && row.getCell(CommentZhIdx)!=null ){
				ind.setComment(row.getCell(CommentZhIdx).getStringCellValue(), "zh");
			}
		}
	}
}
