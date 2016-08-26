import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.IllegalFormatCodePointException;
import java.util.List;
import java.util.Map;

import org.apache.jena.ontology.DatatypeProperty;
import org.apache.jena.ontology.ObjectProperty;
import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.ontology.OntModelSpec;
import org.apache.jena.ontology.OntProperty;
import org.apache.jena.ontology.OntResource;
import org.apache.jena.rdf.model.ModelFactory;
import org.apache.jena.rdf.model.Property;
import org.apache.jena.rdf.model.Resource;
import org.apache.jena.sparql.function.library.max;
import org.apache.jena.util.iterator.ExtendedIterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDocWriter {
	private static Integer idxOffset, cHierarchyDepth, pHierarchyDepth;
	public static void writeDoc( String ontologyPath, String dirPath){
		File file = new File(dirPath);
		file.mkdirs();
		
		OntModel m = ModelFactory.createOntologyModel(OntModelSpec.OWL_DL_MEM_RDFS_INF);;
		try {
			m.read( new FileInputStream(ontologyPath), "");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println("Invalid ontology file path");
		}
		//m.read(ontologyPath);
		
		
		Map<String, String> NsPrefixMap = m.getNsPrefixMap();
		for( String key:NsPrefixMap.keySet() ){
			System.out.println("prefix: "+key+" Ns: "+NsPrefixMap.get(key));
		}
		
		//----the first file is about classes----
		Workbook wb = new XSSFWorkbook();  
		Sheet sheet = wb.createSheet("classes");
		String[] headers = {"Classes", "label xml:lang=\"zh\"",	"label xml:lang=\"en\"", "definition xml:lang=\"zh\"",
				"comment xml:lang=\"zh\"", "URI", "versionInfo", "isDefinedBy", "disjointWith", "equivalentClass"
		};
		// column index of each header, which is predefined
		Map<String, Integer> headerMap = new HashMap<String, Integer>();
		Row Header = sheet.createRow(0);
		for( int i = 0 ; i < headers.length; i++ ){
			headerMap.put( headers[i], i);
			Header.createCell(i).setCellValue(headers[i]);
		}
		
		//int ClassesIdx = 0, LabelZhIdx = 1, LabelEnIdx = 2, CommentZhIdx = 4, URIIdx = 5, IsDefinedByIdx = 7;
		idxOffset = 0;
		cHierarchyDepth = 0;
	    
		List<OntClass> rootClassList = m.listHierarchyRootClasses().toList();
		for( int i = 0; i < rootClassList.size(); i++ ){
			_dfs_c( sheet, m, rootClassList.get(i), headerMap, 0);
		}
		for( int i = 0 ; i < Header.getLastCellNum(); i++ )
			sheet.autoSizeColumn(i);
		
		FileOutputStream fileOut;
		try {
			fileOut = new FileOutputStream(dirPath + "\\" + "classesDoc.xlsx");
			wb.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		
		
		//----the second file is about data type properties----
		wb = new XSSFWorkbook();
		sheet = wb.createSheet("datatypeProperties");
		String[] newHeaders = {"Properties", "Domains", "Ranges", "label xml:lang=\"zh\"",	"label xml:lang=\"en\"", "definition xml:lang=\"zh\"",
				"comment xml:lang=\"zh\"", "URI", "versionInfo", "isDefinedBy", "disjointWith", "equivalentClass"
		};
		headers = newHeaders;
		// column index of each header, which is predefined
		headerMap = new HashMap<String, Integer>();
		Header = sheet.createRow(0);
		for( int i = 0 ; i < headers.length; i++ ){
			headerMap.put( headers[i], i);
			Header.createCell(i).setCellValue(headers[i]);
		}
		
		idxOffset = 0;
		pHierarchyDepth = 0;
	    
		List<DatatypeProperty> dPropertyList = m.listDatatypeProperties().toList();
		for( int i = 0; i < dPropertyList.size(); i++ ){
			List<? extends OntProperty> superP =  dPropertyList.get(i).listSuperProperties().toList();
			
			
			if( superP.size() > 0 ){
				continue;
			}
			_dfs_p( sheet, m, dPropertyList.get(i), headerMap, 0);
		}
		
		try {
			fileOut = new FileOutputStream(dirPath + "\\" + "datatypePropertiesDoc.xlsx");
			wb.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		//----the third file is about object properties----
		// almost the same as data type properties
		wb = new XSSFWorkbook();
		sheet = wb.createSheet("objectProperties");
		Header = sheet.createRow(0);
		for( int i = 0 ; i < headers.length; i++ ){
			Header.createCell(i).setCellValue(headers[i]);
		}

		
		idxOffset = 0;
		pHierarchyDepth = 0;
	    
		List<ObjectProperty> oPropertyList = m.listObjectProperties().toList();
		for( int i = 0; i < oPropertyList.size(); i++ ){
			List<? extends OntProperty> superP =  oPropertyList.get(i).listSuperProperties().toList();
				
			if( superP.size() > 0 ){
				continue;
			}
			_dfs_p( sheet, m, oPropertyList.get(i), headerMap, 0);
		}
		
		try {
			fileOut = new FileOutputStream(dirPath + "\\" + "objectPropertiesDoc.xlsx");
			wb.write(fileOut);
			fileOut.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static void _dfs_c( Sheet sheet, OntModel m, OntClass ontc, Map<String, Integer> headerMap, int deep) {
		if( deep > cHierarchyDepth ){
			cHierarchyDepth++;
			idxOffset++;
			for( Row r:sheet ){
				for( int j = r.getLastCellNum()-1; j >= headerMap.get("Classes")+idxOffset; j--){
					if( r.getCell(j)!=null && r.getCell(j).getStringCellValue()!= null ){
						Cell copyCell = r.createCell(j+1);
						copyCell.setCellValue(r.getCell(j).getStringCellValue());
					}
				}
				r.createCell(headerMap.get("Classes")+idxOffset);// create an empty cell to replace the original cell( same as delete the original cell)
			}
		}
		
		Row row = sheet.createRow(sheet.getLastRowNum()+1);
		for( String key: headerMap.keySet()){						
			if( key .equals("Classes")){
				Cell cell = row.createCell(headerMap.get(key)+deep);
				//Cell cell = row.getCell(headerMap.get(key)+deep);
				cell.setCellValue(ontc.getLocalName());
			}
			else if( key .equals("label xml:lang=\"zh\"")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontc.getLabel("zh")!=null )
					cell.setCellValue(ontc.getLabel("zh"));
			}
			else if( key .equals("label xml:lang=\"en\"")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontc.getLabel("en")!=null )
					cell.setCellValue(ontc.getLabel("en"));
			}
			else if( key .equals("comment xml:lang=\"zh\"")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontc.getComment("zh")!=null )
					cell.setCellValue(ontc.getComment("zh"));
			}
			else if( key .equals("URI")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				
				if( m.getNsURIPrefix(ontc.getNameSpace())==null ){
					System.out.println("There is no defined prefix for the namespace of URI:"+ontc.getURI());
					cell.setCellValue( ontc.getURI() );
				}
				else
					cell.setCellValue(
							m.getNsURIPrefix(ontc.getNameSpace())+":"+ontc.getLocalName());
			}
			else if( key .equals("versionInfo")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontc.getVersionInfo()!=null)
					cell.setCellValue(ontc.getVersionInfo());
			}
			else if( key .equals("isDefinedBy")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontc.getIsDefinedBy()!=null)
					cell.setCellValue(ontc.getIsDefinedBy().getURI());
			}
		}
		
		List<OntClass> subOntc = ontc.listSubClasses().toList();
		if( subOntc.size() > 0 ){			
			for( int i = 0 ; i < subOntc.size(); i++ ){
				_dfs_c(sheet, m, subOntc.get(i), headerMap, deep+1);
			}
		}
					
	}
	
	private static String URIParser( ){
		
		return null;
	}
	
	private static void _dfs_p( Sheet sheet, OntModel m, OntProperty ontdp, Map<String, Integer> headerMap, int deep) {
		if( deep > pHierarchyDepth ){
			pHierarchyDepth++;
			idxOffset++;
			for( Row r:sheet ){
				for( int j = r.getLastCellNum()-1; j >= headerMap.get("Properties")+idxOffset; j--){
					if( r.getCell(j)!=null && r.getCell(j).getStringCellValue()!= null ){
						Cell copyCell = r.createCell(j+1);
						copyCell.setCellValue(r.getCell(j).getStringCellValue());
					}
				}
				r.createCell(headerMap.get("Properties")+idxOffset);// create an empty cell to replace the original cell( same as delete the original cell)
			}
		}
		
		Row row = sheet.createRow(sheet.getLastRowNum()+1);
		for( String key: headerMap.keySet()){						
			if( key .equals("Properties")){
				Cell cell = row.createCell(headerMap.get(key)+deep);
				cell.setCellValue(ontdp.getLocalName());
			}
			else if( key .equals("Domains")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				List<? extends OntResource> domains = ontdp.listDomain().toList();
				String domainNames = "";
				for( int j = 0 ; j < domains.size(); j++ ){
					if(j > 0)
						domainNames = domainNames.concat(" ");
					if( m.getNsURIPrefix(domains.get(j).getNameSpace())==null ){
						System.out.println("There is no defined prefix for the namespace of URI:"+domains.get(j).getURI());
						domainNames = domainNames.concat(domains.get(j).getURI());
					}
					else
						domainNames = domainNames.concat(
								m.getNsURIPrefix(domains.get(j).getNameSpace())+":"+domains.get(j).getLocalName());				
				}
				cell.setCellValue(domainNames);
			}
			else if( key .equals("Ranges")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				List<OntClass> ranges = (List<OntClass>) ontdp.listRange().toList();
				String rangeNames = "";
				for( int j = 0 ; j < ranges.size(); j++ ){
					if(j > 0)
						rangeNames = rangeNames.concat(" ");
					if( m.getNsURIPrefix(ranges.get(j).getNameSpace())==null ){
						System.out.println("There is no defined prefix for the namespace of URI:"+ranges.get(j).getURI());
						rangeNames = rangeNames.concat(ranges.get(j).getURI());	
					}
					else
						rangeNames = rangeNames.concat(
								m.getNsURIPrefix(ranges.get(j).getNameSpace())+":"+ranges.get(j).getLocalName());					
				}
				cell.setCellValue(rangeNames);
			}
			else if( key .equals("label xml:lang=\"zh\"")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontdp.getLabel("zh")!=null )
					cell.setCellValue(ontdp.getLabel("zh"));
			}
			else if( key .equals("label xml:lang=\"en\"")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontdp.getLabel("en")!=null )
					cell.setCellValue(ontdp.getLabel("en"));
			}
			else if( key .equals("comment xml:lang=\"zh\"")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontdp.getComment("zh")!=null )
					cell.setCellValue(ontdp.getComment("zh"));
			}
			else if( key .equals("URI")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				cell.setCellValue(ontdp.getURI());
				if( m.getNsURIPrefix(ontdp.getNameSpace())==null ){
					System.out.println("There is no defined prefix for the namespace of URI:"+ontdp.getURI());
					cell.setCellValue( ontdp.getURI() );
				}
				else
					cell.setCellValue(
							m.getNsURIPrefix(ontdp.getNameSpace())+":"+ontdp.getLocalName());
			}
			else if( key .equals("versionInfo")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontdp.getVersionInfo()!=null)
					cell.setCellValue(ontdp.getVersionInfo());
			}
			else if( key .equals("isDefinedBy")){
				Cell cell = row.createCell(headerMap.get(key)+idxOffset);
				if( ontdp.getIsDefinedBy()!=null)
					cell.setCellValue(ontdp.getIsDefinedBy().getURI());
			}
		}
		
		List<? extends OntProperty> subOntdp =  ontdp.listSubProperties().toList();
		if( subOntdp.size() > 0 ){			
			for( int i = 0 ; i < subOntdp.size(); i++ ){
				// a property is sub-property of itself, we need to exclude itself from the sub-property list.
				if( subOntdp.get(i).getURI().equals(ontdp.getURI()) )
					continue;
				_dfs_p(sheet, m, subOntdp.get(i), headerMap, deep+1);
			}
		}
					
	}
}
