import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.lang3.ObjectUtils.Null;
import org.apache.jena.ontology.DatatypeProperty;
import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.jena.ontology.OntProperty;
import org.apache.jena.ontology.OntResource;
import org.apache.jena.rdf.model.Resource;
import org.apache.jena.sparql.pfunction.library.listIndex;
import org.apache.poi.xwpf.usermodel.*;

public class WordDocWriter {
	public static void writeDoc( OntModel m, String templatePath, String filePath){
		XWPFDocument temp = null;
		try{
			temp = new XWPFDocument(new FileInputStream(templatePath));
		}catch (Exception e) {
			System.out.println( templatePath + ":file not found");
		}
			
		XWPFDocument doc = new XWPFDocument();
		
		XWPFTable tempTbl =  temp.getTables().get(0); // the first table is the template table of class information
		int classCnt = m.listClasses().toList().size();
		for( int i = 0; i < classCnt; i++ ){
			doc.createTable();
			int pos = doc.getTables().size()-1;
			doc.setTable(pos, tempTbl);
			doc.createParagraph();// create an new paragraph to separate tables, or tables will be concatenated together when you view them on MS
		}
		
		tempTbl = temp.getTables().get(1); // // the second table is the template table of property information	
		List<OntProperty> propertiesList = m.listOntProperties().toList();
		int propertyCnt = propertiesList.size();
		for( int i = 0; i < propertyCnt; i++ ){
			// within our document, only properties defined in our name space are under consideration.  
			String pNS = propertiesList.get(i).getNameSpace();
			if( !pNS.equals("Relico:"))
				continue;
			doc.createTable();
			int pos = doc.getTables().size()-1;
			doc.setTable(pos, tempTbl);
			doc.createParagraph();// create an new paragraph to separate tables, or tables will be concatenated together when you view them on MS
		}
		
		// write the file which contains duplicated template tables and then read it again
		// we do this because there's no convenience method to copy the table
		// and we struggle for better solution
        try {
            doc.write(new FileOutputStream(filePath));
        }catch (Exception e) {
			// TODO: handle exception
		}		
        try {
			doc = new XWPFDocument(new FileInputStream(filePath));
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		
        int cnt = 0;
		for( Iterator<OntClass> i = m.listClasses(); i.hasNext();  ){
			OntClass ontc = i.next();
	        XWPFTable tbl = doc.getTables().get(cnt++);
			
			for(XWPFTableRow row:tbl.getRows()) {				
				for(XWPFTableCell cell:row.getTableCells()){
					for (XWPFParagraph p : cell.getParagraphs()) {
				        for (XWPFRun r : p.getRuns()) {
				        	// !!!! customized development are supposed to be made in this block !!!!				        	
				        	
							String text = r.getText(0);
							String str = null;
							if (text.contains("URIPos")) {
								text = text.replace("URIPos", ontc.getURI());
								r.setText(text,0);// the pos argument can not be omitted, or the text will be appended to the end 
							}
							else if( text.contains("LocalNamePos")){
								str = ontc.getLocalName();
								text = text.replace("LocalNamePos", str);
								r.setText(text,0);
							}
							else if( text.contains("LabelZhPos")){
								str = ontc.getLabel("zh");
								if( str == null )
									str = "";
								text = text.replaceAll("LabelZhPos", str);
								r.setText(text,0);
							}
							else if (text.contains("CommentZhPos")) {
								str = ontc.getComment("zh");
								if( str == null )
									str = "";
								text = text.replace("CommentZhPos", str);
								r.setText(text,0);// the pos argument can not be omitted, or the text will be appended to the end 
							}
							else if( text.contains("SuperClassPos")){
								str = "";
								for( Iterator<OntClass> itr = ontc.listSuperClasses(); itr.hasNext();){
									OntClass superC = itr.next();
									str = str.concat(superC.getLocalName()+" ");
								}
								text = text.replace("SuperClassPos", str);
								r.setText(text,0);
							}
							else if( text.contains("IsDefinedByPos")){
								str = ontc.getIsDefinedBy().getURI();
								if( str == null )
									str ="";
								text = text.replace("IsDefinedByPos", str);
								r.setText(text,0);
							}							
				        }		       
			     	}			     	
			  }			
			}			
		}
		
		
		// the table of properties information is almost the same as class information
		for( Iterator<OntProperty> i = m.listOntProperties(); i.hasNext();  ){
			OntProperty ontp = i.next();
			String pNS = ontp.getNameSpace();
			if( !pNS.equals("Relico:"))
				continue;
			
	        XWPFTable tbl = doc.getTables().get(cnt++);
			
			for(XWPFTableRow row:tbl.getRows()) {				
				for(XWPFTableCell cell:row.getTableCells()){
					for (XWPFParagraph p : cell.getParagraphs()) {
				        for (XWPFRun r : p.getRuns()) {
				        	// !!!! customized development are supposed to be made in this block !!!!				        	
				        	
							String text = r.getText(0);
							String str = null;
							if (text.contains("URIPos")) {
								text = text.replace("URIPos", ontp.getURI());
								r.setText(text,0);// the pos argument can not be omitted, or the text will be appended to the end 
							}
							else if( text.contains("LocalNamePos")){
								str = ontp.getLocalName();
								text = text.replace("LocalNamePos", str);
								r.setText(text,0);
							}
							else if( text.contains("LabelZhPos")){
								str = ontp.getLabel("zh");
								if( str == null )
									str = "";
								text = text.replaceAll("LabelZhPos", str);
								r.setText(text,0);
							}
							else if (text.contains("CommentZhPos")) {
								str = ontp.getComment("zh");
								if( str == null )
									str = "";
								text = text.replace("CommentZhPos", str);
								r.setText(text,0);
							}
							else if( text.contains("DomainPos")){
								str = "";
								for( Iterator<OntResource> itr = (Iterator<OntResource>) ontp.listDomain(); itr.hasNext();){
									OntResource propertyDomain = itr.next();
									str = str.concat(propertyDomain.getLocalName()+" ");
								}
								text = text.replace("DomainPos", str);
								r.setText(text,0);
							}
							else if( text.contains("IsDefinedByPos")){
								Resource ontr = ontp.getIsDefinedBy();
								if( ontr != null )
									str = ontr.getURI();
								if( str == null )
									str ="";
								text = text.replace("IsDefinedByPos", str);
								r.setText(text,0);
							}							
				        }		       
			     	}			     	
			  }			
			}			
		}		
		
		
		// write the file
        try {
            doc.write(new FileOutputStream(filePath));
        }catch (Exception e) {
			// TODO: handle exception
		}
	}
}
