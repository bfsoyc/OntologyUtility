import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.commons.lang3.ObjectUtils.Null;
import org.apache.jena.ontology.OntClass;
import org.apache.jena.ontology.OntModel;
import org.apache.poi.xwpf.usermodel.*;

public class WordDocWriter {
	public static void writeDoc( OntModel m, String templatePath, String filePath){
		XWPFDocument temp = null;
		try{
			temp = new XWPFDocument(new FileInputStream(templatePath));
		}catch (Exception e) {
			System.out.println( templatePath + ":file not found");
		}
		

		XWPFTable tempTbl =  temp.getTables().get(0);
		IBody ibody = tempTbl.getBody();
		XWPFDocument doc = new XWPFDocument();
		
		int classCnt = m.listClasses().toList().size();
		for( int i = 0; i < classCnt; i++ ){
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
							if (text.contains("URIPos")) {
								text = text.replace("URIPos", ontc.getURI());
								r.setText(text,0);// the pos argument can not be omitted, or the text will be appended to the end 
							}
							else if (text.contains("CommentZhPos")) {
								String commentStr = ontc.getComment("zh");
								if( commentStr == null )
									commentStr = "";
								text = text.replace("CommentZhPos", commentStr);
								r.setText(text,0);// the pos argument can not be omitted, or the text will be appended to the end 
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
