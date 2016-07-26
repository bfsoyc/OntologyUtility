


import java.io.*;
import org.apache.jena.ontology.*;
import org.apache.jena.rdf.model.ModelFactory;

public class Demo {
	public static void main(String args[]) throws FileNotFoundException{
		//System.out.println("hello jena");
		//String SOURCE = "http://www.OMG.org/paradise/relico";
		//String NS = SOURCE + "#";
		OntModel m = ModelFactory.createOntologyModel();
		
		String classDefintionFile = "D:\\Users\\OuYuanchang\\destop\\单位工作\\20150819_文物数字化保护标准体系及关键标准研究与示范"
				+ "\\virtual\\virtual\\ontology\\classes.xlsx";		
		ClassesReader.read( m, classDefintionFile);
		
		String datatypePropertyDefintionFile = "D:\\Users\\OuYuanchang\\destop\\单位工作\\20150819_文物数字化保护标准体系及关键标准研究与示范"
				+ "\\virtual\\virtual\\ontology\\Properties0710.xlsx";
		PropertiesReader.read(m, datatypePropertyDefintionFile);
		
		try {

			  File file= new File("JenaRelico.rdf");
			  m.write(new FileOutputStream(file));

		} catch (IOException e) {
			  e.printStackTrace();
		}
		
		String templatePath = "template.docx";
		String docSavePath = "ontologyDoc.docx";
		WordDocWriter.writeDoc(m, templatePath, docSavePath);
		
		System.out.println("endOfmain");
	}
}
