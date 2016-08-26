


import java.io.*;
import org.apache.jena.ontology.*;
import org.apache.jena.rdf.model.Model;
import org.apache.jena.rdf.model.ModelFactory;
import org.apache.jena.vocabulary.XSD;
import java.util.HashMap;
import java.util.Map;

public class Demo {
	public static void main(String args[]) throws FileNotFoundException{
		
//		//System.out.println("hello jena");
//		String SOURCE = "http://www.xxx.org.cn/standards/relic/ontology/core";	
//		String NS = SOURCE + "#";
//		String yourOntologyNsPrefix = "relico"; // relico is short for the name space NS. 
//				
//		Map<String, String> NsMap = new HashMap<String, String>(){{
//			put(yourOntologyNsPrefix, NS);
//			put("skos","http://www.w3.org/2004/02/skos/core#");
//			/*  the listing useful prefix are already exist in default Ontmodel:
//				owl for http://www.w3.org/2002/07/owl#
//				rdf for http://www.w3.org/1999/02/22-rdf-syntax-ns#
//				xsd for http://www.w3.org/2001/XMLSchema#
//				rdfs for http://www.w3.org/2000/01/rdf-schema#			
//			*/
//			
//			// To set customized prefix, use the following code:  
//			// put("xsd", "http://www.w3.org/2001/XMLSchema#");
//		}};
//		
//		OntModel m = ModelFactory.createOntologyModel();
//		m.createOntology(SOURCE);
//		// use a Map container to set(add) all prefixes in once
//		m.setNsPrefixes(NsMap);	
//		// use method setNsPrefix to set prefixes one by one
//		// m.setNsPrefix(yourOntologyNsPrefix, NS); // Relico is short for the name space NS. 
//		m.setNsPrefix("defaultNs", NS); // the default prefix recognized by other class 
//
//		
//		String classDefinitionFile = "resource\\classes.xlsx";		
//		ClassesReader.read( m, classDefinitionFile);
//		
//		String datatypePropertyDefinitionFile = "resource\\DatatypeProperties.xlsx";
//		String objectPropertyDefinitionFile = "resource\\ObjectProperties.xlsx";
//		PropertiesReader.read(m, datatypePropertyDefinitionFile);
//		PropertiesReader.read(m, objectPropertyDefinitionFile);
//		
//		String individualDefinitionFile = "resource\\enumeration.xlsx";
//		IndividualsReader.read(m, individualDefinitionFile);
//		
//		try {
//
//			  File file= new File("JenaRelico.rdf");
//			  m.write(new FileOutputStream(file));
//
//		} catch (IOException e) {
//			  e.printStackTrace();
//		}
//		
//		String templatePath = "template.docx";
//		String docSavePath = "ontologyDoc.docx";
//		//WordDocWriter.writeDoc(m, templatePath, docSavePath);
//		
		String ontologyPath = "resource\\radarOnto.owl";
		String dirPath = "resource\\generateExcel" ;
		ExcelDocWriter.writeDoc(ontologyPath, dirPath);
		
		System.out.println("endOfmain");
	}
}
