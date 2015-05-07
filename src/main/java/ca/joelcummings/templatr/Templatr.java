package ca.joelcummings.templatr;

import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.ClassFinder;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Text;
import org.docx4j.wml.Tr;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;

import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

import javax.xml.bind.JAXBException;
import javax.xml.bind.JAXBElement;

/**
 * @author Joel Cummings
 * @version 1.0
 */
public class Templatr {

	private String wordFile, jsonFile;
	private JSONArray items;
	private MainDocumentPart documentPart;
	private final static String TEXT_PATH = "//w:t";
	private WordprocessingMLPackage wordMLPackage;

	// CONSTANTS

	private final static String VALUE_KEY = "value";
	private final static String TYPE_KEY = "type";
	private final static String TEXT_KEY = "text";
	private final static String TABLE_KEY = "table";
	private final static String LIST_KEY = "list";
	private final static String KEY_PLACEHOLDER = "placeholder";

	public Templatr(String wordFile, String jsonFile) throws ParseException,
			IOException, Docx4JException, JAXBException {
		this.wordFile = wordFile;
		this.jsonFile = jsonFile;

		init();
		replaceElements();
		// addTableAtPosition();
	}

	private void init() throws IOException, ParseException, Docx4JException {

		JSONParser parser = new JSONParser();

		JSONObject result = (JSONObject) parser.parse(new FileReader(new File(
				jsonFile)));
		items = (JSONArray) result.get("items");

		wordMLPackage = WordprocessingMLPackage.load(new File(wordFile));

		documentPart = wordMLPackage.getMainDocumentPart();

	}

	private void replaceElements() throws JAXBException, Docx4JException {

		HashMap<String, String> replacements = new HashMap<String, String>();

		for (Object o : items) {
			if (o instanceof JSONObject) {
				JSONObject obj = (JSONObject) o;

				String type = (String) obj.get(TYPE_KEY);

				if (type.equals(TEXT_KEY)) {
					// replacements.put((String)obj.get(KEY_PLACEHOLDER),
					// (String)obj.get(VALUE_KEY));

					this.replaceTextInParagraph(
							(String) obj.get(KEY_PLACEHOLDER),
							(String) obj.get(VALUE_KEY));
				} else if (type.equals(LIST_KEY)) {
					insertList(obj);
				} else if (type.equals(TABLE_KEY)){
					Tbl table = this.createTable(obj);
					int index = this.findIndexOfText((String)obj.get(KEY_PLACEHOLDER));
					insertObject(index, table);
					documentPart.getContent().remove(index +1); // remove the old paragraph with the placeholder
				}
				

			}
		}
	}

	private void insertObject(int index, Object obj) {
		if (index < this.documentPart.getContent().size()) {
			documentPart.getContent().add(index, obj);
		} else {
			documentPart.getContent().add(obj);
		}
	}
	
	private void insertList(JSONObject list) {
		
		int index = this.findIndexOfText((String)list.get(KEY_PLACEHOLDER));
		int newIndex = index;
		this.replaceTextInParagraph((String)list.get(KEY_PLACEHOLDER), "");
		JSONArray elements = (JSONArray)list.get(VALUE_KEY);
		
		for (Object obj: elements) {
			
			JSONObject o = (JSONObject)obj;
			
			String type = (String)o.get(TYPE_KEY);
			
			if (type.equals(TEXT_KEY)) {
				P par = createParagraph((String)o.get(VALUE_KEY));
				// add to the the existing paragraph instead of creating a new one
				if (newIndex == index) {
					insertRun(getParagraph(index), (String)o.get(VALUE_KEY));
				} else if (newIndex < documentPart.getContent().size()) {					
					this.documentPart.getContent().add(newIndex, par);
				} else {
					documentPart.getContent().add(par);
				}
			} 
			else if(type.equals(TABLE_KEY)) {
				Tbl table = createTable(o);
				if (newIndex < this.documentPart.getContent().size()) {
					this.documentPart.getContent().add(newIndex, table);
				} else {
					this.documentPart.getContent().add(table);
				}
				
			}
			
			newIndex++;
			
		}
		
		
		
		
	}
	
	private void insertRun(P par, String text) {
		
		ObjectFactory factory = new ObjectFactory();
		R run = factory.createR();
		Text txt = factory.createText();
		txt.setValue(text);
		run.getContent().add(txt);
		
		par.getContent().add(run);
		
	}
	
	private P createParagraph(String data) {
		
		
		ObjectFactory factory = new ObjectFactory();
		P par = factory.createP();
		
		R run = factory.createR();
		Text txt = factory.createText();
		txt.setValue(data);
		run.getContent().add(txt);
		par.getContent().add(run);
		
		return par;
		
	}
	
	
	

	private Tbl createTable(JSONObject tableObj) {
		ObjectFactory factory = new ObjectFactory();
		Tbl table = factory.createTbl();
		JSONArray items = (JSONArray)tableObj.get(VALUE_KEY);
		JSONObject head = (JSONObject)((JSONObject)items.get(0)).get("columns");
		Tr headRow = factory.createTr();
		Tr row;
		
		// Since HashSets are not ordered we must sort it based on the order specified in the JSON
		Object[] sortedKeySet = head.keySet().toArray();
		Arrays.sort(sortedKeySet);
		
		//create header
		for (Object key: sortedKeySet) {
			
			String s = (String)head.get(key);
			Tc cell = factory.createTc();
			Text t = factory.createText();
			t.setValue(s);
			cell.getContent().add(documentPart.createParagraphOfText(s));
			headRow.getContent().add(cell);
		}
		
		table.getContent().add(headRow);
		
		// remove it to loop through the body;
		items.remove(0);
		
		for (Object obj: items) {
			JSONObject o = (JSONObject)obj;
			row = factory.createTr();
			for (Object key: sortedKeySet) {
				Tc cell = factory.createTc();
				Object value = head.get(key);
				String s = (String)((JSONObject)o.get("row")).get(value);
				Text t = factory.createText();
				t.setValue(s);
				cell.getContent().add(documentPart.createParagraphOfText(s));
				row.getContent().add(cell);
			}
			table.getContent().add(row);
		}
		
		return table;
		
	}
	

	private int findIndexOfText(String toFind) {
		int index = 0;

		for (Object o : documentPart.getContent()) {

			if (o instanceof P) {
				P par = (P) o;

				String content = par.toString();

				if (content.contains(toFind)) {

					index = documentPart.getContent().indexOf(par);
					break;

				}

			}

		}

		return index;
	}

	private void replaceTextInParagraph(String toFind, String toReplace) {

		int index = findIndexOfText(toFind);

		P par = (P) documentPart.getContent().get(index);

		for (Object o : par.getContent()) {

			if (o instanceof R) {
				R run = (R) o;

				for (Object inner : run.getContent()) {

					Text t = (Text) ((JAXBElement) inner).getValue();
					String s = t.getValue();
					
					if (s.contains(toFind)) {
						s = s.replace(toFind, toReplace);						
						t.setValue(s);
					}
				}
			}

		}

	}
	
	private P getParagraph(int index) {
		return (P)documentPart.getContent().get(index);
	}

	public void saveDocument(String fileName) throws Docx4JException {
		wordMLPackage.save(new File(fileName));
	}

	public static void main(String[] args) {

		String wordFile = "TemplatrTest.docx";
		String jsonFile = "input.json";

		try {
			Templatr templatr = new Templatr(wordFile, jsonFile);

			templatr.saveDocument("Document test copy.docx");

		} catch (ParseException e) {
			// e.printStackTrace();
			System.out.println("Parse Exception: " + e.getMessage());
		} catch (IOException e) {
			// e.printStackTrace();
			System.out.println("IO Exception: " + e.getMessage());
		} catch (Docx4JException e) {
			// e.printStackTrace();
			System.out.println("DDCX4J EXCEPTION: " + e.getMessage());
		} catch (JAXBException e) {
			System.out.println("JAXB EXCEPTION: " + e.getMessage());
			// e.printStackTrace();
		}

	}

}
