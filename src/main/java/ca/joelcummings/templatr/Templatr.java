package ca.joelcummings.templatr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Drawing;
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
/**
 * @author Joel Cummings
 * @version 1.0
 */
public class Templatr {

	private String wordFile, jsonFile;
	private JSONArray items;
	private MainDocumentPart documentPart;
	private WordprocessingMLPackage wordMLPackage;

	// CONSTANTS

	private final static String VALUE_KEY = "value";
	private final static String TYPE_KEY = "type";
	private final static String TEXT_KEY = "text";
	private final static String TABLE_KEY = "table";
	private final static String LIST_KEY = "list";
	private final static String IMAGE_KEY = "image";
	private final static String ROW_KEY = "row";
	private final static String COLUMN_KEY = "columns";
	private final static String PLACEHOLDER_KEY = "placeholder";

	/**
	 * Find and Replace contents of word file in memory.
	 * Will not destroy current word document.
	 * @param wordFile - A String of the path to your word file.
	 * @param jsonFile - A String of the path to your json file.
	 * @throws Exception thrown for various reasons including parsing issues, file not found, could not open.
	 * JSON errors etc. Check the messages for more information.
	 */
	public Templatr(String wordFile, String jsonFile) throws Exception {
		this.wordFile = wordFile;
		this.jsonFile = jsonFile;

		init();
		replaceElements();
	}

	/**
	 * Initializes all required resources for the JSON and Docx4j
	 * @throws IOException
	 * @throws ParseException
	 * @throws Docx4JException
	 */
	private void init() throws IOException, ParseException, Docx4JException {

		JSONParser parser = new JSONParser();

		JSONObject result = (JSONObject) parser.parse(new FileReader(new File(
				jsonFile)));
		items = (JSONArray) result.get("items");

		wordMLPackage = WordprocessingMLPackage.load(new File(wordFile));

		documentPart = wordMLPackage.getMainDocumentPart();

	}

	/**
	 * Replaces all placeholder text within the document with the specified content in the JSON 
	 * @throws Exception
	 */
	private void replaceElements() throws Exception {

		for (Object o : items) {
			if (o instanceof JSONObject) {
				JSONObject obj = (JSONObject) o;
				int index = this.findIndexOfText((String) obj.get(PLACEHOLDER_KEY));
				int previousIndex = -1;
				// in a case where it could not replace the text, continue to avoid an infinte loop
				while (index >= 0 && index != previousIndex)  {

					if (index < 0) { break; }
					String type = (String) obj.get(TYPE_KEY);
					if (type == null) {
						continue;
					}
					if (type.equals(TEXT_KEY)) {

						this.replaceTextInParagraph(
								(String) obj.get(PLACEHOLDER_KEY),
								(String) obj.get(VALUE_KEY));
					} else if (type.equals(LIST_KEY)) {
						insertList(obj);
					} else if (type.equals(TABLE_KEY)) {
						Tbl table = this.createTable(obj);

						insertObject(index, table);
						documentPart.getContent().remove(index + 1); // remove the old paragraph with the placeholder
					} else if (type.equals(IMAGE_KEY)) {
						P img = this.createImage((String) obj.get(VALUE_KEY));
						insertObject(index, img);
						documentPart.getContent().remove(index + 1);
					}
					previousIndex = index;
					index = this.findIndexOfText((String) obj.get(PLACEHOLDER_KEY));
				}
			}
		}
	}

	/**
	 * Determines where the object is to be inserted given an index
	 * @param index the location where it should go
	 * @param obj the object to insert. 
	 */
	private void insertObject(int index, Object obj) {
		if (index < 0) { return; }
		if (index < this.documentPart.getContent().size()) {
			documentPart.getContent().add(index, obj);
		} else {
			documentPart.getContent().add(obj);
		}
	}
	
	/**
	 * Inserts a list given the list object from the JSON file
	 * @param list list object
	 * @throws Exception
	 */
	private void insertList(JSONObject list) throws Exception {
		
		int index = this.findIndexOfText((String)list.get(PLACEHOLDER_KEY));
		int newIndex = index;
		this.replaceTextInParagraph((String)list.get(PLACEHOLDER_KEY), "");
		JSONArray elements = (JSONArray)list.get(VALUE_KEY);
		
		for (Object obj: elements) {
			// cast to JSON Object
			if (! (obj instanceof JSONObject)) { 
				throw new Exception("Expected JSONObject inside array, got: " + obj.getClass());
			}
			JSONObject o = (JSONObject)obj;
			String type = (String)o.get(TYPE_KEY);
			
			if (type.equals(TEXT_KEY)) {
				P par = createParagraph((String)o.get(VALUE_KEY));
				// add to the the existing paragraph instead of creating a new one
				this.insertObject(newIndex, par);
				
			} else if(type.equals(TABLE_KEY)) {
				Tbl table = createTable(o);
				this.insertObject(newIndex, createParagraph(""));
				this.insertObject(++newIndex, table);
				
			} else if (type.equals(IMAGE_KEY)) {
				P par = createImage((String)o.get(VALUE_KEY));
				this.insertObject(newIndex, par);
			}
			
			if (index == newIndex) {
				documentPart.getContent().remove(index+1);
			}
			
			newIndex++;
			
		}	
		
	}
	
	/**
	 * Creates an image object inside of a paragraph object given  a full file path
	 * @param filename The full path to the desired image. 
	 * @return paragraph object with the image insde of a run
	 * @throws Exception
	 */
	public P createImage(String filename) throws Exception {
		ObjectFactory factory = new ObjectFactory();
		byte[] imgBytes;
		File f = new File(filename);
		String filenameHint = null, altText = null;
		int id1 = 0, id2 = 1;
		P par = factory.createP();
		R run = factory.createR();
		
		InputStream is = new FileInputStream(f);
		int fileLength = (int)f.length();
		imgBytes = new byte[fileLength];
		int offset = 0;
		int numRead = 0;
		// Read the image in
		while (offset < imgBytes.length && (numRead=is.read(imgBytes, offset, imgBytes.length-offset)) >= 0) {
			offset += numRead;
		}
		
		is.close();

		par.getContent().add(run);
		BinaryPartAbstractImage img = BinaryPartAbstractImage.createImagePart(wordMLPackage, imgBytes);
		
		Inline inline = img.createImageInline(filenameHint, altText, id1, id2, false);
		Drawing drawing = factory.createDrawing();
		run.getContent().add(drawing);
		drawing.getAnchorOrInline().add(inline);
		
		
		return par;
	}
	
	/**
	 * Given an existing paragraph it inserts a run
	 * @param par the paragraph in which to insert
	 * @param text The String to insert into the run.
	 */
	private void insertRun(P par, String text) {
		
		ObjectFactory factory = new ObjectFactory();
		R run = factory.createR();
		Text txt = factory.createText();
		txt.setValue(text);
		run.getContent().add(txt);
		
		par.getContent().add(run);
		
	}
	
	/**
	 * Creates a pargraph in the standard document font.
	 * @param data the string to add into a paragraph.
	 * @return The paragraph object with the text inserted.
	 */
	private P createParagraph(String data) {
		
		
		ObjectFactory factory = new ObjectFactory();
		P par = factory.createP();
		insertRun(par, data);
		
		return par;
		
	}
	
	/**
	 * Creates a Table object given a JSON object in the correct format.
	 * @param tableObj The JSON object in which the table is represented
	 * @return a standard word table
	 */
	private Tbl createTable(JSONObject tableObj) {
		ObjectFactory factory = new ObjectFactory();
		Tbl table = factory.createTbl();
		JSONArray items = (JSONArray)tableObj.get(VALUE_KEY);
		JSONObject head = (JSONObject)((JSONObject)items.get(0)).get(COLUMN_KEY);
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
		
		// remove the columns object to loop through the body of the table.
		items.remove(0);
		
		for (Object obj: items) {
			JSONObject o = (JSONObject)obj;
			row = factory.createTr();
			for (Object key: sortedKeySet) {
				Tc cell = factory.createTc();
				Object value = head.get(key);
				String s = (String)((JSONObject)o.get(ROW_KEY)).get(value);
				Text t = factory.createText();
				t.setValue(s);
				cell.getContent().add(documentPart.createParagraphOfText(s));
				row.getContent().add(cell);
			}
			table.getContent().add(row);
		}
		
		return table;
		
	}
	
	/**
	 * Given a String in the document it returns the index of the paragraph in which it resides.
	 * @param toFind The String to search for (needle)
	 * @return A positive index (or zero) if found, negative if not found. 
	 */
	private int findIndexOfText(String toFind) {
		int index = -1;

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

	/**
	 * Locates text in a paragraph within the document and replaces the text in the appropriate run
	 * @param toFind String you are searching for (needle)
	 * @param toReplace String to replace toFind. 
	 */
	private void replaceTextInParagraph(String toFind, String toReplace) {

		int index = findIndexOfText(toFind);
		if (index < 0) { return; }
		P par = (P) documentPart.getContent().get(index);
		for (Object o : par.getContent()) {

			if (o instanceof R) {
				R run = (R) o;

				for (Object inner : run.getContent()) {
					if (! (inner instanceof JAXBElement)) { continue; }
					if (((JAXBElement) inner).getValue() instanceof Text) {


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
	}
	
	/**
	 * Given an index of a paragraph object it returns the object
	 * @param index of a paragraph object. 
	 * @return
	 */
	private P getParagraph(int index) {
		return (P)documentPart.getContent().get(index);
	}

	/**
	 * Saves the document with replaced values to the given file path.
	 * @param fileName the full path to where you want to store your file. 
	 * @throws Docx4JException
	 */
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
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
