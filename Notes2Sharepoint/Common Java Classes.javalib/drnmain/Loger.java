package drnmain;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.Vector;
import lotus.domino.Database;
import lotus.domino.Document;
import lotus.domino.Log;
import lotus.domino.NotesException;
import lotus.domino.Session;
import lotus.domino.RichTextItem;
import drnmain.Config;

public class Loger {
	private Log notesLog;
	private boolean testMode;
	private Session session;
	private boolean isValid;
	private Config config;
	private Vector sendTo;
	private Vector copyTo;
	
	public Loger(){
		testMode = false;
	}
	
	public Loger(boolean inTestMode){
		testMode = inTestMode;
	}
			
	public Loger(Session inSession){
			session = inSession;			
			config = new Config(inSession);
			isValid = true;
	}
	
	public Loger(Session inSession, String logTitle, boolean inTestMode) {
		try {
			session = inSession;
			notesLog = session.createLog(logTitle);
			testMode = inTestMode;
			config = new Config(inSession);
			isValid = true;
		} catch (NotesException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			isValid = false;
		}
		
		// TODO Auto-generated constructor stub
	}
	
	public void s(String aString) {
		if (testMode) {
			System.out.println(aString);
		}

	}
	
	public void setTestMode(boolean testMode) {
		this.testMode = testMode;
	}
	
	
	public void sendError(Document contextDoc, Exception e, String title){
		StringWriter errors = new StringWriter();
		Document emailDoc; 
		Database currentDB;
		RichTextItem body; 
		System.out.println("Send error");
		
		if (sendTo == null){
			try {
				sendTo = this.config.getKeywords("Error_SendTo");
				copyTo = this.config.getKeywords("Error_BCC"); 
			
			} catch (NotesException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}		
		e.printStackTrace(new PrintWriter(errors));

		if (sendTo != null && !sendTo.isEmpty()) {
			try {
				System.out.println("Send error body");
				currentDB = contextDoc.getParentDatabase();
				emailDoc = currentDB.createDocument();
				emailDoc.replaceItemValue("Form", "Memo");
				emailDoc.replaceItemValue("Subject", title);
				body = emailDoc.createRichTextItem("Body");
				body.addNewLine(2);
				body.appendText("User: " + session.getCommonUserName());
				body.addNewLine(2);
				body.appendText("Server: " + currentDB.getServer());
				body.addNewLine(2);
				body.appendText("Database: " + currentDB.getTitle());	
				body.appendDocLink(currentDB);
				body.addNewLine(2);
				if (contextDoc != null){			
					body.appendText("Document: " + contextDoc.getItemValueString("Form"));	
					body.appendDocLink(contextDoc);
					body.addNewLine(2);							
				}
				body.appendText(errors.toString());		
				emailDoc.send(sendTo);
				if (copyTo != null && !copyTo.isEmpty()) {
					emailDoc.send(copyTo);
				}
			} catch (NotesException e1) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				
				e1.printStackTrace();
			}			
		}		
	}	
}
