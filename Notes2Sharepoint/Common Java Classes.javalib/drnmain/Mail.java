package drnmain;

import lotus.domino.Document;
import lotus.domino.NotesException;

public class Mail {
	public static void mailSafeSend(Document doc, java.lang.String sendTo){
		try{
			doc.send(sendTo);			
		}catch(NotesException e){
			e.printStackTrace();
		}		
	}
	
	
	public static void mailSafeSend(Document doc, java.util.Vector sendTo){
		try{
			doc.send(sendTo);			
		}catch(NotesException e){
			e.printStackTrace();
		}		
	}
	
}
