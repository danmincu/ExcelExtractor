package drnmain;

import java.util.Vector;

import lotus.domino.NotesException;
import lotus.domino.Session;
import lotus.domino.Database;
import lotus.domino.View;
import lotus.domino.Document;

public class Config {
	private Session session;
	private Database currentDb;
	private View dbConfigView; 
	private View keywordsView; 
	
	
	public Config (Session inSession){
		try {
			session = inSession;
			currentDb = session.getCurrentDatabase();
			dbConfigView = currentDb.getView("LUdbProfiles");			
			keywordsView = currentDb.getView("LUKeywordsID");
		} catch (NotesException e) {			
			e.printStackTrace();
		}
	}
	
	public Database getDatabase(String key) throws NotesException{
		Database tmpDB = null;
		Document dbProfileDoc;
		String server;
		String dbName;
		String type;
		
		if (dbConfigView != null){
			dbProfileDoc = dbConfigView.getDocumentByKey(key);
			
			server = dbProfileDoc.getItemValueString("db_server");
			dbName = dbProfileDoc.getItemValueString("db_fileName");
			type = dbProfileDoc.getItemValueString("LocalReplica");
			
			//System.out.println(server);
			//System.out.println(dbName);
			//System.out.println(type);
			
			if(type.equalsIgnoreCase("0")){
				tmpDB = session.getDatabase(server, dbName);
				
				//Use local Replica only|2
				//Use server Replica only|0
				//Use local Replica if server Replica cannot be opened | 1
				//Use server Replica if local Replica cannot be opened | 3
			}
			if(type.equalsIgnoreCase("2")){
				tmpDB = session.getDatabase(null, dbName);			
			}
			if(type.equalsIgnoreCase("1")){
				tmpDB = session.getDatabase(server, dbName);
				if (tmpDB == null) {				
					tmpDB = session.getDatabase(null, dbName);		
				}else if (!tmpDB.isOpen()){
					tmpDB = session.getDatabase(null, dbName);
				}
			}			
			if(type.equalsIgnoreCase("3")){
				tmpDB = session.getDatabase(null, dbName);
				if (tmpDB == null) {				
					tmpDB = session.getDatabase(server, dbName);					
				}else if (!tmpDB.isOpen()){
					tmpDB = session.getDatabase(server, dbName);
				}
			}
			if(type.equalsIgnoreCase("4")){
				tmpDB = session.getDatabase(this.session.getCurrentDatabase().getServer(), dbName);			
			}		
		}		
		return tmpDB;
	}
	
	public String getKeyword(String key) throws NotesException{
		Document dbKeywordDoc;
		
		dbKeywordDoc = keywordsView.getDocumentByKey(key);
		
		if (dbKeywordDoc == null){
			return "";
		}else{
			return dbKeywordDoc.getItemValueString("KeywordValues");	
		}		
	}
	public Vector getKeywords(String key) throws NotesException{
		Document dbKeywordDoc;
		
		dbKeywordDoc = keywordsView.getDocumentByKey(key);
		
		if (dbKeywordDoc == null){
			return null;
		}else{
			return dbKeywordDoc.getItemValue("KeywordValues");	
		}		
	}
}
