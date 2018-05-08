import lotus.domino.AgentBase;
import lotus.domino.Session;
import drnmain.Config;

public class JavaAgent extends AgentBase
{

	public void NotesMain()
	{

		try
		{
			Session session = getSession();
			// AgentContext agentContext = session.getAgentContext();

			Config config = new Config(session);
			// SharepointConnector con = new SharepointConnector();
			ExcelParser parser = new ExcelParser(config);

			parser.setSchdFileName("SourceSchedulerFileName");
			parser.setSchdFolderName("SourceSchedulerFolder");
			parser.setSchdPublishedFileName("SharepointPublishedSchedulerFileName");
			parser.setSchdPublishedFoler("SharepointPublishFolder");
			parser.setExtractorFileName("ExcelExtractorFileName");
			parser.setSheetDetails("SheetDetailsForExtraction");
			parser.setNmbOfDays("NumberOfDaysToExtract");
			parser.setLocalWorkFolder("LocalWorkFolder");

			parser.InitProxyCon("SharepointSite", "SharepointClientId",
					"SharepointClientSecret");

			// Phase 1 - download the file first
			parser.download();

			// Phase 2 - parse the Excel file
			String outputFile = parser.RunParser();

			// Phase 3 - upload the file to Sharepoint
			parser.getCon().setFileName(outputFile);
			parser.upload(outputFile);

		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
}