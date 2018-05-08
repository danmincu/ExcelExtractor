import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Vector;

import org.apache.http.HttpStatus;
import org.json.JSONArray;
import org.json.JSONObject;

import com.mashape.unirest.http.HttpResponse;
import com.mashape.unirest.http.JsonNode;
import com.mashape.unirest.http.Unirest;
import com.mashape.unirest.http.exceptions.UnirestException;

/**
 * 
 * @author danmincu@gmail.com
 * 
 */
public class SharepointProxy
{

    private String token;

    private String digestToken;

    private boolean isLoggedIn = false;

    private boolean isDigestTokenObtained = false;

    private String site;

    private String clientSecret;

    private String clientId;

    /**
     * Constructor
     * 
     * @param site
     *            represents the Sharepoint sub-site e.g. "CATSII"
     * @param clientId
     *            the clientId obtained from the Sharepoint Security Application
     * @param clientSecret
     *            the clientSecret obtained from the Sharepoint Security
     *            Application
     */
    public SharepointProxy(String site, String clientId, String clientSecret)
    {
        this.site = site;
        this.clientId = clientId;
        this.clientSecret = clientSecret;
    }

    private void login(boolean getDigestToken) throws UnirestException
    {
        HttpResponse<JsonNode> response = Unirest
                .post(
                        "https://accounts.accesscontrol.windows.net/5b8221b8-e696-4d6e-870f-20382113efad/tokens/OAuth/2")
                .header("content-type", "application/x-www-form-urlencoded")
                .header("cache-control", "no-cache")
                .body(
                        "grant_type=client_credentials&client_id="
                                + this.clientId
                                + "%405b8221b8-e696-4d6e-870f-20382113efad&client_secret="
                                + this.clientSecret
                                + "&resource=00000003-0000-0ff1-ce00-000000000000%2Fdiscoveryair.sharepoint.com%405b8221b8-e696-4d6e-870f-20382113efad")
                .asJson();

        JSONObject tokenObj = response.getBody().getObject();

        this.token = tokenObj.get("access_token").toString();

        this.isLoggedIn = true;

        if (getDigestToken)
            this.getDigestToken();

    }

    private void checkLoggedIn() throws UnirestException
    {
        if (!this.isLoggedIn)
        {
            this.login(true);
        }
    }

    private void checkDigestToken() throws UnirestException
    {
        if (!this.isDigestTokenObtained)
        {
            this.getDigestToken();
        }
    }

    private void getDigestToken() throws UnirestException
    {

        this.checkLoggedIn();

        HttpResponse<JsonNode> digestResponse = Unirest.post(
                "https://discoveryair.sharepoint.com/" + this.site
                        + "/_api/contextinfo").header("accept",
                "application/json; odata=verbose").header("authorization",
                "bearer " + token).header("cache-control", "no-cache").asJson();

        this.digestToken = digestResponse.getBody().getObject().getJSONObject(
                "d").getJSONObject("GetContextWebInformation").get(
                "FormDigestValue").toString();
        isDigestTokenObtained = true;
    }

    /**
     * Downloads the content of a file in the string format. Suitable for CSV or
     * simple text files
     * 
     * @param folder
     *            the subfolder where the file lives. e.g. "Shared Documents"
     * @param fileName
     *            the name of the file e.g. "test.csv"
     * @return the text content of the file
     * @throws UnirestException
     */
    public String downloadTextFile(String folder, String fileName)
            throws UnirestException
    {
        this.checkLoggedIn();

        String relativeFilePath = folder + "/" + fileName;
        HttpResponse<String> fileResponse = Unirest
                .get(
                        "https://discoveryair.sharepoint.com/"
                                + this.site
                                + "/_api/Web/GetFileByServerRelativePath%28decodedurl=%27/"
                                + this.site + "/" + relativeFilePath
                                + "%27%29/$value").header("accept",
                        "application/json; odata=verbose").header(
                        "authorization", "bearer " + this.token).header(
                        "cache-control", "no-cache").asString();

        return fileResponse.getBody().toString();
    }

    /**
     * Uploads a string as a content of a file on the Sharepoint server
     * 
     * @param folder
     *            the subfolder where the file lives. e.g. "Shared Documents"
     * @param fileName
     *            the name of the file e.g. "test.csv"
     * @param content
     *            the text content to be uploaded
     * @throws UnirestException
     */
    public void uploadTextFile(String folder, String fileName, String content)
            throws UnirestException
    {
        this.checkLoggedIn();
        this.checkDigestToken();

        HttpResponse<String> uploadTextResponse = Unirest.post(
                "https://discoveryair.sharepoint.com/" + this.site
                        + "/_api/web/GetFolderByServerRelativeUrl%28%27/"
                        + this.site + "/" + folder
                        + "%27%29/Files/add%28url=%27" + fileName
                        + "%27,overwrite=true%29").header("accept",
                "application/json; odata=verbose").header("authorization",
                "bearer " + this.token).header("x-requestdigest",
                this.digestToken).header("cache-control", "no-cache").body(
                content).asString();
    }

    /**
     * uploads a file to the server. The fileName is preserved on the Sharepoint
     * server
     * 
     * @param localFilePath
     *            the complete path of the file to upload
     * @param folder
     *            the subfolder on the server where the file should be uploaded.
     *            e.g. "Shared Documents"
     * @throws UnirestException
     * @throws IOException
     */
    public void uploadBinaryFile(String localFilePath, String folder)
            throws UnirestException, IOException
    {
        this.checkLoggedIn();
        this.checkDigestToken();

        File file = new File(localFilePath);
        final InputStream stream = new FileInputStream(file);
        final byte[] bytes = new byte[stream.available()];
        stream.read(bytes);
        stream.close();

        HttpResponse<String> uploadBinaryResponse = Unirest.post(
                "https://discoveryair.sharepoint.com/" + this.site
                        + "/_api/web/GetFolderByServerRelativeUrl%28%27/"
                        + this.site + "/" + folder
                        + "%27%29/Files/add%28url=%27" + file.getName()
                        + "%27,overwrite=true%29").header("accept",
                "application/json; odata=verbose").header("authorization",
                "bearer " + this.token).header("x-requestdigest",
                this.digestToken).header("cache-control", "no-cache").body(
                bytes).asString();
    }

    /**
     * 
     * @param localFolderPath the path of the folder where the downloaded content is stored.
     *            The file name is being preserved
     * @param folder the subfolder on the server where the file to be downloaded
     *            resides. e.g. "Shared Documents"
     * @param fileName the name of the file on the server
     * @throws UnirestException
     * @throws IOException
     */
    public void downloadBinaryFile(String localFolderPath, String folder,
            String fileName) throws UnirestException, IOException
    {
        this.checkLoggedIn();
        this.checkDigestToken();

        String relativeFilePath = folder + "/" + fileName;
        HttpResponse<InputStream> fileResponse = Unirest
                .get(
                        "https://discoveryair.sharepoint.com/"
                                + this.site
                                + "/_api/Web/GetFileByServerRelativePath%28decodedurl=%27/"
                                + this.site + "/" + relativeFilePath
                                + "%27%29/$value").header("accept",
                        "application/json; odata=verbose").header(
                        "authorization", "bearer " + this.token).header(
                        "x-requestdigest", this.digestToken).header(
                        "cache-control", "no-cache").asBinary();

        byte[] buffer = new byte[fileResponse.getBody().available()];
        fileResponse.getBody().read(buffer);

        File targetFile = new File(localFolderPath + "\\" + fileName);
        OutputStream outStream = new FileOutputStream(targetFile);
        outStream.write(buffer);
        outStream.close();
    }

    /**
     * This is returning a list of file names from the server. Similar to the
     * "dir" command
     * 
     * @param folder the folder on the server to perform the action. e.g.
     *            "Shared Documents"
     * @return - a vector containing the names of the files residing in the
     *         queried folder.
     * @throws UnirestException
     */
    public Vector<String> getFolderList(String folder) throws UnirestException
    {
        this.checkLoggedIn();

        HttpResponse<JsonNode> responseListFiles = Unirest.get(
                "https://discoveryair.sharepoint.com/" + this.site
                        + "/_api/web/GetFolderByServerRelativeUrl%28%27/"
                        + this.site + "/" + folder + "%27%29/Files").header(
                "accept", "application/json; odata=verbose").header(
                "authorization", "bearer " + this.token).header(
                "cache-control", "no-cache").asJson();

        JSONArray filesArray = responseListFiles.getBody().getObject()
                .getJSONObject("d").getJSONArray("results");
        Vector<String> result = new Vector<String>();
        for (int i = 0; i < filesArray.length(); i++)
        {
            result.add(filesArray.getJSONObject(i).get("Name").toString());
        }
        return result;
    }
    
    /**
     * This command deletes a file from the server
     * @param folder the folder on the server to perform the action. e.g. "Shared Documents"
     * @param fileName the name of the file to be deleted
     * @throws UnirestException
     * @return True if the file was deleted. If the file is not existent it returns false.
     */
    public boolean deleteFile(String folder, String fileName) throws UnirestException
    {
        this.checkLoggedIn();
        
        HttpResponse<String> response = Unirest.post("https://discoveryair.sharepoint.com/" + this.site + 
                "/_api/Web/GetFileByServerRelativePath%28decodedurl=%27/"+this.site + "/" 
                + folder + "/" + fileName + "%27%29")
        .header("accept", "application/json; odata=verbose")
        .header("content-type", "application/json;odata=verbose")
        .header("authorization", "bearer " + this.token)
        .header("x-requestdigest", this.digestToken)
        .header("x-http-method", "DELETE")
        .header("if-match", "*")
        .header("cache-control", "no-cache")        
        .asString();
        return (response.getStatus() == HttpStatus.SC_OK);
    }
    
}