package ua.com.programmer.webapp;

import org.json.simple.JSONArray;
import org.json.simple.parser.JSONParser;
import org.json.simple.JSONObject;
import org.json.simple.parser.ParseException;

import javax.ws.rs.*;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import java.io.BufferedReader;
import java.io.InputStream;
import java.io.InputStreamReader;

@Path("/rc")
public class DataLoadService {

    private final Connector connector = new Connector();

    @GET
    @Path("/test")
    public Response testService(){
        String result="";
        try {
            connector.openConnection();
            connector.closeConnection();
            result = "Status: OK; Database connected";
        }catch (Exception ex){
            result = "Status: Error; "+ex.toString();
        }
        return Response.status(200).entity(result).build();
    }

    @POST
    @Path("/createOrder")
    @Consumes(MediaType.APPLICATION_JSON)
    public Response createOrder(InputStream inputStream){

        JSONObject outputDocument = new JSONObject();
        outputDocument.put("status", true);
        outputDocument.put("error", "");


        if (connector.isBusy()){
            connector.errorLog("busy state");
            return Response.status(200).entity(outputDocument.toString()).build();
        }

        if (!connector.openConnection()){
            connector.errorLog("quit without processing request");
            connector.closeConnection();
            return Response.status(200).entity(outputDocument.toString()).build();
        }

        StringBuilder builder = new StringBuilder();
        try {

            BufferedReader in = new BufferedReader(new InputStreamReader(inputStream));
            String line;
            while ((line = in.readLine()) != null) {
                builder.append(line);
            }

            connector.createOrder(builder.toString());

        }catch (Exception ex){
            //bad request
            connector.errorLog("parse request: "+ex.toString());
            outputDocument.put("status", false);
            outputDocument.put("error", "can't process the request");
        }

        connector.closeConnection();

        return Response.status(200).entity(outputDocument.toString()).build();
    }

    @POST
    @Path("/utf8")
    @Consumes(MediaType.TEXT_PLAIN)
    public Response convertToUTF8(InputStream inputStream){
        String responseData="";
        StringBuilder builder = new StringBuilder();
        try {

            BufferedReader in = new BufferedReader(new InputStreamReader(inputStream));
            String line;
            while ((line = in.readLine()) != null) {
                builder.append(line);
            }

            responseData = connector.convertToUTF8(builder.toString());

        }catch (Exception ex){
            connector.errorLog("convertToUTF8: "+ex.toString());
        }

        return Response.status(200).entity(responseData).build();
    }

    @POST
    @Path("/pst/{userID}")
    @Consumes(MediaType.APPLICATION_JSON)
    public Response postOperation(InputStream inputStream){

        JSONObject outputDocument = new JSONObject();
        outputDocument.put("result","ok");
        outputDocument.put("data", new JSONArray());
        outputDocument.put("message","");
        outputDocument.put("token","");

        if (connector.isBusy()){
            connector.errorLog("busy state");
            return Response.status(200).entity(outputDocument.toString()).build();
        }

        if (!connector.openConnection()){
            connector.errorLog("quit without processing request");
            connector.closeConnection();
            return Response.status(200).entity(outputDocument.toString()).build();
        }

        JSONArray responseData = new JSONArray();

        JSONObject inputDocument;
        StringBuilder builder = new StringBuilder();
        try {

            BufferedReader in = new BufferedReader(new InputStreamReader(inputStream));
            String line;
            while ((line = in.readLine()) != null) {
                builder.append(line);
            }

            JSONParser parser = new JSONParser();
            inputDocument = (JSONObject) parser.parse(builder.toString());

            responseData = processRequestData(inputDocument);

        }catch (Exception ex){
            //bad request
            connector.errorLog("parse request: "+ex.toString());
            outputDocument.put("result","error");
            outputDocument.put("message","Invalid request data");
        }

        connector.closeConnection();

        outputDocument.put("data",responseData);
        String output = outputDocument.toString();
        return Response.status(200).entity(output).build();
    }

    private JSONArray processRequestData(JSONObject inputDocument) throws ParseException{

        String userID = (String) inputDocument.get("userID");
        String type = (String) inputDocument.get("type");
        String data = (String) inputDocument.get("data");

        JSONArray responseData = new JSONArray();
        JSONParser parser = new JSONParser();
        JSONObject parameters;
        JSONObject result;
        String resultString;

        switch (type){
            case "check":
                responseData = connector.getOptionsData(userID);
                break;
            case "documents":
                responseData = connector.getDocumentsData(data,"");
                break;
            case "documentContent":
                parameters = (JSONObject) parser.parse(data);
                String docType = (String) parameters.get("type");
                String number = (String) parameters.get("number");
                String guid = (String) parameters.get("guid");

                resultString = connector.getDocumentContent(docType,number,guid);

                result = (JSONObject) parser.parse(resultString);
                responseData = (JSONArray) result.get("result");
                break;
            case "catalog":
                parameters = (JSONObject) parser.parse(data);
                String catalogType = (String) parameters.get("type");
                String groupCode = (String) parameters.get("group");
                String filter = (String) parameters.get("searchFilter");

                if (filter.equals("")) {
                    responseData = connector.getReferenceData(catalogType, groupCode);
                }else {
                    responseData = connector.getCatalogDataWithFilter(catalogType,filter);
                }
                break;
            case "barcode":
                responseData = connector.findBarcode(data);
                break;
            case "saveDocument":
                parameters = (JSONObject) parser.parse(data);
                responseData = connector.saveDocument(parameters);
                break;
        }
        return responseData;
    }
}