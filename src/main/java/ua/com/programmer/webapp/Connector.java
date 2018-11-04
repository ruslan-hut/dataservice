package ua.com.programmer.webapp;

import org.jawin.COMException;
import org.jawin.DispatchPtr;
import org.jawin.win32.Ole32;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

import java.io.FileReader;
import java.io.FileWriter;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Scanner;

class Connector {

    private DispatchPtr app;

    private String convert(String message){
        String result;
        try {
            result = new String(message.getBytes(StandardCharsets.ISO_8859_1), "Windows-1251");
        }catch (Exception stringEx){
            result = message;
        }
        return result;
    }

    void errorLog(String message){
        final String errorsFile = "connector.errors";
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        message = simpleDateFormat.format(calendar.getTime())+" "+message+"\n";
        try (FileWriter fileWriter = new FileWriter(errorsFile, true)) {
            fileWriter.write(convert(message));
            fileWriter.flush();
        }catch (Exception ex){
            System.out.println(ex.toString());
        }
    }

    boolean isBusy(){
        return app != null;
    }

    boolean openConnection(){
        if (isBusy()){
            return true;
        }
        final String settingsFile = "connector.settings";
        String init = "";
        String line;
        Scanner scanner;
        try(FileReader fileReader = new FileReader(settingsFile)) {
            scanner = new Scanner(fileReader);
            while (scanner.hasNextLine()){
                line = scanner.nextLine();
                if (line.contains("path=")){
                    init = line.replace("path=", "");
                    break;
                }
            }
        }catch (Exception fileReaderException){
            errorLog("FileReader: "+fileReaderException.toString());
            return false;
        }
        if (init.equals("")){
            errorLog("openConnection: No 'path' parameter in settings file");
            return false;
        }
        try {
            Ole32.CoInitialize();
            app = new DispatchPtr("V77.Application");
            app.invoke("Initialize",app.get("RMTrade"),init,"NO_SPLASH_SHOW");
            return true;
        }catch (Exception ex){
            errorLog("openConnection: "+ex.toString());
        }
        return false;
    }

    private DispatchPtr invokeProcedure(String commandName, String par1, String par2, String par3){
        DispatchPtr result;
        try {
            result = (DispatchPtr) app.invoke("MobileService", commandName, par1, par2, par3);
        }catch (COMException ex){
            errorLog("MobileService: "+commandName+"("+par1+","+par2+","+par3+") "+ex.toString());
            result = null;
        }
        return result;
    }

    JSONArray getReferenceData(String refType, String groupCode){
        JSONArray resultArray = new JSONArray();
        if (openConnection()) {
            DispatchPtr refItems = invokeProcedure("getReferenceData", refType, groupCode, "");
            if (refItems != null) {
                try {
                    Double listSize = (Double) refItems.invoke("GetListSize");
                    int listItemNumber = 1;
                    if ( listSize > 0.0){
                        while (listItemNumber <= listSize) {
                            readGoodsItemData((DispatchPtr) refItems.invoke("GetValue", listItemNumber), resultArray);
                            listItemNumber++;
                        }
                    }
                } catch (Exception ex) {
                    errorLog("getReferenceData: " +refType+"; "+groupCode+"; "+ex.toString());
                }
            }else{
                errorLog("getReferenceData: no data received from app");
            }
        }
        return resultArray;
    }

    JSONArray getDocumentsData(String docType, String period){
        JSONArray resultArray = new JSONArray();
        if (openConnection()){
            DispatchPtr refItems = invokeProcedure("getDocumentsData", docType, period, "");
            if (refItems != null) {
                try {
                    Double listSize = (Double) refItems.invoke("GetListSize");
                    int listItemNumber = 1;
                    if ( listSize > 0.0){
                        while (listItemNumber <= listSize) {
                            readDocumentItemData((DispatchPtr) refItems.invoke("GetValue", listItemNumber), resultArray);
                            listItemNumber++;
                        }
                    }
                }catch (Exception ex) {
                    errorLog("getDocumentsData: "+docType+"; "+ex.toString());
                }
            }else{
                errorLog("getDocumentsData: no data received from app");
            }
        }
        return resultArray;
    }

    String getDocumentContent(String docType, String number, String guid){
        String result="";
        if (openConnection()){
            try {
                result = (String) app.invoke("MobileService", "getDocumentContent", docType, number, guid);
            }catch (Exception ex) {
                errorLog("getDocumentContent: "+docType+"; "+number+"; "+ex.toString());
            }
        }
        return result;
    }

    JSONArray findBarcode(String barcode){
        JSONArray resultArray = new JSONArray();
        if (openConnection()){
            try {
                DispatchPtr refItems = invokeProcedure("findBarcode", barcode, "", "");
                Double listSize = (Double) refItems.invoke("GetListSize");
                int listItemNumber = 1;
                if ( listSize > 0.0){
                    while (listItemNumber <= listSize) {
                        readGoodsItemData((DispatchPtr) refItems.invoke("GetValue", listItemNumber), resultArray);
                        listItemNumber++;
                    }
                }
            }catch (Exception ex) {
                errorLog("findBarcode: "+barcode+"; "+ex.toString());
            }
        }
        return resultArray;
    }

    JSONArray getCatalogDataWithFilter(String type, String filter){
        JSONArray resultArray = new JSONArray();
        if (openConnection()){
            try {
                DispatchPtr refItems = invokeProcedure("getCatalogDataWithFilter", type, filter, "");
                Double listSize = (Double) refItems.invoke("GetListSize");
                int listItemNumber = 1;
                if ( listSize > 0.0){
                    while (listItemNumber <= listSize) {
                        readGoodsItemData((DispatchPtr) refItems.invoke("GetValue", listItemNumber), resultArray);
                        listItemNumber++;
                    }
                }
            }catch (Exception ex) {
                errorLog("getCatalogDataWithFilter: "+type+"; "+filter+"; "+ex.toString());
            }
        }
        return resultArray;
    }

    JSONArray getOptionsData(String userID){
        JSONArray resultArray = new JSONArray();
        if (openConnection()){
            try {
                DispatchPtr refItems = invokeProcedure("readOptions", userID, "", "");
                Double listSize = (Double) refItems.invoke("GetListSize");
                if (listSize > 0.0){
                    readOptionsData((DispatchPtr) refItems.invoke("GetValue", 1), resultArray);
                }
            }catch (Exception ex) {
                errorLog("getOptionsData: " + ex.toString());
            }
        }
        return resultArray;
    }

    void closeConnection(){
        app = null;
        try {
            Ole32.CoUninitialize();
        }catch (Exception ex){
            errorLog("closeConnection: "+ex.toString());
        }
    }

    private void readGoodsItemData(DispatchPtr item, JSONArray array) throws COMException {
        int isGroup;
        if ((Double) item.invoke("Get","isGroup") != 0.0) {
            isGroup = 1;
        }else {
            isGroup = 0;
        }
        JSONObject jsonItem = new JSONObject();
        jsonItem.put("id", item.invoke("Get","id").toString());
        jsonItem.put("isGroup", isGroup);
        jsonItem.put("code", item.invoke("Get","code").toString());
        jsonItem.put("description", item.invoke("Get","description").toString());
        jsonItem.put("art", item.invoke("Get","art").toString());
        jsonItem.put("unit", item.invoke("Get","unit").toString());
        jsonItem.put("groupName", item.invoke("Get","groupName").toString());
        jsonItem.put("groupCode", item.invoke("Get","groupCode").toString());
        array.add(jsonItem);
    }

    private void readDocumentItemData(DispatchPtr item, JSONArray array) throws COMException {
        int isProcessed = 0;
        if ((Double) item.invoke("Get","isProcessed") != 0.0) isProcessed = 1;
        int isDeleted = 0;
        if ((Double) item.invoke("Get","isDeleted") != 0.0) isDeleted = 1;

        JSONObject jsonItem = new JSONObject();
        jsonItem.put("guid", item.invoke("Get", "guid").toString());
        jsonItem.put("isProcessed", isProcessed);
        jsonItem.put("isDeleted", isDeleted);
        jsonItem.put("number", item.invoke("Get", "number").toString());
        jsonItem.put("date", item.invoke("Get", "date").toString());
        jsonItem.put("contractor", item.invoke("Get", "contractor").toString());
        jsonItem.put("company", item.invoke("Get", "company").toString());
        jsonItem.put("warehouse", item.invoke("Get", "warehouse").toString());
        jsonItem.put("sum", item.invoke("Get", "sum").toString());

        array.add(jsonItem);
    }

    private void readOptionsData(DispatchPtr item, JSONArray array) throws COMException {
        JSONObject jsonItem = new JSONObject();
        boolean read = item.invoke("Get", "read").toString().equals("true");
        boolean write = item.invoke("Get", "write").toString().equals("true");
        jsonItem.put("read", read);
        jsonItem.put("write", write);
        jsonItem.put("user", item.invoke("Get", "user").toString());
        jsonItem.put("catalog", item.invoke("Get", "catalog").toString());
        jsonItem.put("document", item.invoke("Get", "document").toString());

        array.add(jsonItem);
    }

    JSONArray saveDocument(JSONObject document){
        JSONArray resultArray = new JSONArray();
        if (openConnection()){
            try {

                DispatchPtr values = (DispatchPtr) app.invoke("CreateObject", "ValueList");
                values.invoke("Set", "number", document.get("number"));
                values.invoke("Set", "guid", document.get("guid"));
                values.invoke("Set", "type", document.get("type"));

                DispatchPtr lines = (DispatchPtr) app.invoke("CreateObject", "ValueList");
                JSONArray jsonLines = (JSONArray) document.get("lines");
                for (Object objectLine: jsonLines){
                    JSONObject jsonLine = (JSONObject) objectLine;
                    DispatchPtr line = (DispatchPtr) app.invoke("CreateObject", "ValueList");
                    line.invoke("Set", "art", jsonLine.get("art"));
                    line.invoke("Set", "code", jsonLine.get("code"));
                    line.invoke("Set", "quantity", jsonLine.get("quantity"));
                    line.invoke("Set", "price", jsonLine.get("price"));
                    lines.invoke("AddValue", line);
                }
                values.invoke("Set", "lines", lines);

                DispatchPtr result77 = (DispatchPtr) app.invoke("MobileService", "saveDocument", values);

                readSaveDocumentResult(result77, resultArray);

            }catch (Exception ex){
                errorLog("saveDocument: "+ex.toString());
            }
        }
        return resultArray;
    }

    private void readSaveDocumentResult(DispatchPtr item, JSONArray array) throws COMException {
        JSONObject jsonItem = new JSONObject();
        jsonItem.put("saved", item.invoke("Get", "saved").toString());
        jsonItem.put("error", item.invoke("Get", "error").toString());
        array.add(jsonItem);
    }
}
