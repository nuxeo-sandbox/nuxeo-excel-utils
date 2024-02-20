package nuxeo.excel.utils.operations;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Context;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.Blobs;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.openxmlformats.schemas.officeDocument.x2006.customProperties.CTProperty;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

/**
 * Returns a JSON blob with the following objects:
 * <pre>
 * {
 *   "sheets": an array with the names of thsheets
 *   "coreProperties": An object with the core properties:
 *                     title, creator, category, ...
 *   "customProperties": An object with the core properties.
 *                       WARNING: Only String properties are returned with a value. Others are set to null.
 *                     
 * }
 * </pre>
 *
 */
@Operation(id = GetPropertiesOp.ID, category = Constants.CAT_BLOB, label = "Excel.GetProperties", description = "Returns a JSONBlob with the properties."
        + " If input is a document, xpath is used to get the blob (default file:content)")
public class GetPropertiesOp {

    public static final String ID = "Excel.GetProperties";

    public static final SimpleDateFormat dateFormatter = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ssZ");
    
    @Context
    protected CoreSession session;

    @Param(name = "xpath", required = false, values = { "file:content" })
    protected String xpath = "file:content";

    @OperationMethod
    public Blob run(DocumentModel doc) {
        Blob b = (Blob) doc.getPropertyValue(xpath);

        return getProperties(b);
    }

    @OperationMethod
    public Blob run(Blob blob) {

        return getProperties(blob);
    }

    protected Blob getProperties(Blob blob) {

        ObjectNode objNode = null;

        if (blob == null) {
            ObjectMapper mapper = new ObjectMapper();
            ObjectNode jsonObject = mapper.createObjectNode();
            jsonObject.put("status", "No input blob");
        } else {
            try (InputStream is = blob.getStream()) {
                Workbook workbook = WorkbookFactory.create(is);
                if (workbook instanceof XSSFWorkbook) {
                    objNode = readXLSXProperties((XSSFWorkbook) workbook);
                } else {
                    throw new IOException("The plugin supports only Excel 2007 OOXML (.xlsx) format.");
                }
            } catch (IOException e) {
                throw new NuxeoException(e);
            }
        }

        return Blobs.createJSONBlob(objNode.toString());
    }

    protected ObjectNode readXLSXProperties(XSSFWorkbook workbook) {

        ObjectMapper mapper = new ObjectMapper();
        
        // Sheets
        ArrayNode jsonSheets = mapper.createArrayNode();
        for(int i = 0; i < workbook.getNumberOfSheets(); i++) {
            workbook.getSheetName(i);
            jsonSheets.add(workbook.getSheetName(i));
        }

        POIXMLProperties props = workbook.getProperties();

        // Core properties
        ObjectNode jsonCoreProps = mapper.createObjectNode();

        POIXMLProperties.CoreProperties coreProps = props.getCoreProperties();
        jsonCoreProps.put("title", coreProps.getTitle());
        jsonCoreProps.put("creator", coreProps.getCreator());
        jsonCoreProps.put("category", coreProps.getCategory());
        jsonCoreProps.put("contentStatus", coreProps.getContentStatus());
        jsonCoreProps.put("contentType", coreProps.getContentType());
        jsonCoreProps.put("created", formatDate(coreProps.getCreated()));
        jsonCoreProps.put("description", coreProps.getDescription());
        jsonCoreProps.put("identifier", coreProps.getIdentifier());
        jsonCoreProps.put("keywords", coreProps.getKeywords());
        jsonCoreProps.put("lastModifiedByUser", coreProps.getLastModifiedByUser());
        jsonCoreProps.put("lastPrinted", formatDate(coreProps.getLastPrinted()));
        jsonCoreProps.put("modified", formatDate(coreProps.getModified()));
        jsonCoreProps.put("revision", coreProps.getRevision());
        jsonCoreProps.put("subject", coreProps.getSubject());
        jsonCoreProps.put("version", coreProps.getVersion());
        
        // Custom properties
        // Only Strings. Other types are hard to get, actually.
        ObjectNode jsonCustomProps = mapper.createObjectNode();
        POIXMLProperties.CustomProperties customProps = props.getCustomProperties();
        List<CTProperty> ctProperties = customProps.getUnderlyingProperties().getPropertyList();
        for (CTProperty ctProperty : ctProperties) {
            String name = ctProperty.getName();

            if(ctProperty.isSetLpstr()) {
                jsonCustomProps.put(name, ctProperty.getLpstr());
            } else if(ctProperty.isSetLpwstr()) {
                jsonCustomProps.put(name, ctProperty.getLpwstr());
            } else if(ctProperty.isSetDate()) {
                Calendar cal = ctProperty.getDate();
                if(cal == null) {
                    jsonCustomProps.put(name, "");
                } else {
                    jsonCustomProps.put(name, formatDate(cal.getTime()));
                }
            } else if(ctProperty.isSetDecimal()) {
                jsonCustomProps.put(name, ctProperty.getDecimal());
            } else if(ctProperty.isSetInt()) {
                jsonCustomProps.put(name, ctProperty.getInt());
            }  else if(ctProperty.isSetBool()) {
                jsonCustomProps.put(name, ctProperty.getBool());
            } else if(ctProperty.isSetBstr()) {
                jsonCustomProps.put(name, ctProperty.getBstr());
            } else {
                jsonCustomProps.put(name, ctProperty.toString());
            }
        }
        
        customProps.getUnderlyingProperties().getPropertyList().forEach(property -> {
            jsonCustomProps.put(property.getName(), property.getLpwstr());
        });

        // Finalize
        ObjectNode jsonProps = mapper.createObjectNode();
        jsonProps.set("sheets", jsonSheets);
        jsonProps.set("coreProperties", jsonCoreProps);
        jsonProps.set("customProperties", jsonCustomProps);

        return jsonProps;
    }
    
    protected String formatDate(Date date) {
        if(date == null) {
            return "";
        }
        
        return dateFormatter.format(date);
    }
}
