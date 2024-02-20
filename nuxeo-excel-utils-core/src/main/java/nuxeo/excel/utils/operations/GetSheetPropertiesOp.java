package nuxeo.excel.utils.operations;

import org.apache.commons.lang3.StringUtils;
import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Context;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.PathRef;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 *
 */
@Operation(id=GetSheetPropertiesOp.ID, category=Constants.CAT_DOCUMENT, label="Excel.GetProperties", description="Describe here what your operation does.")
public class GetSheetPropertiesOp {

    public static final String ID = "Document.GetSheetPropertiesOp";

    @Context
    protected CoreSession session;

    @Param(name = "path", required = false)
    protected String path;

    @OperationMethod
    public DocumentModel run() {
        if (StringUtils.isBlank(path)) {
            return session.getRootDocument();
        } else {
            return session.getDocument(new PathRef(path));
        }
    }

    @OperationMethod
    public void run(Blob input) {

        try (InputStream is = new FileInputStream(input.getFile())) {
            Workbook workbook = WorkbookFactory.create(is);
            if (workbook instanceof XSSFWorkbook) {
                readXLSXProperties((XSSFWorkbook) workbook);
            } else if (workbook instanceof HSSFWorkbook) {
                readXLSProperties((HSSFWorkbook) workbook);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void readXLSXProperties(XSSFWorkbook workbook) {
        POIXMLProperties props = workbook.getProperties();
        POIXMLProperties.CoreProperties coreProps = props.getCoreProperties();
        System.out.println("Title: " + coreProps.getTitle());
        System.out.println("Author: " + coreProps.getCreator());

        // To read custom properties
        POIXMLProperties.CustomProperties customProps = props.getCustomProperties();
        /*
        customProps.getUnderlyingProperties().forEach(property -> {
            System.out.println(property.getName() + ": " + property.getLpwstr());
        });*/
    }

    private static void readXLSProperties(HSSFWorkbook workbook) {
        ExcelExtractor extractor = new ExcelExtractor(workbook);
        System.out.println("Author: " + extractor.getSummaryInformation().getAuthor());
        System.out.println("Title: " + extractor.getSummaryInformation().getTitle());
        // HSSF does not directly support reading custom properties like XSSF,
        // custom handling might be required for detailed custom properties.
    }
}
