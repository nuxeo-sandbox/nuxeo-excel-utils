package nuxeo.excel.utils;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import java.io.File;

import javax.inject.Inject;

import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.nuxeo.common.utils.FileUtils;
import org.nuxeo.ecm.automation.AutomationService;
import org.nuxeo.ecm.automation.OperationContext;
import org.nuxeo.ecm.automation.test.AutomationFeature;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.core.test.DefaultRepositoryInit;
import org.nuxeo.ecm.core.test.annotations.Granularity;
import org.nuxeo.ecm.core.test.annotations.RepositoryConfig;
import org.nuxeo.runtime.test.runner.Deploy;
import org.nuxeo.runtime.test.runner.Features;
import org.nuxeo.runtime.test.runner.FeaturesRunner;

import nuxeo.excel.utils.operations.GetPropertiesOp;

@RunWith(FeaturesRunner.class)
@Features(AutomationFeature.class)
@RepositoryConfig(init = DefaultRepositoryInit.class, cleanup = Granularity.METHOD)
@Deploy("nuxeo.excel.utils.nuxeo-excel-utils-core")
public class TestGetSheetPropertiesOp {

    @Inject
    protected CoreSession session;

    @Inject
    protected AutomationService automationService;

    @Test
    public void shouldGetProperties() throws Exception {
        
        File testFile = FileUtils.getResourceFileFromContext("test-properties.xlsx");
        
        OperationContext ctx = new OperationContext(session);
        ctx.setInput(new FileBlob(testFile));
        Blob blob = (Blob) automationService.run(ctx, GetPropertiesOp.ID);
        
        assertNotNull(blob);
        
        JSONObject mainJson = new JSONObject(blob.getString());
        
        JSONArray sheets = mainJson.getJSONArray("sheets");
        assertNotNull(sheets);
        assertEquals(2, sheets.length());
        assertEquals("List", sheets.get(0));
        assertEquals("otherSheet", sheets.get(1));
        
        JSONObject coreProperties = mainJson.getJSONObject("coreProperties");
        assertNotNull(coreProperties);
        assertEquals("Title for Test", coreProperties.get("title"));
        assertEquals("The Comment", coreProperties.get("description"));
        assertEquals("The Author", coreProperties.get("creator"));
        assertEquals("kw1 kw2 kw3", coreProperties.get("keywords"));
        assertEquals("2017-02-03T05:31:38+0100", coreProperties.get("created"));
        
        
        JSONObject customProperties = mainJson.getJSONObject("customProperties");
        assertNotNull(customProperties);
        assertEquals("Custom Department", customProperties.get("Department"));
        assertEquals("Custom Client", customProperties.get("Client"));
        //Not handling date/numerical/etc. types
        
    }

}
