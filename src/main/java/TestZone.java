import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XReplaceDescriptor;
import com.sun.star.util.XReplaceable;
import ooo.connector.BootstrapSocketConnector;

import java.io.File;


public class TestZone {
    public static void main(String[] args) throws Exception {

        //Initialise
        String oooExeFolder = "C:\\Program Files (x86)\\OpenOffice 4\\program";
        XComponentContext xContext = BootstrapSocketConnector.bootstrap(oooExeFolder);
        XMultiComponentFactory xMCF = xContext.getServiceManager();
        Object oDesktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
        XDesktop xDesktop = (XDesktop) UnoRuntime.queryInterface(XDesktop.class, oDesktop);

        // Load the Document
        String workingDir = "D:/Test/";
        String myTemplate = "test.docx";

        if (!new File(workingDir + myTemplate).canRead()) {
            throw new RuntimeException("Cannot load template:" + new File(workingDir + myTemplate));
        }

        XComponentLoader xCompLoader = (XComponentLoader) UnoRuntime
                .queryInterface(com.sun.star.frame.XComponentLoader.class, xDesktop);

        String sUrl = "file:///" + workingDir + myTemplate;

        PropertyValue[] propertyValues = new PropertyValue[0];

        propertyValues = new PropertyValue[1];
        propertyValues[0] = new PropertyValue();
        propertyValues[0].Name = "Hidden";
        propertyValues[0].Value = new Boolean(true);

        XComponent xComp = xCompLoader.loadComponentFromURL(
                sUrl, "_blank", 0, propertyValues);

        // Search and replace
        XReplaceDescriptor xReplaceDescr = null;
        XReplaceable xReplaceable = null;

        XTextDocument xTextDocument = (XTextDocument) UnoRuntime
                .queryInterface(XTextDocument.class, xComp);

        xReplaceable = (XReplaceable) UnoRuntime
                .queryInterface(XReplaceable.class, xTextDocument);

        xReplaceDescr = (XReplaceDescriptor) xReplaceable
                .createReplaceDescriptor();

        // mail merge the date
        xReplaceDescr.setSearchString("${name}");
        xReplaceDescr.setReplaceString("Test Human");
        xReplaceable.replaceAll(xReplaceDescr);

        // Export to PDF
        XStorable xStorable = (XStorable) UnoRuntime
                .queryInterface(XStorable.class, xComp);

        propertyValues = new PropertyValue[2];
        propertyValues[0] = new PropertyValue();
        propertyValues[0].Name = "Overwrite";
        propertyValues[0].Value = new Boolean(true);
        propertyValues[1] = new PropertyValue();
        propertyValues[1].Name = "FilterName";
        propertyValues[1].Value = "writer_pdf_Export";

        // Appending the favoured extension to the origin document name
        String myResult = workingDir + "test01.pdf";
        xStorable.storeToURL("file:///" + myResult, propertyValues);

        System.out.println("Saved " + myResult);

        // shutdown
        xDesktop.terminate();
    }
}
