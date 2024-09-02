package font.updater;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import java.io.File;
import java.util.Objects;

public class App {
	// Define the target directory path
	private static final String TARGET_DIR = "/home/alfa/Downloads/eclipse/workspace/Test/";

	public static void main(String[] args) {
		try {
			// Bootstrap LibreOffice
			XComponentContext xContext = Bootstrap.bootstrap();
			if (xContext == null) {
				System.out.println("ERROR: Could not bootstrap LibreOffice.");
				return;
			}

			// Get the ServiceManager
			XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class,
					xContext.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", xContext));

			// Iterate over all .pptx files in the target directory
			File dir = new File(TARGET_DIR);
			File[] pptxFiles = dir.listFiles((d, name) -> name.toLowerCase().endsWith(".pptx"));

			if (pptxFiles == null || pptxFiles.length == 0) {
				System.out.println("No .pptx files found in the directory.");
				return;
			}

			for (File pptxFile : Objects.requireNonNull(pptxFiles)) {
				String filePath = "file:///" + pptxFile.getAbsolutePath().replace("\\", "/");

				// Load the Impress document
				XComponent xComponent = xComponentLoader.loadComponentFromURL(filePath, "_blank", 0,
						new PropertyValue[0]);

				// Access the draw pages supplier
				XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent);
				XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();

				// Iterate over all slides in the presentation
				for (int i = 0; i < xDrawPages.getCount(); i++) {
					XDrawPage xDrawPage = UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex(i));

					// Iterate over all shapes on the slide
					XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, xDrawPage);
					for (int j = 0; j < xShapes.getCount(); j++) {
						XShape xShape = UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(j));

						// Check if the shape contains text
						XText xText = UnoRuntime.queryInterface(XText.class, xShape);
						if (xText != null) {
							System.out.println("Updating text in " + pptxFile.getName() + ": " + xText.getString());

							// Create a text cursor
							XTextCursor xTextCursor = xText.createTextCursor();

							// Select all text
							xTextCursor.gotoStart(false);
							xTextCursor.gotoEnd(true);

							// Access the properties of the selected text
							XPropertySet xCursorProps = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);

							// Change the font to another font
							xCursorProps.setPropertyValue("CharFontName", "Titillium Web");
						}
					}
				}

				// Save the modified presentation as a .pptx file using impress8 filter
				XStorable xStorable = UnoRuntime.queryInterface(XStorable.class, xComponent);
				if (xStorable != null) {
					PropertyValue[] saveProps = new PropertyValue[1];
					saveProps[0] = new PropertyValue();
					saveProps[0].Name = "FilterName";
					saveProps[0].Value = "impress8";

					xStorable.store();

					// Save the file with "_converted" appended to the original file name
//                    String outputFilePath = "file:///" + pptxFile.getAbsolutePath().replace("\\", "/").replace(".pptx", "_converted.pptx");
//                    xStorable.storeAsURL(outputFilePath, saveProps);

				}

				// Close the document
				xComponent.dispose();
			}

			System.out.println("FINISHED!");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
