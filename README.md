# font.updater
A Project to Update Font Type in PPTX documents 

1. Install **LibreOffice** on your machine.
2. Open or Import the project into your **Eclipse IDE** or (other IDEs).
3. If there is an error in the project, right click on the project. Click `Properties`. Click `Java Build Path`. Select `Classpath`. Select `program - /usr/lib/libreoffice` and then click the `Edit` button. Replace the path with the directory of the `program` directory under the directory of the LibreOffice installed on your machine. 
4. Replace the `TARGET_DIR` in the `App.java` with the path where your `PPTX` files are located.
	```Java
    private static final String TARGET_DIR = "/home/alfa/Downloads/eclipse/workspace/Test/";
    ```
5. To change the target font type, find this line in the code. Replace `Titillium Web` with your expected font type (e.g., Arial, Courier New, etc). 
   ```Java
    xCursorProps.setPropertyValue("CharFontName", "Titillium Web");
   ```
6. Run the program. Program will load all `PPTX` files in that directory, update the font types, save them, and close them.
