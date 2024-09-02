# font.updater
A Project to Update Font Type in PPTX documents 

1. Install **LibreOffice** on your machine.
2. Import the project into your **Eclipse IDE** or (other IDEs).
3. Replace the `TARGET_DIR` in the `App.java` with the path where your `PPTX` files are located.
	```Java
    private static final String TARGET_DIR = "/home/alfa/Downloads/eclipse/workspace/Test/";
    ```
4. To change the target font type, find this line in the code. Replace `Titillium Web` with your expected font type (e.g., Arial, Courier New, etc). 
   ```Java
    xCursorProps.setPropertyValue("CharFontName", "Titillium Web");
   ```
5. Run the program. Program will load all `PPTX` files in that directory, update the font types, save them, and close them.
