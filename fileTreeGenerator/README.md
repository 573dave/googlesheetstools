
# Google Drive File Tree Generator

This script generates a file tree in Google Sheets from a Google Drive folder structure. It provides a visual representation of the files and folders, complete with icons and hyperlinks. The script also uses a sidebar for interaction and displays progress as the tree is being generated.

## Features

- Generate a file tree of the current folder in Google Sheets.
- Supports various Google Drive file types and assigns appropriate icons.
- Provides hyperlinks to open files directly from the sheet.
- Warns the user if the sheet is not empty before generating the file tree.
- Custom styling, including dark mode for the sheet.

## Components

### 1. **Main Script (`Code.gs`)**

This script defines the core functionality to generate the file tree and handle user interactions.

#### Key Functions:
- `onOpen()`: Adds a custom menu to Google Sheets to launch the sidebar.
- `showSidebar()`: Opens the sidebar where the user can initiate the file tree generation process.
- `generateTree()`: Generates the file tree and displays progress.
- `processFolder()`: Recursively processes folders and files to create the tree structure.
- `formatItem()`: Formats each item (file or folder) with an icon and a hyperlink.
- `getIconForType()`: Assigns appropriate icons to different file types.
- `resizeColumns()`: Adjusts the columns for better visualization.
- `styleSheet()`: Applies dark mode and custom styles to the sheet.
- `isSheetEmpty()`: Checks if the sheet is empty before overwriting data.
- `setProgress()` & `getProgress()`: Handles progress tracking.
- `getFoldersInRoot()`: Retrieves folders from the root directory.

### 2. **Sidebar (`Sidebar.html`)**

The HTML file for the sidebar interface. It allows users to trigger the file tree generation.

#### Key Features:
- Displays the directory name.
- Provides a "Generate Tree" button.
- Shows a progress bar as the tree generation progresses.

## How to Use

1. **Setup:**
   - Create a new Google Sheet.
   - Open the Script Editor (`Extensions` > `Apps Script`).
   - Create two files: `Code.gs` for the script and `Sidebar.html` for the sidebar.

2. **Insert Code:**
   - Paste the provided JavaScript code into `Code.gs`.
   - Paste the sidebar HTML code into `Sidebar.html`.

3. **Run the Script:**
   - Refresh the Google Sheet.
   - You should now see a new custom menu named "Drive Tree" in the Google Sheets UI.
   - Click `Drive Tree > Show Sidebar` to open the sidebar.

4. **Generate File Tree:**
   - The sidebar will display the current directory name.
   - Click "Generate Tree" to start generating the file tree.
   - A progress bar will appear, showing the percentage of completion.
   - The tree will be displayed in the active sheet with icons and hyperlinks.

## Customization

- **Icons:** You can modify the `getIconForType()` function to use different emojis or icons for different file types.
- **Styling:** The `styleSheet()` function can be updated to change the appearance of the sheet (e.g., fonts, colors).
- **Progress Bar:** The progress tracking in `setProgress()` can be customized to display different messages or use a different method to show progress.

## License

This code is licensed under the MIT License.

