
# insertDataWithStyle

This function allows you to insert data into a specific cell or range in Google Sheets while applying customizable styles such as font size, color, bold, italics, background color, alignment, borders, and more. It supports all styling options provided by Google Sheets.

## Features

- Insert data into a cell or range
- Customize font size, color, and family
- Apply bold, italic, underline styles
- Set background color
- Align text horizontally and vertically
- Enable/disable text wrapping
- Rotate text
- Add borders (customizable for each side)
- Works with single cells or entire ranges

## Usage

### Function Signature

```javascript
function insertDataWithStyle(sheet, rangeNotation, data, options)
```

### Parameters

- `sheet`: A reference to the active Google Sheet where data should be inserted.
- `rangeNotation`: The target cell or range (e.g., `"A1"`, `"A2:B5"`).
- `data`: The data to insert. Can be a string for single cells or a 2D array for ranges.
- `options`: An object that specifies various styling options. See the full list of available options below.

### Available Styling Options

You can pass the following options in the `options` object to style the cell or range:

| Option                | Type    | Description                                                                 |
|-----------------------|---------|-----------------------------------------------------------------------------|
| `fontSize`            | Number  | The font size of the text (e.g., `14`)                                      |
| `fontColor`           | String  | The font color in hex format (e.g., `'#FF0000'` for red)                    |
| `backgroundColor`     | String  | The background color in hex format (e.g., `'#FFFF00'` for yellow)           |
| `fontFamily`          | String  | The font family (e.g., `'Arial'`, `'Roboto'`, `'Ubuntu'`)                   |
| `bold`                | Boolean | Whether the text is bold (`true` for bold, `false` for normal)              |
| `italic`              | Boolean | Whether the text is italic (`true` for italic, `false` for normal)          |
| `underline`           | Boolean | Whether the text is underlined (`true` for underline, `false` for normal)   |
| `horizontalAlignment` | String  | Horizontal alignment (`'left'`, `'center'`, `'right'`)                      |
| `verticalAlignment`   | String  | Vertical alignment (`'top'`, `'middle'`, `'bottom'`)                        |
| `wrapText`            | Boolean | Whether to enable text wrapping (`true` to wrap, `false` to not wrap)       |
| `textRotation`        | Number  | Rotation of the text (e.g., `45` for 45-degree rotation)                    |
| `border`              | Object  | Border options (customizable for each side: top, bottom, left, right)       |

### Border Options

If you use the `border` option, you can specify individual sides and customize their style and color:

- `top`: `true` or `false` (show or hide top border)
- `bottom`: `true` or `false` (show or hide bottom border)
- `left`: `true` or `false` (show or hide left border)
- `right`: `true` or `false` (show or hide right border)
- `color`: Border color in hex format (e.g., `'#000000'` for black)
- `style`: Border style. Available styles include:
  - `SpreadsheetApp.BorderStyle.SOLID`
  - `SpreadsheetApp.BorderStyle.SOLID_MEDIUM`
  - `SpreadsheetApp.BorderStyle.SOLID_THICK`
  - `SpreadsheetApp.BorderStyle.DOTTED`
  - `SpreadsheetApp.BorderStyle.DASHED`

### Example: Insert Data with Full Styling

```javascript
function testInsertDataWithStyle() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var options = {
    fontSize: 14,
    fontColor: '#FF0000',  // Red font
    backgroundColor: '#FFFF00',  // Yellow background
    fontFamily: 'Arial',
    bold: true,
    italic: false,
    underline: true,
    horizontalAlignment: 'center',
    verticalAlignment: 'middle',
    wrapText: true,
    textRotation: 45,  // Rotate text 45 degrees
    border: {
      top: true,
      bottom: true,
      left: true,
      right: true,
      color: '#000000',  // Black border
      style: SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    }
  };
  
  // Insert a single value into cell A1
  insertDataWithStyle(sheet, 'A1', 'Styled Data', options);

  // Insert multiple values into range A2:B3
  var data = [
    ['A', 'B'],
    ['C', 'D']
  ];
  insertDataWithStyle(sheet, 'A2:B3', data, options);
}
```

### Options Breakdown

#### Font Size
```javascript
fontSize: 12
```
Sets the font size to 12.

#### Font Color
```javascript
fontColor: '#00FF00'
```
Sets the font color to green.

#### Background Color
```javascript
backgroundColor: '#FFFF00'
```
Sets the background color to yellow.

#### Font Family
```javascript
fontFamily: 'Courier New'
```
Sets the font to "Courier New."

#### Bold Text
```javascript
bold: true
```
Makes the text bold.

#### Italic Text
```javascript
italic: true
```
Italicizes the text.

#### Underline Text
```javascript
underline: true
```
Underlines the text.

#### Horizontal Alignment
```javascript
horizontalAlignment: 'center'
```
Centers the text horizontally.

#### Vertical Alignment
```javascript
verticalAlignment: 'middle'
```
Aligns the text vertically in the middle of the cell.

#### Text Wrapping
```javascript
wrapText: true
```
Enables text wrapping within the cell.

#### Text Rotation
```javascript
textRotation: 90
```
Rotates the text by 90 degrees.

#### Borders
```javascript
border: {
  top: true,
  bottom: true,
  left: true,
  right: true,
  color: '#0000FF',
  style: SpreadsheetApp.BorderStyle.DOTTED
}
```
Adds dotted blue borders to all sides of the cell.

### Multiple Cells (Range)
If you're applying this function to a range, you can pass a 2D array to the `data` parameter:

```javascript
var data = [
  ['A', 'B', 'C'],
  ['D', 'E', 'F']
];
insertDataWithStyle(sheet, 'A1:C2', data, options);
```

This will insert data across the specified range while applying all the given styles.

## License

This code is licensed under the MIT License.
