function insertDataWithStyle(sheet, rangeNotation, data, options) {
  var range = sheet.getRange(rangeNotation);

  // Set data
  if (Array.isArray(data)) {
    range.setValues(data); // If data is an array, set multiple values
  } else {
    range.setValue(data); // Set single value
  }
  
  // Set font size
  if (options.fontSize) {
    range.setFontSize(options.fontSize);
  }
  
  // Set font color
  if (options.fontColor) {
    range.setFontColor(options.fontColor);
  }

  // Set background color
  if (options.backgroundColor) {
    range.setBackground(options.backgroundColor);
  }

  // Set font family
  if (options.fontFamily) {
    range.setFontFamily(options.fontFamily);
  }
  
  // Set bold
  if (options.bold !== undefined) {
    range.setFontWeight(options.bold ? 'bold' : 'normal');
  }
  
  // Set italic
  if (options.italic !== undefined) {
    range.setFontStyle(options.italic ? 'italic' : 'normal');
  }
  
  // Set underline
  if (options.underline !== undefined) {
    range.setUnderline(options.underline);
  }
  
  // Set horizontal alignment
  if (options.horizontalAlignment) {
    range.setHorizontalAlignment(options.horizontalAlignment);
  }

  // Set vertical alignment
  if (options.verticalAlignment) {
    range.setVerticalAlignment(options.verticalAlignment);
  }

  // Set text wrapping
  if (options.wrapText !== undefined) {
    range.setWrap(options.wrapText);
  }

  // Set text rotation (angle)
  if (options.textRotation) {
    range.setTextRotation(options.textRotation);
  }
  
  // Set borders (optional: specify sides and border styles)
  if (options.border) {
    var borderOptions = {
      top: options.border.top || false,
      bottom: options.border.bottom || false,
      left: options.border.left || false,
      right: options.border.right || false,
      color: options.border.color || 'black',
      style: options.border.style || SpreadsheetApp.BorderStyle.SOLID
    };
    range.setBorder(
      borderOptions.top, 
      borderOptions.left, 
      borderOptions.bottom, 
      borderOptions.right, 
      false, 
      false, 
      borderOptions.color, 
      borderOptions.style
    );
  }
}
