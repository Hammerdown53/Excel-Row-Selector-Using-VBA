# Excel VBA Subroutine: SelectRowsContainingText

This VBA subroutine, `SelectRowsContainingText`, empowers users to quickly select rows in an Excel worksheet that contain specific text strings. By defining an array of texts to search for, users can efficiently highlight rows that match their criteria. I created this becuase everymonth I'm sent an email with hundreds of rows when I only need a few to do my calculations. Using this I can define spefic text found in rows that does not change for it to highlight with one click making a 30 mintue jobs turn into 5 minutes.

## How to Use:
1. Open the Excel workbook containing the worksheet you want to search.
2. Navigate to the worksheet where you wish to perform the text-based selection.
3. Run the `SelectRowsContainingText` subroutine.
4. Enter the desired text strings to search for in the `searchText` array.
5. The subroutine will automatically select rows containing any of the specified texts, simplifying data analysis tasks.

To create this macro, you will need to use the personal file as show below within VBA and add the module within that one so that this is applied to every excel doc you open.
![alt text](https://raw.githubusercontent.com/Hammerdown53/Excel-Row-Selector-Using-VBA/main/Screenshot%202024-07-15%20095411.png)

## Key Features:
- **Flexible array-based search:** Define multiple text strings to search for within the worksheet.
- **Dynamic row selection:** Automatically highlights entire rows containing the specified text, aiding in data visualization and analysis.

## Contributing:
Contributions to enhance the functionality of this VBA subroutine are welcome. Feel free to suggest improvements or report any issues by opening an issue or submitting a pull request.

## License:
This project is licensed under the MIT License.

## Credits:
Developed by Trey McBride. Enjoy streamlining your Excel data manipulation tasks with this handy VBA tool!
