🔗 ConcatenateColumnsCAndD_AB — Merge Column Values into a Single Cell
This VBA macro concatenates values from columns C and D in the "BL-BID" sheet starting from row 3, and inserts the combined result into a merged cell in column F (rows 3 to 30).

🔹 What it does:
Loops through all rows (starting from row 3) until the last non-empty cell in either column C or D.
For each row, if both column C and D contain data, it creates a string in the format:

ValueC1(ValueD1), ValueC2(ValueD2), ...
After building the string, it:

Removes the trailing comma and space.
Inserts the result into a merged block of cells from F3:F30.
Aligns the text to top-left for readability.

🧾 Example Output:
If your data looks like this:


C	D
Cable A	100m
Conduit B	200m
The merged cell in F3:F30 will display:


Cable A(100m), Conduit B(200m)
✅ Highlights:
Dynamically calculates the maximum of used rows in both C and D, ensuring complete coverage.

Clean output format, removing trailing punctuation.

Supports customization by adjusting the range or sheet name.

Ready to use in summary sections or documentation fields.

💡 Example Usage:

Sub ConcatenateColumnsCAndD_AB()
    ' Run this to merge values from C & D into one formatted string in F3:F30
End Sub
Tip: If you want to include a fixed prefix or title before the string (e.g., "Material List: "), just assign a value to the fixedText variable before it's combined.

