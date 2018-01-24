# Templar

Templar is a letter-template suite management tool. It allows for letter templates to be generated on demand by pulling together pre-composed document fragments and merging them together. It was designed for a specific use-case, but the variable names have been changed to general terms to help you adapt it to your needs.

To set up a Templar letter suite, create a new Google Spreadsheet with at least 3 sheets, to work with the current code. 

<ul>
<li>Sheet 1 is a landing page. Put a greeting here for your users or something similar, if you'd like.
<li>Sheet 2 (the "logic sheet") should have a single row up top reserved for categories on the X axis and two columns on the left reserved for Y-axis categories (note only column B is actually referenced - column A can be used for higher-level grouping of the categories in column B). 
<li>Sheet 3 (the "key sheet") should consist of 2 columns with a row reserved up top for the column names: Document # and Document Description. Under Document #, you'd put an arbitrary number to reference the document fragment to the right. Under Document Description, you'd put a human-readable title or description of the document fragment. This text should then be hyperlinked to a Google Document that houses the formatted fragment in question. <b>NOTE</b> This link must be derived from the sharing link from within the actual view of the document. It does not work with the sharing link copied from Google Drive for some reason. All other cells within the grid should consist of numerical formulas that tell Templar which Document #s (from the below "key sheet") to pull. Each number should be separated by ", " (a comma and a space). 
</ul>

You can then add the attached code to the Google Spreadsheet itself to get a working suite! Generated templates will appear in your Google Drive home folder.
