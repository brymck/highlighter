Highlighter
===========

Toggles highlight of editable fields in Excel

Purpose
-------

This add-in allows a user to declare a cell as editable or uneditable and
to toggle the cell color of editable cells. The goal is to avoid two common
methods of dealing with variables in Excel sheets: setting all relevant
variables to the side of the print area and referencing them with a formula
(which can make sheets bulky and adds pointless formulas), and making all
variables the intended print color (which is just confusing).

Instead, this add-in reserves two color indexes (1 and 23 by default). By
default, Excel sets 1 to black, which is also the automatic color. 23 is
a light blue. Thus, uneditable cells can be given any other color index, and
editable cells can be toggled between light blue and black. As the automatic
color is also black, there should be no difference in the display of editable
and uneditable cells when highlight mode is off.

Usage
-----

This add-in adds a toolbar called Highlighter to all Excel workbooks. You can
click "Mark as editable/uneditable" to change whether a cell will change color
when highlight mode is on. "Highlight/unhighlight" then toggles that highlight
mode.

Note that for some worksheets (particularly Excel 2007+), you may need to run
"Prepare sheet" first, which will set all cells to the automatic color. This
will overwrite any existing font color information for the current sheet,
although it leaves other formatting unaffected.

Installation
------------

1. Close Excel.
2. Copy Highlighter.xla to `%AppData%\Microsoft\Addins`
3. If this is your first install, do the following in Excel:
   * Excel 2007+
     * Click Office Button > Excel Options... > Add-Ins.
     * In the Manage: dropdown at the bottom, click Excel Add-ins and Go... .
     * Place a checkmark next to Highlighter.
   * Excel 2003-
     * Go to Tools > Add-Ins.
     * Place a checkmark next to Highlighter.