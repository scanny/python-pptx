.. _table_merge:

Cell Merge
==========

PowerPoint allows a user to *merge* table cells, causing multiple *grid
cells* to appear and behave as a single cell.


Terminology
-----------

The following distinctions are important to the implementation and also to
using the API.

table grid
   A PowerPoint (DrawingML) table has an underlying grid which is a strict
   two-dimensional (n X m) array; each row in the table grid has m cells and
   each column has n cells. Grid cell boundaries are aligned both vertically
   and horizontally.

grid cell
   Each of the n x m cells in a table is a grid cell. Every grid cell is
   addressable, even those "shadowed" by a merge-origin cell.

   *Apparent* cell boundaries can be modified by *merging* two or more grid
   cells.

merge-origin cell
   The top-left grid cell in a merged cell. This cell will be visible, have
   the combined extent of all spanned cells, and contain any visible content.

spanned cell
   All grid cells in a merged cell are spanned cells, excepting the one at
   the top-left (the merge-origin cell). Intuitively, the merge-origin cell
   "spans" all the other grid cells in the merge range.

   Content in a spanned cell is not visible, and is typically moved to the
   merge-origin cell as part of the merge operation, leaving each spanned
   cell with exactly one empty paragraph. This content becomes visible again
   if the merged cell is split.

unmerged cell
   While not a distinction directly embodied in the code, any cell which is
   not a merge-origin cell and is also not a spanned cell is a "regular" cell
   and not part of a merged cell.


Protocol
--------

A cell is accessed by grid location (regardless of any merged regions). All
grid cells are addressable, even those "shadowed" by a merge (such a shadowed
cell is a *spanned* cell)::

    >>> table = shapes.add_table(rows=3, cols=3, x, y, cx, cy).table
    >>> len(table.rows)
    3
    >>> len(table.columns)
    3
    >>> len(table.rows[0].cells)
    3

Merge 2 x 2 cells at top left::

    >>> cell = table.cells(0, 0)
    >>> other_cell = table.cells(1, 1)
    >>> cell.merge(other_cell)

The top-left cell of a merged cell is the merge-origin cell::

    >>> origin_cell = table.cells(0, 0)
    >>> origin_cell
    <pptx.table._MergeOriginCell object at 0x...>
    >>> origin_cell.is_merge_origin
    True
    >>> origin_cell.is_spanned
    False
    >>> origin_cell.span_height
    1
    >>> origin_cell.span_width
    2

A spanned cell has |True| on its `.is_spanned` property::

    >>> spanned_cell = table.cell(0, 1)
    >>> spanned_cell
    <pptx.table._SpannedCell object at 0x...>
    >>> spanned_cell.is_merge_origin
    False
    >>> spanned_cell.is_spanned
    True

A "regular" cell not participating in a merge is neither the merge origin nor
spanned::

    >>> cell = table.cell(0, 2)
    >>> cell
    <pptx.table._Cell object at 0x...>
    >>> cell.is_merge_origin
    False
    >>> cell.is_spanned
    False

Cell instances proxying the same `a:tc` element compare equal::

    >>> origin_cell == table.cell(0, 0)
    True
    >>> spanned_cell == table.cell(0, 1)
    True
    >>> cell == table.cell(0, 2)
    True
    >>> origin_cell == spanned_cell
    False


Use Cases
---------

Use Case: Interrogate table for merged cells::

    def iter_merge_origins(table):
        """Generate each merge-origin cell in *table*.

        Cell objects are ordered by their position in the table,
        left-to-right, top-to-bottom.
        """
        return (cell for cell in table.iter_cells() if cell.is_merge_origin)

    def merged_cell_report(cell):
        """Return str summarizing position and size of merged *cell*."""
        return (
            'merged cell at row %d, col %d, %d cells high and %d cells wide'
            % (cell.row_idx, cell.col_idx, cell.span_height, cell.span_width)
        )

    # ---Print a summary line for each merged cell in *table*.---
    for merge_origin_cell in iter_merge_origins(table):
        print(merged_cell_report(merge_origin_cell))

Use Case: Access only cells that display text (are not spanned)::

    def iter_visible_cells(table):
        return (cell for cell in table.iter_cells() if not cell.is_spanned)

Use Case: Determine whether table contains merged cells::

    def has_merged_cells(table):
        for cell in table.iter_cells():
            if cell.is_merge_origin:
                return True


PowerPoint behaviors
--------------------

* Two or more cells are merged by selecting them using the mouse, then
  selecting "Merge cells" from the context menu.

* Content from spanned cells is moved to the merge origin cell.

* A merged cell can be split ("unmerged" roughly). The UI allows the merge to
  be split into an arbitrary number of rows and columns and adjusts the table
  grid and row heights etc. to accommodate, adding (potentially very many)
  new merged cells as required.

  `python-pptx` just removes the merge, restoring the underlying table grid
  cells to regular (unmerged) cells.


Specimen XML
------------

.. highlight:: xml

Super-simplified 3-cell horizontal merge::

  <a:tr>
    <a:tc gridSpan="3"/>
    <a:tc hMerge="true"/>  <!-- PowerPoint uses boolean value "1" -->
    <a:tc hMerge="true"/>
  </a:tr>

Super-simplified 3-cell vertical merge::

  <a:tr>
    <a:tc rowSpan="3"/>
  </a:tr>
  <a:tr>
    <a:tc vMerge="true"/>  <!-- PowerPoint uses boolean value "1" -->
  </a:tr>
  <a:tr>
    <a:tc vMerge="true"/>
  </a:tr>

Super-simplified 2D merge::

  <a:tr>
    <a:tc rowSpan="3" gridSpan="3"/>
    <a:tc rowSpan="3" hMerge="true"/>
    <a:tc rowSpan="3" hMerge="true"/>
  </a:tr>
  <a:tr>
    <a:tc gridSpan="3" vMerge="true"/>
    <a:tc hMerge="true" vMerge="true"/>
    <a:tc hMerge="true" vMerge="true"/>
  </a:tr>
  <a:tr>
    <a:tc gridSpan="3" vMerge="true"/>
    <a:tc hMerge="true" vMerge="true"/>
    <a:tc hMerge="true" vMerge="true"/>
  </a:tr>

Simplified 2 x 3 table with first two horizontal cells merged::

  <a:tbl>
    <a:tblGrid>
      <a:gridCol w="2032000"/>
      <a:gridCol w="2032000"/>
      <a:gridCol w="2032000"/>
    </a:tblGrid>
    <a:tr h="370840">
      <a:tc gridSpan="2">
        <a:txBody>...</a:txBody>
        <a:tcPr/>
      </a:tc>
      <a:tc hMerge="1">
        <a:txBody>...</a:txBody>
        <a:tcPr/>
      </a:tc>
      <a:tc>
        <a:txBody>...</a:txBody>
        <a:tcPr/>
      </a:tc>
    </a:tr>
    <a:tr h="370840">
      <a:tc>...</a:tc>
      <a:tc>...</a:tc>
      <a:tc>...</a:tc>
    </a:tr>
  </a:tbl>

Simplified 2 x 3 table with first two vertical cells merged::

  <a:tbl>
    <a:tr h="370840">
      <a:tc rowSpan="2">
        <a:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>Vertical</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </a:p>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>Span</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </a:p>
        </a:txBody>
        <a:tcPr/>
      </a:tc>
      <a:tc>...</a:tc>
      <a:tc>...</a:tc>
    </a:tr>
    <a:tr h="370840">
      <a:tc vMerge="1">
        <a:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </a:p>
        </a:txBody>
        <a:tcPr/>
      </a:tc>
      <a:tc>...</a:tc>
      <a:tc>...</a:tc>
    </a:tr>
  </a:tbl>


Schema excerpt
--------------

::

  <xsd:complexType name="CT_TableCell">
    <xsd:sequence>
      <xsd:element name="txBody" type="CT_TextBody"               minOccurs="0"/>
      <xsd:element name="tcPr"   type="CT_TableCellProperties"    minOccurs="0"/>
      <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
    </xsd:sequence>
    <xsd:attribute name="rowSpan"  type="xsd:int"     default="1"/>
    <xsd:attribute name="gridSpan" type="xsd:int"     default="1"/>
    <xsd:attribute name="hMerge"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="vMerge"   type="xsd:boolean" default="false"/>
    <xsd:attribute name="id"       type="xsd:string"/>
  </xsd:complexType>
