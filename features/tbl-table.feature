Feature: Table properties and methods
  In order to discover and adjust the formatting of a table
  As a developer using python-pptx
  I need properties and methods on Table


  Scenario: Table.cell()
     Given a 2x2 Table object as table
      Then table.cell(0, 0) is a _Cell object


  Scenario: Table.columns
     Given a Table object as table
      Then table.columns is a _ColumnCollection object


  Scenario: Table.first_col setter
     Given a 2x2 Table object as table
      When I assign table.first_col = True
      Then table.first_col is True


  Scenario: Table.first_row setter
     Given a 2x2 Table object as table
      When I assign table.first_row = True
      Then table.first_row is True


  Scenario: Table.horz_banding setter
     Given a 2x2 Table object as table
      When I assign table.horz_banding = True
      Then table.horz_banding is True


  Scenario: Table.iter_cells()
    Given a 3x3 Table object as table
     Then len(list(table.iter_cells())) == 9


  Scenario: Table.last_col setter
     Given a 2x2 Table object as table
      When I assign table.last_col = True
      Then table.last_col is True


  Scenario: Table.last_row setter
     Given a 2x2 Table object as table
      When I assign table.last_row = True
      Then table.last_row is True


  Scenario: Table.rows
     Given a Table object as table
      Then table.rows is a _RowCollection object


  Scenario: Table.vert_banding setter
     Given a 2x2 Table object as table
      When I assign table.vert_banding = True
      Then table.vert_banding is True
