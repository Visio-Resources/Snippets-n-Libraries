# VBA Basics - Getting a handle on shapes
## Modifiying properties of shapes
```VBA
Sub Foo()
  DIM shp as Shape
  Set shp = ... 'here comes one of the methods described bellow to get a handle on a shape
  'Properties of a shape are modified by setting the value of a cell of its shapesheet.
  shp.Cells("PinX").Formula = "1"
  shp.Cells("LineWeight").Formula = "1pt."
End sub
```
**Remarks:**
- Mind the differences between Formula and FormulaU.
- There are properties which are not included in the shapesheet. e.g. shp.Text
### Reading properties of a shape
```VBA
Sub Foo()
  DIM shp as Shape
  DIM t as string
  Set shp = ... 'here comes one of the methods described bellow to get a handle on a shape
  'Properties of a shape are read by using of the multiple "ResultX" methods of a cell of its shapesheet.
  Debug.print shp.Cells("PinX").ResultStr("")
End sub
```

## Getting a handle on the shapes/objects to process
Regardless of the operations to perform, you need first to get a handle on the object to process.

**handling the first selected shape**
**handling exactly one shape**

**handling shapes with direct processing on the current page**
```VBA
for each shp in ActivePage.Shapes
  'do stuff directly
  
  'of check first if the shape complies with certain requirements
  if shp.Cells(...).ResultIU = ... then
    'do stuff here
  end if
next shp
```
**handling every shape on every page**
```VBA
Dim pg as page
Dim shp as page
for each pg in ActiveDocument.Pages
  for each shp in pg.Shapes
    'do stuff
   next shp
next pg
```
**handling shapes and their sub-shapes**
```VBA
Sub Foo()
  Bar ActivePage.Shapes
End sub

Sub Bar(shps as Shapes)
  Dim shp as Shape
  for each shp in shps
    'do stuff to shp itself, then call the same procedure on its sub-shapes
    Bar shp.Shapes
  next shp
End sub
```
**processing manually selected shapes**
```VBA
for each shp in ActiveWindow.Selection
  'do stuff
next shp
```
**working with selections**
```VBA
Sub Foo()
  ActiveWindow.DeselectAll
  for each shp in ActivePage.Shapes
    if ... then
      ActiveWindow.Select shp, visSelect
    end if
  next shp
  Bar ActiveWindow.Selection
End sub

Sub Bar(sel as Selection)
  Dim shp as shp
  For each shp in sel
    'do stuff
  next shp
End sub
```

#### Preparing shapes for selection
If certain certain selections will be needed frequently it makes sence to shorten the selection process.
**Define a "category" for shapes**
```VBA
Sub SetupShapes
  Dim shp as Shape
  'Use the ActiveWindow.Selection or one of the previous other methods
  For each shp in ActiveWindow.Selection
    If not shp.SectionExists(visSectionUser, False) then
      shp.AddSection visSectionUser
    End if
    If not shp.CellExists("user.foo", False) then
      shp.AddNamedRow visSectionUser, "shapeType", visTagDefault
    End if
    shp.Cells("user.shapeType").Formula = chr(34) & "xyz" & chr(34)
  Next shp
End sub
```
When needed the shapes are they called by
```VBA
For each shp in ActivePage.Shapes
  If shp.SectionExists(visSectionUser, False) then
    If shp.CellExists("user.shapeType", False) then
      If shp.Cells("user.shapeType").ResultStr("") = "xyz" then
        'Do stuff
      End if
    End if
  End if
Next shp
```
This can be shortened by:
```VBA
Sub Foo
  On Error Goto Errhandler
  Dim shp as Shape
  For each shp in ActivePage.Shapes
    If shp.Cells("user.shapeType").ResultStr("") = "xyz" then 'This line will raise an error if the cell does not exist.
      'Do stuff 1
      'Do stuff 2
      'Do stuff 3
ResumePoint:
    End if
  Next shp
Exit sub
ErrHandler:
  'optional Debug information
  Err.clear
  Resume ResumePoint
End sub
```
