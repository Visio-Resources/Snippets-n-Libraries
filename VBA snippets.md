# VBA snippets
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

### Interesting objects and "sub-objects"
#### Shapes
Shapes are the certainly the most important objects, since they are the visible result of anything you do in Visio.
#### Application

#### Documents

#### Pages
#### Windows
#### Documents

### Shapes
#### Selecting all shapes
**direct processing on the current page**
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
    'do stuff
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
#### Using the manually selected shapes

#### Selecting shape by property

#### Preparing shapes for selection
