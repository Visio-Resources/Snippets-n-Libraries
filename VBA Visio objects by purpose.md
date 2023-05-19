
Objects commonly handled in Macros, ordered by (percieved) commonness and the tasks associated with them.

### Shape
- modify the geometry of a shape. Targetting cells like PinX, PinY, Width, Height, Angle, etc.
```VBA
shp.Cells("Width").Formula = shp.Cells("Height").ResultIU
```
- modify the path of the geometry of a shape
```VBA
for i = 1 to 10
  shp.Cells("Geometry1.X"&i).Formula = R*cos(i)
  shp.Cells("Geometry1.Y"&i).Formula = R*sin(i)
 next i
```
- Adding user or prop cells to shapes
```VBA
if not shp.SectionExists(visSectionUser,False) then
  shp.AddSection visSectionUser
end if
if not shp.CellExists("user.spam", False) then
  shp.AddNamedRow visSectionUser, "spam", visTagDefault
end if
```

### Selection and Shapes
- iterate over shapes

### Application, Document and Documents

### Page and Pages

### Layer and Layers
