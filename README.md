# SKeeper

![alt text](https://github.com/voyteca/SKeeper/raw/master/SKeeper.png)

 SKeeper creates collections of shapes/objects which can be added to selection.

It put a registry to shape's name attribute (you can see it in Object Manager). When you click on the list with groups names it finds all shapes with the same record and adds them to selection.

Registry/records looks like this :    SK:[groupName]

where [groupName] will appear in the SKeeper list box

You can put anything in the object's name property, just keep SKeeper record at the end.

The limit is that the groups cannot overlap, in other words object can only belong to one group. Therefore current group record will be replaced with new one.

The macro uses CQL (Corel Query Language) to search for shapes.

for example:  
    ActiveDocument.SelectableShapes.FindShapes(, , , "@name.Contains('SK:')")
    
first 3 parameters are omitted (notice empty spaces between commas) and last parameter is a String  representing the CQL syntax.

