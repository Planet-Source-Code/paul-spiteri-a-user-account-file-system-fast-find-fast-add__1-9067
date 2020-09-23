<div align="center">

## A User Account File System\. Fast Find, Fast Add\!


</div>

### Description

This code is a very quick way of getting data from a file.
 
### More Info
 
RAFSearch, used to get data returns the MyData variable type.

You need to pass a parameter for the ID to find.

RAFAdd needs to be passed the ID of the new record, and as it currently stands, the Name and Password of the 'user' being added. This will need to be changed to suit you.

RAFSearch returns the MyData variable type.

Therefore you should do:

Dim TempData as MyData

TempData = RAFSearch(50)

TempData will now have the data for ID 50.

If TempData.ID = -1, the record does not exist.

RAFAdd returns a boolean. True if it was added successfully, false if it was not - probably a full file, increase the MaxRecords variable and start again.

The MaxRecords can only be changed once, when the RAFClear is executed, which clears and sets up the pointer file.


<span>             |<span>
---                |---
**Submitted On**   |2000-06-23 14:20:38
**By**             |[Paul Spiteri](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-spiteri.md)
**Level**          |Intermediate
**User Rating**    |4.4 (40 globes from 9 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD70206232000\.zip](https://github.com/Planet-Source-Code/paul-spiteri-a-user-account-file-system-fast-find-fast-add__1-9067/archive/master.zip)








