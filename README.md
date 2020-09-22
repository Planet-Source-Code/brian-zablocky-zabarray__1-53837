<div align="center">

## ZabArray


</div>

### Description

ZabArray stores data in Key/Value pairs like a collection. The key difference is that ZabArrays allow you to serialize the data into a string (PHP function) which can be saved to a text file or the registry. The string can be loaded into a different ZabArray to restore the entire structure.
 
### More Info
 
A key (string) and and associated value. Value is solid data types like Int, Lng, Sng, Dbl, Cur, String, Byte, Bool, and Date. Also accepts a single-line string representing the data that was created from another ZabArray.

You cannot store arrays or objects in a ZabArray because of problems converting them to a string.

The same data that was input. Also returns a single-line string representing the data for restoration later.

The memory used by the data is 10% larger than the actual data. During serialization and unserialization the data is 120% larger until the process is complete.


<span>             |<span>
---                |---
**Submitted On**   |2004-05-19 17:35:48
**By**             |[Brian Zablocky](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-zablocky.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[ZabArray1748015202004\.zip](https://github.com/Planet-Source-Code/brian-zablocky-zabarray__1-53837/archive/master.zip)

### API Declarations

```
'
' No API. Everything is hard-coded.
'
```





