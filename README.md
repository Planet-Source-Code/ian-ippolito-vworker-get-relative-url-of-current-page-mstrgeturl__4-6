<div align="center">

## Get relative URL of current page: mstrGetURL


</div>

### Description

Returns relative URL of current page. Example: /vb/test.asp. See http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=5 to get the fully qualified URL string.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Beginner
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-get-relative-url-of-current-page-mstrgeturl__4-6/archive/master.zip)





### Source Code

```
function mstrGetRelativeURL()
	mstrGetRelativeURL=Request.serverVariables("PATH_INFO")
end function
```

