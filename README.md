<div align="center">

## Allow users to post "Safe" HTML


</div>

### Description

This code pulls out all the nasty tags that a user sholdn't use when posting content. It also pulls out any javascript events assigned to any tags. A must have if you allow people to post HTML on your site.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-allow-users-to-post-safe-html__4-6656/archive/master.zip)

### API Declarations

(c)Copyright 2001 Lewis Edward Moten III, All rights reserved.


### Source Code

```
Function SafeHTML(ByVal pStrHTML)
	Dim lObjRegExp
	If VarType(pStrHTML) = vbNull Then Exit Function
	If pStrHTML = "" Then Exit Function
	Set lObjRegExp = New RegExp
	lObjRegExp.Global = True
	lObjRegExp.IgnoreCase = True
	lObjRegExp.Pattern = "<(/)?SCRIPT|META|STYLE([^>]*)>"
	pStrHTML = lObjRegExp.Replace(pStrHTML, "&lt;$1SCRIPT$3&gt;")
	lObjRegExp.Pattern = "<(/)?(LINK|IFRAME|FRAMESET|FRAME|APPLET|OBJECT)([^>]*)>"
	pStrHTML = lObjRegExp.Replace(pStrHTML, "&lt;$1LINK$3&gt;")
	lObjRegExp.Pattern = "(<A[^>]+href\s?=\s?""?javascript:)[^""]*(""[^>]+>)"
	pStrHTML = lObjRegExp.Replace(pStrHTML, "$1//protected$2")
	lObjRegExp.Pattern = "(<IMG[^>]+src\s?=\s?""?javascript:)[^""]*(""[^>]+>)"
	pStrHTML = lObjRegExp.Replace(pStrHTML, "$1//protected$2")
	lObjRegExp.Pattern = "<([^>]*) on[^=\s]+\s?=\s?([^>]*)>"
	pStrHTML = lObjRegExp.Replace(pStrHTML, "<$1$3>")
	Set lObjRegExp = Nothing
	SafeHTML = pStrHTML
End Function
```

