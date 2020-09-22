<div align="center">

## Remove HTML \+ Optional Ingnore Tags


</div>

### Description

This function will strip a string of all html. An optional parameter (sIgnoreTags) allows specified HTML tags to be ignored from stripping.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Garske](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-garske.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-garske-remove-html-optional-ingnore-tags__1-37981/archive/master.zip)





### Source Code

```
Function RemoveHTML(ByVal sHTML, ByVal sIgnoreTags)
	Dim I, J, arr_sIgnoreTags, sIgnoreTag, bIgnoreTags, bIgnoreTag, iIndex
	bIgnoreTags = False
	If Len(sIgnoreTags) > 0 Then
		arr_sIgnoreTags = Split(sIgnoreTags, ",")
		bIgnoreTags = True
	End If
	sHTML = Trim(sHTML)
	If IsNull(sHTML) Then sHTML = ""
	sHTML = Replace(sHTML, vbCrLf, "") 'Makes easier
	I = InStr(1, sHTML, "<")
	Do While I <> 0
		bIgnoreTag = False
		If bIgnoreTags Then
			For iIndex = 0 To UBound(arr_sIgnoreTags)
				sIgnoreTag = Trim(arr_sIgnoreTags(iIndex))
				If UCase(Mid(sHTML, I + 1, Len(sIgnoreTag))) = UCase(sIgnoreTag) Then
					bIgnoreTag = True
					Exit For
				End If
			Next
		End If
		If Not bIgnoreTag Then
			J = InStr(I + 1, sHTML, ">")
			If J <> 0 Then
				sHTML = Left(sHTML, I - 1) & Mid(sHTML, J + 1)
			Else
				'Chop off rest off sHTML since bad HTML
				sHTML = Left(sHTML, I - 1)
			End If
		Else
			I = I + 1 'So next tag is searched
		End If
		I = InStr(I, sHTML, "<")
	Loop
	If Len(sHTML) = 0 Then sHTML = "&nbsp;"
	RemoveHTML = sHTML
End function
```

