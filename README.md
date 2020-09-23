<div align="center">

## CRecordEdit v\. 2\.0


</div>

### Description

creates an HTML form given an ADO recordset. Uses different controls (textbox, checkbox, textarea) for different datatypes. Determines maxlength property based on datatype. Keeps some useful info in hidden fields.
 
### More Info
 
Add your own HTML and JavaScript to sections of the form:

.moreTblTags

.moreCaptionTags

.moreCellTags

.moreRowTags

Other:

.uniqueField - defaults to 0. change if needed.

.print(rs) - prints one record from the recordset.

Tested on all JET/ADO datatypes. Extensive SQL server testing coming soon. Should work though.

There are some hidden fields with no value that you can update with JS if you like (example 'selected field').

You should use cascading style sheets to format the form.

NS & IE event management in next version.

uses response.write to output nicely formatted HTML (with tabs and linefeeds).

Writes query string to a hidden field if you don't use a view.

Does not correctly determine maxlength for numeric types (except short int, and byte).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Neil McGuigan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/neil-mcguigan.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/neil-mcguigan-crecordedit-v-2-0__4-6745/archive/master.zip)

### API Declarations

Copyright (C) 2001, Neil McGuigan. All rights reserved. This software is licenced.


### Source Code

```
<%
'TO DO:
' cross-browser event handling (caption onclick, field on change)
'shouldn't use HTML tags for formatting, use CSS
'verify with all SQL Server types.
'add 'goto hyperlink' and 'sendmail' icons if field contains (only) a url
'fix maxlength. shows number of bits for numbers.
' end to do
class CRecordEdit
	public moreTblTags
	public moreRowTags
	public moreCaptionTags
	public moreCellTags
	private m_iPKFld
	public property let uniqueField(byval p)
		m_iPKFld=uniqueField
	end property
	private sub class_initialize()
		m_iPKFld=0
	end sub
	public sub print(byref rs)
		with response
			.write "<table"
			.write " " & moreTblTags
			.write ">" & vbCR
			dim fld
			for each fld in rs.fields
				.write vbTab & "<tr"
				.write " " & moreRowTags
				.write ">"
				.write "<th"
				select case fld.type
					case 3,17,2,131,5,6,4,130,129,202,200,72,7,135,203,201
						.write " onClick=""" & fld.name & ".focus();"""
					case 11 : .write " onClick=""" & fld.name & ".checked=!" & fld.name & ".checked;""" 'boolean
				end select
				.write " " & moreCaptionTags
				.write ">"
				.write fld.name
				.write "</th>"
				.write "<td"
'				.write " " & moreCellTags
				.write ">"
				call showControl(fld)
				.write "</td>"
				.write "</tr>" & vbCR
			next
			.write "</table>"
			.write vbCR & "<input type=""hidden"" name=""query"" value=""" & rs.source & """>"
			.write vbCR & "<input type=""hidden"" name=""dateAccessed"" value=""" & now() & """>"
			.write vbCR & "<input type=""hidden"" name=""uniqueField"" value=""" & rs.fields(m_iPKFld).name & """>"
			.write vbCR & "<input type=""hidden"" name=""changedFields"" value="""">"
			.write vbCR & "<input type=""hidden"" name=""selectedField"" value="""">"
			.write vbCR & "<input type=""hidden"" name=""selectedValue"" value="""">"
		end with
	end sub
	private sub showControl(byref fld)
		dim name,val,maxLength,width,ftype
		name=fld.name
		val=fld.value
		maxLength=fld.definedSize
		width=""
		ftype=fld.type
		'took out widths, use CSS
		select case ftype
			case 7,135 'dates
				maxLength=22
'				width=21
			case 3,4,5,6
				maxLength=99 'should figure this out actually
			case 2 'adSmallInt (-32,000)
				maxLength=7
'				width=7
			case 72 'GUID
				maxLength=38
'				width=43
			case 17 'byte
				maxlength=3
'				width=3
		end select
		select case ftype
			case 3,17,2,131,5,6,4,130,129,202,200,72,7,135 'regular text
				with response
					.write "<input"
					.write " type=""text"""
					.write " name=""" & name & """"
					.write " value=""" & val & """"
					if len(maxLength)>0 then .write " maxlength=""" & maxLength & """"
'					if len(width)>0 then .write " size=""" & width & """"
					.write " onFocus=""this.select();"""
					.write " " & moreCellTags
					.write ">"
				end with
			case 203,201 'memo
				with response
					.write "<textarea"
					.write " name=""" & name & """"
					.write " rows=""4"""
					.write " cols=""40"""
					.write " onFocus=""this.select();"""
					.write " " & moreCellTags
					.write ">"
					.write val
					.write "</textArea>"
				end with
			case 11 'boolean
				with response
					.write "<input"
					.write " type=""checkBox"""
					.write " name=""" & name & """"
					.write " value=""true"""
					if val then .write " checked "
					.write moreCellTags
					.write ">"
				end with
			case else
				response.write "&lt;binary&gt;"
		end select
				with response
					.write vbCR & vbTab & "<input type=""hidden"" name=""" & name & "UNDERLYING"" value=""" & val & """>"
					.write vbCR & vbTab & "<input type=""hidden"" name=""" & name & "ADOTYPE"" value=""" & ftype & """>"
				end with
	end sub
end class
%>
```

