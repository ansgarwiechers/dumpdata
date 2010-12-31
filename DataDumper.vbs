'! Dictionary-based data structures are handy in various situations. However,
'! when the structures grow, they tend to become messy pretty fast. To deal
'! with this issue I wrote a function to inspect a given variable and return a
'! string representation of its data. It was inspired by Perl's Data::Dumper
'! module, although DumpData() is far less sophisticated (just in case you
'! were wondering).
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2010-12-29
'! @version 1.1

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'! Return a "printable" representation of the data stored in the given variable.
'! Primitive data types are displayed "as is". Structured data types (mainly
'! arrays and dictionaries) are expanded. Objects other than dictionaries are
'! represented by their respective type names without further introspection.
'!
'! @param  var    The variable to display.
'! @return A string representation of the data stored in the given variable.
'!
'! @raise  Unknown primitive data type (23)
Public Function DumpData(var)
	DumpData = RDump(var, 0)
End Function

'! Worker function to recursively generate the printable representation of
'! the given variable.
'!
'! @param  var    The variable to display.
'! @param  indent Level of indention.
'! @return A string representation of the data stored in the given variable.
'!
'! @raise  Unknown primitive data type (23)
Private Function RDump(var, indent)
	Dim data, i, key, hasDict, spacer

	data = ""

	If IsEmpty(var) Then
		data = data & "/Empty/"
	ElseIf IsNull(var) Then
		data = data & "/Null/"
	ElseIf IsDate(var) Then
		data = data & "#" & var & "#"
	ElseIf IsObject(var) Then
		If var Is Nothing Then
			data = "/Nothing/"
		Else
			Select Case TypeName(var)
			Case "Dictionary"
				If var.Count = 0 Then
					data = "{}"
				Else
					data = "{" & vbNewLine
					For Each key In var.Keys
						data = data & String(indent+1, vbTab) & RDump(key, indent) & " => " & RDump(var(key), indent+1) & vbNewLine
					Next
					data = data & String(indent, vbTab) & "}"
				End If
			Case "CArray"
				' handle my wrapper class for arrays
				' <http://www.planetcobalt.net/download/CArray-1.0.zip>
				RDump(var.ToArray, indent)
			Case Else
				data = data & "<" & TypeName(var) & ">"
			End Select
		End If
	Else
		Select Case TypeName(var)
		Case "Boolean"
			If var Then
				data = data & "True"
			Else
				data = data & "False"
			End If
		Case "Double"
			data = data & var
		Case "Integer"
			data = data & var
		Case "String"
			data = data & """" & var & """"
		Case "Variant()"
			If UBound(var) < 0 Then
				data = data & "[]"
			Else
				hasDict = False
				For i = 0 To UBound(var)
					If TypeName(var(i)) = "Dictionary" Then hasDict = True
				Next
				If hasDict Then
					spacer = vbNewLine & String(indent+1, vbTab)
				Else
					spacer = " "
				End If
				data = data & "[" & spacer & RDump(var(0), indent+1)
				For i = 1 To UBound(var)
					data = data & "," & spacer & RDump(var(i), indent+1)
				Next
				data = data & Replace(spacer, vbTab, "", 1, 1) & "]"
			End If
		Case Else
			Err.Raise 23, WScript.Script.Name, "Unknown primitive data type: " & TypeName(var)
		End Select
	End If

	RDump = data
End Function
