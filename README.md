<div align="center">

## Convert an ADO Recordset to Text


</div>

### Description

This code snippet will show you how you can convert an ADO recordset to a delimited

text file in just a couple lines of code using the ADO GetString Method. You can easily export a recordset to a csv file using this method.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.7 (47 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/convert-an-ado-recordset-to-text__1-29947/archive/master.zip)





### Source Code

<br>
Dim rs As New ADODB.Recordset<br>
Dim fName As String, fNum As Integer<br>
<br>
    rs.Open "Select * from myTable", db, adOpenKeyset,
adLockReadOnly<br>
<br>
    fName = "C:\MyTestFile.csv"<br>
    fNum = FreeFile</p>
<p>    Open fName For Output As fNum<br>
<br>
    Do Until rs.EOF = True<br>
<br>
        Print #fNum, rs.GetString(adClipString, 1,
",", vbCr)<br>
<br>
    Loop<br>
<br>
    rsA.Close<br>
    Close #fNum</p>
<p>______________________________________________________________________</p>
<h1><a name="mdmthgetstringmethod(recordset)ado"></a>GetString Method</h1>
<p>Returns the <a href="mdobjodbrec.htm">Recordset</a> as a string.</p>
<h4>Syntax</h4>
<pre class="syntax"><i>Variant</i> = <i>recordset.</i><b>GetString</b><i>(<a
class="synParam" onclick="showTip(this)" href>StringFormat</a>, <a class="synParam"
onclick="showTip(this)" href>NumRows</a>, <a class="synParam" onclick="showTip(this)" href>ColumnDelimiter</a>, <a
class="synParam" onclick="showTip(this)" href>RowDelimiter</a>, <a class="synParam"
onclick="showTip(this)" href>NullExpr</a>)</i></pre>
<div class="reftip" id="reftip"
style="VISIBILITY: hidden; OVERFLOW: visible; POSITION: absolute"></div>
<h4>Return Value</h4>
<p>Returns the <b>Recordset</b> as a string-valued <b>Variant</b> (BSTR).</p>
<h4>Parameters</h4>
<dl>
 <dt><i>StringFormat</i> </dt>
 <dd>A <a href="mdcststringformatenum.htm">StringFormatEnum</a> value that specifies how the <b>Recordset</b>
 should be converted to a string. The <i>RowDelimiter</i>, <i>ColumnDelimiter</i>, and <i>NullExpr</i>
 parameters are used only with a <i>StringFormat</i> of <b>adClipString</b>. </dd>
 <dt><i>NumRows</i> </dt>
 <dd>Optional. The number of rows to be converted in the <b>Recordset</b>. If <i>NumRows </i>is
 not specified, or if it is greater than the total number of rows in the <b>Recordset</b>,
 then all the rows in the <b>Recordset</b> are converted. </dd>
 <dt><i>ColumnDelimiter</i> </dt>
 <dd>Optional. A delimiter used between columns, if specified, otherwise the TAB character. </dd>
 <dt><i>RowDelimiter</i> </dt>
 <dd>Optional. A delimiter used between rows, if specified, otherwise the CARRIAGE RETURN
 character. </dd>
 <dt><i>NullExpr</i> </dt>
 <dd>Optional. An expression used in place of a null value, if specified, otherwise the empty
 string. </dd>
</dl>
<h4>Remarks</h4>
<p>Row data, but no schema data, is saved to the string. Therefore, a <b>Recordset</b>
cannot be reopened using this string.</p>
<p>This method is equivalent to the RDO <b>GetClipString</b> method.</p>

