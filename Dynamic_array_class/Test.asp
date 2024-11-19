<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="dArray.class.asp"-->
<%
Response.Write("--- Starting test --- <br>")

Response.Write("--- Initialize and test array status --- <br>")

Dim da 
Set da = new dArray 
Response.Write("Array status: ")
Response.Write(da.is_initializated() & "<br>")

Response.Write("--- Add elements into array and ckeck status --- <br>")

da.add_element("A")
da.add_element("B")
da.add_element("C")
da.add_element("D")
da.add_element("A")
da.add_element("B")
da.add_element("C")
da.add_element("D")
da.add_element("A")
Response.Write("Elements inside: ")
da.write_array()

Response.Write("--- Get element with index 1 --- <br>")

Response.Write("Element: ")
Response.Write(da.get_element(1) & "<br>")

Response.Write("--- Remove last element --- <br>")

da.remove_last_element()
Response.Write("Elements: ")
da.write_array()

Response.Write("--- Remove all occurence of B --- <br>")

da.remove_all_occurences("B")
Response.Write("Elements: ")
da.write_array()

Response.Write("--- Remove first occurence of C --- <br>")

da.remove_first_occurence("C")
Response.Write("Elements: ")
da.write_array()

Response.Write("--- Remove elements from posiction: 3 - 4 - 1 --- <br>")

Dim temp_array(2)
temp_array(0) = 3
temp_array(1) = 4
temp_array(2) = 1

da.remove_these_elements(temp_array)
Response.Write("Elements: ")
da.write_array()

Response.Write("--- Check if A is into array --- <br>")

Response.Write("Is A present? : ")
Response.Write(da.contains("A") & "<br>")

Response.Write("--- Check if C is into array --- <br>")

Response.Write("Is C present? : ")
Response.Write(da.contains("C") & "<br>")

Response.Write("--- Add other elements into array --- <br>")

da.add_element("B")
da.add_element("C")
da.add_element("D")
da.add_element("A")
Response.Write("Elements: ")
da.write_array()

Response.Write("--- Check first index of A --- <br>")

Response.Write("Index of A: ")
Response.Write(da.get_first_index_occurence_of("A") & "<br>")

Response.Write("--- Check indices of A --- <br>")

Response.Write("Indices of A: ")
Dim temp
For Each temp In da.get_all_indeces_occurence_of("A")
    Response.Write(temp & " - ")
Next

Response.Write("<br>")
Response.Write("--- Remove element from posiction: 1 --- <br>")

da.remove_this_element(1)
Response.Write("Elements: ")
da.write_array()
%>