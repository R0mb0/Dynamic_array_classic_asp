<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Dynamic_array.asp"-->
<%
Response.Write("--- Starting test --- <br>")

Response.Write("--- Initialize and test array status --- <br>")

initialize_array()
Response.Write("Array status: ")
Response.Write(is_array_initializated() & "<br>")

Response.Write("--- Add elements into array and ckeck status --- <br>")

add_element_to_array("A")
add_element_to_array("B")
add_element_to_array("C")
add_element_to_array("D")
Response.Write("Elements inside: ")
write_array()

Response.Write("--- Get element with 1 --- <br>")

Response.Write("Element: ")
Response.Write(get_element_from_array(1) & "<br>")

Response.Write("--- Remove last element --- <br>")

remove_last_element_from_array()
Response.Write("Elements: ")
write_array()

Response.Write("--- Remove B from array --- <br>")

remove_element_from_array("B")
Response.Write("Elements: ")
write_array()

Response.Write("--- Check if A is into array --- <br>")

Response.Write("Is A present? : ")
Response.Write(array_contains("A") & "<br>")

Response.Write("--- Check index of C --- <br>")

Response.Write("Index of C: ")
Response.Write(from_array_get_index_of("C"))
%>