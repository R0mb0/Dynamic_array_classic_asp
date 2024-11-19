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
add_element_to_array("A")
add_element_to_array("B")
add_element_to_array("C")
add_element_to_array("D")
add_element_to_array("A")
Response.Write("Elements inside: ")
write_array()

Response.Write("--- Get element with index 1 --- <br>")

Response.Write("Element: ")
Response.Write(get_element_from_array(1) & "<br>")

Response.Write("--- Remove last element --- <br>")

remove_last_element_from_array()
Response.Write("Elements: ")
write_array()

Response.Write("--- Remove all occurence of B --- <br>")

remove_all_occurences_from_array("B")
Response.Write("Elements: ")
write_array()

Response.Write("--- Remove first occurence of C --- <br>")

remove_first_occurence_from_array("C")
Response.Write("Elements: ")
write_array()

Response.Write("--- Remove elements from posiction: 3 - 4 - 1 --- <br>")

Dim temp_array(2)
temp_array(0) = 3
temp_array(1) = 4
temp_array(2) = 1

remove_these_elements_from_array(temp_array)
Response.Write("Elements: ")
write_array()

Response.Write("--- Check if A is into array --- <br>")

Response.Write("Is A present? : ")
Response.Write(array_contains("A") & "<br>")

Response.Write("--- Check if C is into array --- <br>")

Response.Write("Is C present? : ")
Response.Write(array_contains("C") & "<br>")

Response.Write("--- Add other elements into array --- <br>")

add_element_to_array("B")
add_element_to_array("C")
add_element_to_array("D")
add_element_to_array("A")
Response.Write("Elements: ")
write_array()

Response.Write("--- Check first index of A --- <br>")

Response.Write("Index of A: ")
Response.Write(from_array_get_first_index_occurence_of("A") & "<br>")

Response.Write("--- Check indices of A --- <br>")

Response.Write("Indices of A: ")
Dim temp
For Each temp In from_array_get_all_indeces_occurence_of("A")
Response.Write(temp & " - ")
Next

Response.Write("<br>")
Response.Write("--- Remove element from posiction: 1 --- <br>")

remove_this_element_from_array(1)
Response.Write("Elements: ")
write_array()
%>