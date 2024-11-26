<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Dynamic_arrays.asp"-->
<%

Response.Write("--- Starting test --- <br>")

Response.Write("--- Initialize test array --- <br>")

Dim test_array
test_array = Array()
test_array = get_initializated_dynamic_array()

Response.Write("--- Add elements into the test array and ckeck status --- <br>")

add_element_to_dynamic_array test_array,"A"
add_element_to_dynamic_array test_array,"B"
add_element_to_dynamic_array test_array,"C"
add_element_to_dynamic_array test_array,"D"
add_element_to_dynamic_array test_array,"A"
add_element_to_dynamic_array test_array,"B"
add_element_to_dynamic_array test_array,"C"
add_element_to_dynamic_array test_array,"D"
add_element_to_dynamic_array test_array,"A"
Response.Write("Elements inside: ")
write_dynamic_array(test_array)

Response.Write("--- Get element with index 1 --- <br>")

Response.Write("Element: ")
Response.Write(get_element_from_dynamic_array(test_array,1) & "<br>")

Response.Write("--- Remove last element --- <br>")

remove_last_element_from_dynamic_array test_array
Response.Write("Elements: ")
write_dynamic_array(test_array)

Response.Write("--- Remove all occurence of B --- <br>")
remove_all_occurences_from_dynamic_array test_array,"B"
Response.Write("Elements: ")
write_dynamic_array(test_array)

Response.Write("--- Remove first occurence of C --- <br>")

remove_first_occurence_from_dynamic_array test_array,"C"
Response.Write("Elements: ")
write_dynamic_array(test_array)

Response.Write("--- Remove elements from posiction: 3 - 4 - 1 --- <br>")

Dim temp_array(2)
temp_array(0) = 3
temp_array(1) = 4
temp_array(2) = 1

remove_these_elements_from_dynamic_array test_array,temp_array
Response.Write("Elements: ")
write_dynamic_array(test_array)

Response.Write("--- Check if A is into array --- <br>")

Response.Write("Is A present? : ")
Response.Write(dynamic_array_contains(test_array,"A") & "<br>")

Response.Write("--- Check if C is into array --- <br>")

Response.Write("Is C present? : ")
Response.Write(dynamic_array_contains(test_array,"C") & "<br>")

Response.Write("--- Add other elements into array --- <br>")

add_element_to_dynamic_array test_array,"B"
add_element_to_dynamic_array test_array,"C"
add_element_to_dynamic_array test_array,"D"
add_element_to_dynamic_array test_array,"A"
Response.Write("Elements: ")
write_dynamic_array(test_array)

Response.Write("--- Check first index of A --- <br>")

Response.Write("Index of A: ")
Response.Write(from_dynamic_array_get_first_index_occurence_of(test_array,"A") & "<br>")

Response.Write("--- Check indices of A --- <br>")

Response.Write("Indices of A: ")
Dim temp
For Each temp In from_dynamic_array_get_all_indeces_occurence_of(test_array,"A")
Response.Write(temp & " - ")
Next

Response.Write("<br>")
Response.Write("--- Remove element from posiction: 1 --- <br>")

remove_this_element_from_dynamic_array test_array,1
Response.Write("Elements: ")
write_dynamic_array(test_array)
%>
