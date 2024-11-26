# Dynamic array library in Classic ASP

## `Dynamic_array.asp`'s avaible Functions

- **Initialize the array** -> `initialize_array()`
- **Check if the array is initializated** -> `is_array_initializated()`
- **Add an element into the array** -> `add_element_to_array(element)`
- **Get element from index** -> `get_element_from_array(idx)`
- **Remove last element from array** -> `remove_last_element_from_array()`
- **Remove an element from array** -> `remove_all_occurences_from_array(element)`
- **Remove the first element occurence from array** -> `remove_first_occurence_from_array(element)`
- **Remove element from index** -> `remove_this_elements_from_array(idx)`
- **Remove elements from indices** -> `remove_these_elements_from_array(indices_array)`
- **Reset the array** -> `initialize_array()`
- **Check if an element is in the array** -> `array_contains(element)`
- **Retrieve the first index of an element in the array** -> `from_array_get_first_index_occurence_of(element)`
- **Retrieve all indeces of an element inside the array (return an array)** -> `from_array_get_all_indeces_occurence_of(element)`
- **Retrieve the entire array** -> `get_array()`
- **Retrieve the array dimension** -> `get_array_dimension()`
- **Write the entire array** -> `write_array()`

## How to use: 

> From `Test1.asp`

1. Initialize the array and check it's status
   ```
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="Dynamic_array.asp"-->
   <%
   initialize_array()
   Response.Write("Array status: ")
   Response.Write(is_array_initializated() & "<br>")
   ```
2. Use the functions to manage the array
   ```
   add_element_to_array("A")
   add_element_to_array("B")
   add_element_to_array("C")
   add_element_to_array("D")
   Response.Write("Elements inside: ")
   write_array()
   %>
   ```
