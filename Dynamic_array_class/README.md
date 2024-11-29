# Dynamic array class in Classic ASP

## `dArray.class.asp`'s avaible Functions

- **Initialize class** -> `sub class_initialize()`
- **Terminate class** -> `sub class_terminate()`
- **Add an element** -> `Public Function add_element(element)`
- **Get element from index** -> `Public Function get_element(idx)`
- **Get array dimension** -> `Public Function get_dimension()`
- **Remove last element** -> `Public Function remove_last_element()`
- **Remove all element occurences** -> `Public Function remove_all_occurences(element)`
- **Remove first element occurence** -> `Public Function remove_first_occurence(element)`
- **Remove element from index** -> `Public Function remove_this_element(idx)`
- **Remove elements from indices** -> `Public Function remove_these_elements(indices_array)`
- **Reset** -> Re-initialize the object
- **Check if an element is present** -> ` Public Function contains(element)`
- **Retrieve the first index of an element** -> `Public Function get_first_index_occurence_of(element)`
- **Retrieve all indeces of an element (return an array)** -> `Public Function get_all_indeces_occurence_of(element)`
- **Write the entire array** -> ` Public Function write_array()`

## How to use: 

> From `Test.asp`

1. Initialize the class
  ```
  <%@LANGUAGE="VBSCRIPT"%>
  <!--#include file="dArray.class.asp"-->
  <%
      Dim da 
      Set da = new dArray
  ```
2. Use the class
  ```
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
  %>
  ```
